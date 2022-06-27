VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{63BEADB1-20E1-478A-9B40-DDDAFBF3624F}#1.0#0"; "bsGradientLabel.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Object = "{FF19AA0C-2968-41B8-A906-E80997A9C394}#208.0#0"; "WSAFIPFEOCX.ocx"
Object = "{706C3604-A82B-4400-9EE4-3433F1D8DB08}#1.8#0"; "EpsonFPHostControlX.ocx"
Begin VB.Form frmRemito 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos de Ventas. 240418"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   14865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   14865
   Begin MSDataGridLib.DataGrid dgClientes 
      Height          =   1545
      Left            =   1170
      TabIndex        =   229
      Top             =   1620
      Visible         =   0   'False
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   2725
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      ForeColor       =   255
      HeadLines       =   1
      RowHeight       =   15
      TabAction       =   1
      WrapCellPointer =   -1  'True
      RowDividerStyle =   4
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
            LCID            =   3082
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
            LCID            =   3082
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
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1875
      Left            =   10590
      TabIndex        =   213
      Top             =   3870
      Visible         =   0   'False
      Width           =   3375
      Begin VB.CommandButton cmdEstado 
         Caption         =   "Estado"
         Height          =   315
         Left            =   1470
         TabIndex        =   220
         Top             =   1470
         Width           =   1815
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "TCF"
         Height          =   585
         Left            =   270
         TabIndex        =   219
         Top             =   1170
         Width           =   825
      End
      Begin VB.CommandButton cmdCommand1 
         Caption         =   "Conecta"
         Height          =   315
         Left            =   1470
         TabIndex        =   216
         Top             =   1170
         Width           =   1815
      End
      Begin VB.TextBox vvelocidad 
         Height          =   405
         Left            =   1590
         TabIndex        =   215
         Text            =   "Text1"
         Top             =   720
         Width           =   1485
      End
      Begin VB.TextBox vpuerto 
         Height          =   375
         Left            =   1590
         TabIndex        =   214
         Text            =   "Text1"
         Top             =   240
         Width           =   1485
      End
      Begin VB.Label lblVelocidad 
         Caption         =   "Velocidad"
         Height          =   285
         Left            =   360
         TabIndex        =   218
         Top             =   780
         Width           =   825
      End
      Begin VB.Label lblPuerto 
         Caption         =   "Puerto"
         Height          =   285
         Left            =   570
         TabIndex        =   217
         Top             =   420
         Width           =   495
      End
   End
   Begin EpsonFPHostControlX.EpsonFPHostControl EpsonFP 
      Left            =   6930
      OleObjectBlob   =   "frmRemito.frx":0000
      Top             =   7380
   End
   Begin XtremeSuiteControls.GroupBox GroupBox6 
      Height          =   435
      Left            =   90
      TabIndex        =   202
      Top             =   690
      Width           =   10065
      _Version        =   851968
      _ExtentX        =   17754
      _ExtentY        =   767
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   375
         Left            =   3210
         TabIndex        =   206
         Top             =   30
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "F6"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.FlatEdit vcodEmpresa 
         Height          =   315
         Left            =   1485
         TabIndex        =   203
         Top             =   45
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   -2147483633
         Appearance      =   3
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit vdescEmpresa 
         Height          =   315
         Left            =   3600
         TabIndex        =   204
         Top             =   90
         Width           =   6375
         _Version        =   851968
         _ExtentX        =   11245
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   4210752
         BackColor       =   -2147483633
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   285
         Left            =   120
         TabIndex        =   205
         Top             =   90
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Empresa:"
         ForeColor       =   255
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.ListBox log 
      Height          =   2205
      Left            =   315
      TabIndex        =   200
      Top             =   3930
      Visible         =   0   'False
      Width           =   14115
   End
   Begin WSAFIPFEOCX.WSAFIPFEx fe 
      Left            =   13140
      Top             =   7335
      _ExtentX        =   2143
      _ExtentY        =   820
   End
   Begin MSDataGridLib.DataGrid dgArticulos 
      Height          =   2580
      Left            =   90
      TabIndex        =   7
      Tag             =   "0"
      Top             =   3645
      Visible         =   0   'False
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   4551
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   255
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
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
   Begin XtremeSuiteControls.GroupBox GroupBox4 
      Height          =   195
      Left            =   30
      TabIndex        =   188
      Top             =   540
      Width           =   14775
      _Version        =   851968
      _ExtentX        =   26061
      _ExtentY        =   344
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.TabControl TabTotales 
      Height          =   1845
      Left            =   90
      TabIndex        =   128
      Top             =   6990
      Width           =   14835
      _Version        =   851968
      _ExtentX        =   26167
      _ExtentY        =   3254
      _StockProps     =   68
      ItemCount       =   6
      Item(0).Caption =   "Totales"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "fraTotales"
      Item(1).Caption =   "Comentarios"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "txtObservaciones"
      Item(2).Caption =   "Comentarios Fiscales"
      Item(2).ControlCount=   1
      Item(2).Control(0)=   "cmdComentarios"
      Item(3).Caption =   "Cant. de volquetes"
      Item(3).ControlCount=   2
      Item(3).Control(0)=   "Label5"
      Item(3).Control(1)=   "txt_vcantidadVolquete"
      Item(4).Caption =   "Choferes"
      Item(4).ControlCount=   7
      Item(4).Control(0)=   "Label1(0)"
      Item(4).Control(1)=   "btn_CargarChoferes(5)"
      Item(4).Control(2)=   "grd_Choferes"
      Item(4).Control(3)=   "btn_pasar"
      Item(4).Control(4)=   "vchofer"
      Item(4).Control(5)=   "PushButton1"
      Item(4).Control(6)=   "PushButton2"
      Item(5).Caption =   "Datos del Remito"
      Item(5).ControlCount=   4
      Item(5).Control(0)=   "GroupBox1"
      Item(5).Control(1)=   "GroupBox2"
      Item(5).Control(2)=   "GroupBox3"
      Item(5).Control(3)=   "GroupBox7"
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   555
         Left            =   -69940
         TabIndex        =   168
         Top             =   330
         Visible         =   0   'False
         Width           =   6525
         _Version        =   851968
         _ExtentX        =   11509
         _ExtentY        =   979
         _StockProps     =   79
         Caption         =   "Datos del receptor:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vRemitoRecibio 
            Height          =   315
            Left            =   1380
            TabIndex        =   177
            Top             =   210
            Width           =   5055
            _Version        =   851968
            _ExtentX        =   8916
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label Label11 
            Height          =   285
            Left            =   300
            TabIndex        =   178
            Top             =   210
            Width           =   1035
            _Version        =   851968
            _ExtentX        =   1826
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Recibió:"
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   315
         Left            =   -56470
         TabIndex        =   158
         Top             =   1050
         Visible         =   0   'False
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Vaciar "
         Appearance      =   5
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   -56470
         TabIndex        =   156
         Top             =   630
         Visible         =   0   'False
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Verlo como lista"
         Appearance      =   4
      End
      Begin VB.TextBox vchofer 
         Height          =   315
         Left            =   -68410
         TabIndex        =   155
         Top             =   870
         Visible         =   0   'False
         Width           =   4605
      End
      Begin XtremeSuiteControls.PushButton btn_pasar 
         Height          =   315
         Left            =   -63700
         TabIndex        =   154
         Top             =   870
         Visible         =   0   'False
         Width           =   435
         _Version        =   851968
         _ExtentX        =   767
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   ">"
         Appearance      =   3
      End
      Begin XtremeSuiteControls.FlatEdit txt_vcantidadVolquete 
         Height          =   345
         Left            =   -65950
         TabIndex        =   150
         Top             =   780
         Visible         =   0   'False
         Width           =   2295
         _Version        =   851968
         _ExtentX        =   4048
         _ExtentY        =   609
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Frame fraTotales 
         Caption         =   "Totales :"
         ForeColor       =   &H00808080&
         Height          =   1305
         Left            =   60
         TabIndex        =   129
         Top             =   390
         Width           =   14685
         Begin XtremeSuiteControls.PushButton PushButton5 
            Height          =   345
            Left            =   7590
            TabIndex        =   201
            Top             =   450
            Width           =   705
            _Version        =   851968
            _ExtentX        =   1244
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Log"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.TextBox txtDescuento 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   6300
            TabIndex        =   139
            Top             =   345
            Width           =   850
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   12990
            TabIndex        =   138
            Top             =   525
            Width           =   1575
         End
         Begin VB.TextBox txtIva 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   1
            Left            =   9720
            TabIndex        =   137
            Top             =   540
            Width           =   1575
         End
         Begin VB.TextBox txtSubtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1620
            TabIndex        =   136
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox txtPDescuento 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5550
            TabIndex        =   135
            Top             =   345
            Width           =   675
         End
         Begin VB.TextBox txtIva 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000080&
            Height          =   285
            Index           =   0
            Left            =   9720
            TabIndex        =   134
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtImpuesto 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   5550
            TabIndex        =   133
            Top             =   645
            Width           =   1605
         End
         Begin VB.CommandButton cmdActualizarTotal 
            Caption         =   "Actualizar Total"
            Height          =   315
            Left            =   180
            TabIndex        =   132
            Top             =   750
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.CheckBox chkTotalManual 
            Caption         =   "Total Manual"
            Height          =   195
            Left            =   210
            TabIndex        =   131
            Top             =   900
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.TextBox txtIva 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H00800000&
            Height          =   285
            Index           =   2
            Left            =   9720
            TabIndex        =   130
            Top             =   840
            Width           =   1575
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "> Descuento %  :"
            Height          =   195
            Index           =   12
            Left            =   4140
            TabIndex        =   146
            Top             =   375
            Width           =   1395
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "> I.V.A. 21 %:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   8310
            TabIndex        =   145
            Top             =   570
            Width           =   1395
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "> Subtotal :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   144
            Top             =   495
            Width           =   1395
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "> Impuesto  :"
            Height          =   195
            Index           =   13
            Left            =   4140
            TabIndex        =   143
            Top             =   675
            Width           =   1395
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "> Total  :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   11640
            TabIndex        =   142
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "> I.V.A. 10,5 %:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   8310
            TabIndex        =   141
            Top             =   270
            Width           =   1395
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            Caption         =   "> I.V.A. 27 %:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   8310
            TabIndex        =   140
            Top             =   870
            Width           =   1395
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtObservaciones 
         Height          =   1125
         Left            =   -69880
         TabIndex        =   147
         Top             =   450
         Visible         =   0   'False
         Width           =   14625
         _Version        =   851968
         _ExtentX        =   25797
         _ExtentY        =   1984
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   255
      End
      Begin XtremeSuiteControls.PushButton cmdComentarios 
         Height          =   255
         Left            =   -69730
         TabIndex        =   148
         Top             =   810
         Visible         =   0   'False
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Comentarios Pre-Cargados"
         Appearance      =   1
      End
      Begin XtremeSuiteControls.PushButton btn_CargarChoferes 
         Height          =   255
         Index           =   5
         Left            =   -69100
         TabIndex        =   152
         Tag             =   "Vendedor"
         Top             =   900
         Visible         =   0   'False
         Width           =   375
         _Version        =   851968
         _ExtentX        =   661
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "..."
         Appearance      =   3
      End
      Begin Grid.KlexGrid grd_Choferes 
         Height          =   1335
         Left            =   -63100
         TabIndex        =   153
         Top             =   390
         Visible         =   0   'False
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   2355
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
         MouseIcon       =   "frmRemito.frx":009B
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   915
         Left            =   -63280
         TabIndex        =   169
         Top             =   420
         Visible         =   0   'False
         Width           =   7995
         _Version        =   851968
         _ExtentX        =   14102
         _ExtentY        =   1614
         _StockProps     =   79
         Caption         =   "Datos del transportista:"
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vTransportistaCuit 
            Height          =   315
            Left            =   4920
            TabIndex        =   175
            Top             =   540
            Width           =   2955
            _Version        =   851968
            _ExtentX        =   5212
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vTransportistaDomicilio 
            Height          =   315
            Left            =   930
            TabIndex        =   173
            Top             =   600
            Width           =   2985
            _Version        =   851968
            _ExtentX        =   5265
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit vTransportistaNombre 
            Height          =   315
            Left            =   930
            TabIndex        =   171
            Top             =   240
            Width           =   2985
            _Version        =   851968
            _ExtentX        =   5265
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton vmdBuscarChofer 
            Height          =   285
            Left            =   4020
            TabIndex        =   176
            Top             =   210
            Width           =   375
            _Version        =   851968
            _ExtentX        =   661
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label vCuitTransportista 
            Height          =   285
            Left            =   4110
            TabIndex        =   174
            Top             =   570
            Width           =   765
            _Version        =   851968
            _ExtentX        =   1349
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "CUIT:"
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label13 
            Height          =   285
            Left            =   90
            TabIndex        =   172
            Top             =   600
            Width           =   765
            _Version        =   851968
            _ExtentX        =   1349
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Domicilio:"
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label Label12 
            Height          =   285
            Left            =   60
            TabIndex        =   170
            Top             =   240
            Width           =   765
            _Version        =   851968
            _ExtentX        =   1349
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Nombre:"
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   495
         Left            =   -69940
         TabIndex        =   179
         Top             =   840
         Visible         =   0   'False
         Width           =   6525
         _Version        =   851968
         _ExtentX        =   11509
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.FlatEdit vlentrega 
            Height          =   315
            Left            =   1350
            TabIndex        =   180
            Top             =   150
            Width           =   5055
            _Version        =   851968
            _ExtentX        =   8916
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label Label14 
            Height          =   285
            Left            =   90
            TabIndex        =   181
            Top             =   120
            Width           =   1125
            _Version        =   851968
            _ExtentX        =   1984
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Lugar entrega: "
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox7 
         Height          =   495
         Left            =   -69910
         TabIndex        =   224
         Top             =   1290
         Visible         =   0   'False
         Width           =   14655
         _Version        =   851968
         _ExtentX        =   25850
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin VB.CommandButton cmdAct 
            Caption         =   "Act."
            Height          =   285
            Left            =   13950
            TabIndex        =   227
            Top             =   150
            Width           =   615
         End
         Begin XtremeSuiteControls.FlatEdit vleyenda 
            Height          =   315
            Left            =   1380
            TabIndex        =   225
            Top             =   120
            Width           =   12375
            _Version        =   851968
            _ExtentX        =   21828
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.Label Label18 
            Height          =   285
            Left            =   90
            TabIndex        =   226
            Top             =   120
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Leyenda Factura:"
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Choferes:"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   -69910
         TabIndex        =   151
         Top             =   930
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label Label5 
         Caption         =   "Cantidad de volquete utilizados en este documento: "
         Height          =   315
         Left            =   -69820
         TabIndex        =   149
         Top             =   840
         Visible         =   0   'False
         Width           =   3885
      End
   End
   Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
      Height          =   285
      Left            =   11970
      TabIndex        =   85
      Top             =   840
      Width           =   2565
      _ExtentX        =   4524
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
   Begin XtremeSuiteControls.GroupBox GBOtrosDocumentos 
      Height          =   4065
      Left            =   15030
      TabIndex        =   58
      Top             =   2520
      Visible         =   0   'False
      Width           =   14835
      _Version        =   851968
      _ExtentX        =   26167
      _ExtentY        =   7170
      _StockProps     =   79
      Caption         =   " Otros Documentos"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.GroupBox GBCaja 
         Height          =   3135
         Left            =   9780
         TabIndex        =   86
         Top             =   300
         Visible         =   0   'False
         Width           =   4815
         _Version        =   851968
         _ExtentX        =   8493
         _ExtentY        =   5530
         _StockProps     =   79
         Caption         =   "Ingreso de Pago de Contado:"
         ForeColor       =   14737632
         BackColor       =   8421504
         Appearance      =   6
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   30
            Picture         =   "frmRemito.frx":00B7
            ScaleHeight     =   555
            ScaleWidth      =   4785
            TabIndex        =   87
            TabStop         =   0   'False
            Top             =   2520
            Width           =   4785
            Begin XtremeSuiteControls.PushButton cmdCerrarPago 
               Height          =   375
               Left            =   3580
               TabIndex        =   88
               Top             =   90
               Width           =   1155
               _Version        =   851968
               _ExtentX        =   2028
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Cerrar"
               Appearance      =   6
               Picture         =   "frmRemito.frx":516A
            End
            Begin XtremeSuiteControls.PushButton cmdGuardarPago 
               Height          =   375
               Left            =   2450
               TabIndex        =   89
               Top             =   90
               Width           =   1155
               _Version        =   851968
               _ExtentX        =   2028
               _ExtentY        =   661
               _StockProps     =   79
               Caption         =   "Guardar"
               Appearance      =   6
               Picture         =   "frmRemito.frx":556A
               BorderGap       =   10
            End
            Begin VB.Label lblWGESTION2010 
               AutoSize        =   -1  'True
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
               Index           =   3
               Left            =   75
               TabIndex        =   91
               Top             =   170
               Width           =   1770
            End
            Begin VB.Label lblWGESTION2010 
               AutoSize        =   -1  'True
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
               Index           =   2
               Left            =   50
               TabIndex        =   90
               Top             =   150
               Width           =   1770
            End
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   315
            Index           =   0
            Left            =   1440
            TabIndex        =   92
            Top             =   240
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   2
            Left            =   1995
            TabIndex        =   93
            Tag             =   "CajaBanco"
            Top             =   240
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   315
            Index           =   1
            Left            =   2400
            TabIndex        =   94
            Top             =   240
            Width           =   2295
            _Version        =   851968
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   315
            Index           =   2
            Left            =   1440
            TabIndex        =   95
            Top             =   600
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   3
            Left            =   1995
            TabIndex        =   96
            Tag             =   "BancoCuenta"
            Top             =   600
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   315
            Index           =   3
            Left            =   2400
            TabIndex        =   97
            Top             =   600
            Width           =   2295
            _Version        =   851968
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   315
            Index           =   6
            Left            =   1440
            TabIndex        =   98
            Top             =   1320
            Width           =   1695
            _Version        =   851968
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   435
            Index           =   8
            Left            =   1440
            TabIndex        =   99
            Top             =   2040
            Width           =   3255
            _Version        =   851968
            _ExtentX        =   5741
            _ExtentY        =   767
            _StockProps     =   77
            BackColor       =   -2147483643
            MaxLength       =   255
            ScrollBars      =   2
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   315
            Index           =   4
            Left            =   1440
            TabIndex        =   101
            Top             =   960
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   4
            Left            =   1995
            TabIndex        =   102
            Tag             =   "TipoValor"
            Top             =   960
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   315
            Index           =   5
            Left            =   2400
            TabIndex        =   103
            Top             =   960
            Width           =   2295
            _Version        =   851968
            _ExtentX        =   4048
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtCaja 
            Height          =   315
            Index           =   7
            Left            =   1440
            TabIndex        =   100
            Top             =   1680
            Width           =   1695
            _Version        =   851968
            _ExtentX        =   2990
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.Label lblOtrosDocumentos 
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   109
            Top             =   280
            Width           =   1500
            _Version        =   851968
            _ExtentX        =   2646
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Caja/Banco:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblOtrosDocumentos 
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   108
            Top             =   640
            Width           =   1500
            _Version        =   851968
            _ExtentX        =   2646
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Cuenta Banco:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblOtrosDocumentos 
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   107
            Top             =   1000
            Width           =   1500
            _Version        =   851968
            _ExtentX        =   2646
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Tipo de Valor:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblOtrosDocumentos 
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   106
            Top             =   1720
            Width           =   1500
            _Version        =   851968
            _ExtentX        =   2646
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Importe:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblOtrosDocumentos 
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   105
            Top             =   2040
            Width           =   1500
            _Version        =   851968
            _ExtentX        =   2646
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Observaciones:"
            Transparent     =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblOtrosDocumentos 
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   104
            Top             =   1360
            Width           =   1500
            _Version        =   851968
            _ExtentX        =   2646
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "Numero de Valor:"
            Transparent     =   -1  'True
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   5
         Left            =   6480
         TabIndex        =   68
         Top             =   840
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoMovimiento 
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   82
         Top             =   480
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox cboBienesServicios 
         Height          =   315
         Left            =   6480
         TabIndex        =   75
         Top             =   1560
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         DropDownItemCount=   2
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   555
         Index           =   7
         Left            =   2040
         TabIndex        =   78
         Top             =   2280
         Width           =   7215
         _Version        =   851968
         _ExtentX        =   12726
         _ExtentY        =   979
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   0
         Left            =   2040
         TabIndex        =   69
         Top             =   840
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   1
         Left            =   2040
         TabIndex        =   70
         Top             =   1200
         Width           =   735
         _Version        =   851968
         _ExtentX        =   1296
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   2
         Left            =   2880
         TabIndex        =   71
         Top             =   1200
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   3
         Left            =   2040
         TabIndex        =   62
         Top             =   1560
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   4
         Left            =   2040
         TabIndex        =   65
         Top             =   1920
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   6
         Left            =   6480
         TabIndex        =   73
         Top             =   1200
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtIB 
         Height          =   315
         Index           =   10
         Left            =   6600
         TabIndex        =   77
         Top             =   2880
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
         MaxLength       =   254
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   2800
         TabIndex        =   83
         Tag             =   "TipoMovimientos"
         Top             =   480
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtTipoMovimiento 
         Height          =   315
         Index           =   1
         Left            =   3160
         TabIndex        =   84
         Top             =   480
         Width           =   1535
         _Version        =   851968
         _ExtentX        =   2708
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.Label lblTipoDocumento 
         Height          =   255
         Left            =   6480
         TabIndex        =   80
         Top             =   1920
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   450
         _StockProps     =   79
         ForeColor       =   -2147483635
         BackColor       =   65535
         Alignment       =   2
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   76
         Top             =   2320
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Leyenda :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   7
         Left            =   4800
         TabIndex        =   74
         Top             =   1600
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Bienes/Servicios :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   6
         Left            =   4800
         TabIndex        =   72
         Top             =   1240
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Impuesto Exento :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   5
         Left            =   4800
         TabIndex        =   67
         Top             =   885
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe NO Gravado :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   10
         Left            =   4800
         TabIndex        =   66
         Top             =   2920
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Total :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   64
         Top             =   1960
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Retenciones :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   63
         Top             =   1600
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Deducciones :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   61
         Top             =   1240
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "IVA :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   60
         Top             =   880
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Importe Gravado :"
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblOtrosDocumentos 
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   520
         Width           =   1750
         _Version        =   851968
         _ExtentX        =   3087
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Tipo Comprobantes:"
         Transparent     =   -1  'True
      End
   End
   Begin VB.Frame fraDetalle 
      BorderStyle     =   0  'None
      Height          =   3870
      Left            =   90
      TabIndex        =   40
      Top             =   3210
      Width           =   14805
      Begin XtremeSuiteControls.FlatEdit vlineasDetalles 
         Height          =   255
         Left            =   4680
         TabIndex        =   183
         Top             =   90
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   2
         UseVisualStyle  =   0   'False
      End
      Begin Grid.KlexGrid KlexDetalle 
         Height          =   2835
         Left            =   0
         TabIndex        =   49
         Top             =   420
         Width           =   14775
         _ExtentX        =   26061
         _ExtentY        =   5001
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
         MouseIcon       =   "frmRemito.frx":594E
      End
      Begin VB.Frame fraCargaDetalle 
         BorderStyle     =   0  'None
         Height          =   975
         Left            =   0
         TabIndex        =   50
         Top             =   3150
         Width           =   14805
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "B"
            Height          =   195
            Left            =   9945
            TabIndex        =   222
            Top             =   360
            Width           =   255
         End
         Begin XtremeSuiteControls.CheckBox chkfijo 
            Height          =   165
            Left            =   780
            TabIndex        =   163
            Top             =   135
            Width           =   825
            _Version        =   851968
            _ExtentX        =   1455
            _ExtentY        =   291
            _StockProps     =   79
            Caption         =   "Fijo F12"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtDetalle 
            Height          =   375
            Index           =   1
            Left            =   1635
            TabIndex        =   52
            Top             =   150
            Width           =   8295
            _Version        =   851968
            _ExtentX        =   14631
            _ExtentY        =   661
            _StockProps     =   77
            ForeColor       =   4210752
            BackColor       =   16777215
            BackColor       =   16777215
         End
         Begin XtremeSuiteControls.FlatEdit txtDetalle 
            Height          =   345
            Index           =   0
            Left            =   15
            TabIndex        =   51
            Top             =   135
            Width           =   705
            _Version        =   851968
            _ExtentX        =   1244
            _ExtentY        =   609
            _StockProps     =   77
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777215
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtDetalle 
            Height          =   375
            Index           =   2
            Left            =   10200
            TabIndex        =   53
            Top             =   120
            Width           =   930
            _Version        =   851968
            _ExtentX        =   1640
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   16777215
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   16777215
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtDetalle 
            Height          =   375
            Index           =   3
            Left            =   11130
            TabIndex        =   54
            Top             =   120
            Width           =   840
            _Version        =   851968
            _ExtentX        =   1482
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   16777215
            BackColor       =   16777215
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtDetalle 
            Height          =   375
            Index           =   4
            Left            =   11970
            TabIndex        =   55
            Top             =   120
            Width           =   840
            _Version        =   851968
            _ExtentX        =   1482
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtDetalle 
            Height          =   375
            Index           =   5
            Left            =   12810
            TabIndex        =   56
            Top             =   120
            Width           =   840
            _Version        =   851968
            _ExtentX        =   1482
            _ExtentY        =   661
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   1
         End
         Begin XtremeSuiteControls.FlatEdit txtDetalle 
            Height          =   375
            Index           =   6
            Left            =   13680
            TabIndex        =   57
            Top             =   120
            Width           =   1080
            _Version        =   851968
            _ExtentX        =   1905
            _ExtentY        =   661
            _StockProps     =   77
            ForeColor       =   0
            BackColor       =   16777215
            BackColor       =   16777215
            Alignment       =   1
         End
         Begin XtremeSuiteControls.CheckBox chkIncCod 
            Height          =   255
            Left            =   780
            TabIndex        =   184
            Tag             =   "IncluyeCodEnDoc"
            Top             =   315
            Width           =   900
            _Version        =   851968
            _ExtentX        =   1587
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Inc.Cod."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PusTextil 
            Height          =   195
            Left            =   9945
            TabIndex        =   228
            Top             =   135
            Width           =   240
            _Version        =   851968
            _ExtentX        =   423
            _ExtentY        =   344
            _StockProps     =   79
            Caption         =   "T"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.Label lblF10 
            Height          =   225
            Left            =   180
            TabIndex        =   221
            Top             =   450
            Width           =   540
            _Version        =   851968
            _ExtentX        =   952
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "<F10>"
         End
      End
      Begin VB.Frame fraDocAbrir 
         BorderStyle     =   0  'None
         Height          =   585
         Left            =   -150
         TabIndex        =   47
         Top             =   -150
         Width           =   135
         Begin MSComctlLib.Toolbar BarraCliente 
            Height          =   330
            Left            =   30
            TabIndex        =   48
            Top             =   150
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   582
            ButtonWidth     =   1588
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList2"
            DisabledImageList=   "ImageList2"
            HotImageList    =   "ImageList2"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   4
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "Cliente"
                  Object.ToolTipText     =   "Cliente"
                  ImageIndex      =   11
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Nuevo"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Buscar"
                  ImageIndex      =   10
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Style           =   3
               EndProperty
            EndProperty
            Begin MSComctlLib.ImageList ImageList1 
               Left            =   8580
               Top             =   0
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483633
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   -2147483644
               UseMaskColor    =   0   'False
               _Version        =   393216
               BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
                  NumListImages   =   12
                  BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":596A
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":5A7C
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":5B8E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":5CA0
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":5DB2
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":5EC4
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":5FD6
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":60E8
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":61FA
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":630C
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":641E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "frmRemito.frx":6530
                     Key             =   ""
                  EndProperty
               EndProperty
            End
         End
      End
      Begin VB.Frame fraPrecio 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   5700
         TabIndex        =   41
         Top             =   30
         Width           =   8895
         Begin XtremeSuiteControls.FlatEdit vDesRepartidor 
            Height          =   315
            Left            =   4380
            TabIndex        =   209
            Top             =   30
            Width           =   4455
            _Version        =   851968
            _ExtentX        =   7858
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4210752
            BackColor       =   -2147483633
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit vcodRepartidor 
            Height          =   315
            Left            =   3060
            TabIndex        =   208
            Top             =   30
            Width           =   795
            _Version        =   851968
            _ExtentX        =   1402
            _ExtentY        =   556
            _StockProps     =   77
            ForeColor       =   4210752
            BackColor       =   -2147483633
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin VB.ComboBox cboLista 
            Height          =   315
            ItemData        =   "frmRemito.frx":6642
            Left            =   1140
            List            =   "frmRemito.frx":6644
            TabIndex        =   42
            Text            =   "1"
            Top             =   30
            Width           =   585
         End
         Begin XtremeSuiteControls.FlatEdit txtEmpleados 
            Height          =   315
            Index           =   0
            Left            =   3150
            TabIndex        =   43
            Top             =   30
            Width           =   615
            _Version        =   851968
            _ExtentX        =   1085
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   6513507
            BackColor       =   6513507
            Locked          =   -1  'True
            MaxLength       =   3
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtEmpleados 
            Height          =   315
            Index           =   1
            Left            =   4380
            TabIndex        =   44
            Top             =   30
            Width           =   4440
            _Version        =   851968
            _ExtentX        =   7832
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   6513507
            BackColor       =   6513507
            Appearance      =   2
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton PushButton7 
            Height          =   345
            Left            =   3960
            TabIndex        =   207
            Top             =   0
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "F7"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.Label Label17 
            Height          =   315
            Left            =   1830
            TabIndex        =   210
            Top             =   30
            Width           =   1185
            _Version        =   851968
            _ExtentX        =   2090
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Vendedor :"
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label lbllista 
            AutoSize        =   -1  'True
            Caption         =   "> Lista de Precio:"
            Height          =   195
            Left            =   -150
            TabIndex        =   45
            Top             =   90
            Width           =   1230
         End
         Begin VB.Label Label1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   375
            Index           =   7
            Left            =   1770
            TabIndex        =   46
            Top             =   0
            Width           =   7095
         End
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   345
         Index           =   1
         Left            =   1380
         TabIndex        =   160
         ToolTipText     =   "Depura la Grilla de Detalles"
         Top             =   30
         Width           =   1395
         _Version        =   851968
         _ExtentX        =   2461
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Vaciar Detalle"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   0
         Picture         =   "frmRemito.frx":6646
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   345
         Index           =   0
         Left            =   0
         TabIndex        =   161
         ToolTipText     =   "Borra el Detalle Seleccionado de la Grilla"
         Top             =   30
         Width           =   1365
         _Version        =   851968
         _ExtentX        =   2408
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Borrar Detalle"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   0
         Picture         =   "frmRemito.frx":6A5F
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   345
         Index           =   2
         Left            =   2790
         TabIndex        =   162
         ToolTipText     =   "Depura la Grilla de Detalles"
         Top             =   30
         Width           =   1305
         _Version        =   851968
         _ExtentX        =   2302
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cargar Excel"
         UseVisualStyle  =   -1  'True
         TextAlignment   =   0
         Picture         =   "frmRemito.frx":6E7D
         TextImageRelation=   4
      End
      Begin XtremeSuiteControls.Label Label15 
         Height          =   165
         Left            =   4140
         TabIndex        =   182
         Top             =   120
         Width           =   555
         _Version        =   851968
         _ExtentX        =   979
         _ExtentY        =   291
         _StockProps     =   79
         Caption         =   "Lineas:"
      End
   End
   Begin TabDlg.SSTab TabTipoDetalle 
      Height          =   1650
      Left            =   16290
      TabIndex        =   39
      Top             =   4890
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   2910
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "frmRemito.frx":7417
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "frmRemito.frx":7433
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
   End
   Begin XtremeSuiteControls.GroupBox GBClientes 
      Height          =   1185
      Left            =   120
      TabIndex        =   25
      Top             =   1080
      Width           =   10065
      _Version        =   851968
      _ExtentX        =   17754
      _ExtentY        =   2090
      _StockProps     =   79
      ForeColor       =   255
      UseVisualStyle  =   -1  'True
      Begin VB.TextBox txtCodigoCliente 
         Height          =   330
         Left            =   7965
         TabIndex        =   230
         Top             =   135
         Width           =   915
      End
      Begin XtremeSuiteControls.PushButton PusNuevo 
         Height          =   315
         Left            =   9390
         TabIndex        =   127
         Top             =   150
         Width           =   615
         _Version        =   851968
         _ExtentX        =   1085
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Nuevo"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton cmdCambiarNombre 
         Height          =   315
         Left            =   9030
         TabIndex        =   81
         Top             =   150
         Width           =   345
         _Version        =   851968
         _ExtentX        =   609
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ComboBox cboTipoIva 
         Height          =   315
         Left            =   1035
         TabIndex        =   26
         Top             =   810
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   315
         Index           =   0
         Left            =   1050
         TabIndex        =   27
         Top             =   150
         Width           =   6690
         _Version        =   851968
         _ExtentX        =   11800
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   315
         Index           =   1
         Left            =   1035
         TabIndex        =   28
         Top             =   480
         Width           =   3720
         _Version        =   851968
         _ExtentX        =   6562
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   315
         Index           =   2
         Left            =   5625
         TabIndex        =   29
         Top             =   510
         Width           =   4380
         _Version        =   851968
         _ExtentX        =   7726
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   315
         Index           =   3
         Left            =   5640
         TabIndex        =   30
         Top             =   840
         Width           =   4365
         _Version        =   851968
         _ExtentX        =   7699
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblCliente 
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente: F1"
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
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   35
         Top             =   195
         Width           =   1245
      End
      Begin VB.Label lblCliente 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de IVA :"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   34
         Top             =   840
         Width           =   1245
      End
      Begin VB.Label lblCliente 
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección :"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   33
         Top             =   495
         Width           =   1245
      End
      Begin VB.Label lblCliente 
         BackStyle       =   0  'Transparent
         Caption         =   "Localidad :"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   4830
         TabIndex        =   32
         Top             =   555
         Width           =   825
      End
      Begin VB.Label lblCliente 
         Caption         =   "C.U.I.T:"
         Height          =   195
         Index           =   4
         Left            =   5040
         TabIndex        =   31
         Top             =   870
         Width           =   705
      End
   End
   Begin VB.CommandButton cmdBuscarArticulo 
      Enabled         =   0   'False
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
      Left            =   9900
      Picture         =   "frmRemito.frx":744F
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7890
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmdGuardarPrecio 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   315
      Left            =   10260
      Picture         =   "frmRemito.frx":7876
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Actualiza el precio del Articulo"
      Top             =   7260
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   345
   End
   Begin MSDataGridLib.DataGrid dgEmpleados 
      Height          =   495
      Left            =   15060
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   873
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      HeadLines       =   1
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
   Begin XtremeSuiteControls.GroupBox GBDocAFactura 
      Height          =   855
      Index           =   0
      Left            =   10530
      TabIndex        =   16
      Top             =   8100
      Visible         =   0   'False
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   1508
      _StockProps     =   79
      Caption         =   "IVA"
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton RBIva 
         Height          =   285
         Index           =   1
         Left            =   1530
         TabIndex        =   18
         Top             =   -60
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1411
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Sumar"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBIva 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   795
         _Version        =   851968
         _ExtentX        =   1411
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Restar"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
   End
   Begin MSAdodcLib.Adodc bfactura 
      Height          =   330
      Left            =   14160
      Top             =   8370
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   582
      ConnectMode     =   2
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
      Caption         =   "bfactura"
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
   Begin MSAdodcLib.Adodc bdetalle 
      Height          =   330
      Left            =   14160
      Top             =   8730
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
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
      Caption         =   "bdetalle"
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
   Begin MSAdodcLib.Adodc barticulo 
      Height          =   330
      Left            =   14160
      Top             =   8010
      Visible         =   0   'False
      Width           =   2505
      _ExtentX        =   4419
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
      Caption         =   "barticulo"
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
   Begin VB.Frame fraConfig 
      Caption         =   "Configuración :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   16260
      TabIndex        =   0
      Top             =   1740
      Visible         =   0   'False
      Width           =   2475
      Begin VB.CheckBox bienes 
         Alignment       =   1  'Right Justify
         Caption         =   "Bienes de Capital :"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   330
         TabIndex        =   2
         Top             =   990
         UseMaskColor    =   -1  'True
         Width           =   1635
      End
      Begin VB.ComboBox tprecio 
         Height          =   315
         ItemData        =   "frmRemito.frx":7C78
         Left            =   600
         List            =   "frmRemito.frx":7C85
         TabIndex        =   1
         Text            =   "Pesos ($)"
         Top             =   570
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "Tomar precio en:"
         Height          =   225
         Left            =   600
         TabIndex        =   3
         Top             =   330
         Width           =   1425
      End
   End
   Begin XtremeSuiteControls.GroupBox GBDocAFactura 
      Height          =   495
      Index           =   1
      Left            =   11910
      TabIndex        =   19
      Top             =   8520
      Visible         =   0   'False
      Width           =   2535
      _Version        =   851968
      _ExtentX        =   4471
      _ExtentY        =   873
      _StockProps     =   79
      Enabled         =   0   'False
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkBorrarDocOriginal 
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   180
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Borrar Documento Original"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.Frame fraNroInterno 
      Height          =   1125
      Left            =   16950
      TabIndex        =   21
      Top             =   450
      Visible         =   0   'False
      Width           =   1455
      Begin VB.TextBox txtNroRemito 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.Frame fraTipoDocumento 
      ForeColor       =   &H00808080&
      Height          =   540
      Left            =   30
      TabIndex        =   8
      Top             =   2670
      Width           =   14715
      Begin XtremeSuiteControls.CheckBox chkCaeTest 
         Height          =   315
         Left            =   8760
         TabIndex        =   192
         Top             =   150
         Width           =   675
         _Version        =   851968
         _ExtentX        =   1191
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Test"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vcaeFecha2 
         Height          =   285
         Left            =   12720
         TabIndex        =   191
         Top             =   150
         Width           =   1935
         _Version        =   851968
         _ExtentX        =   3413
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit vcae 
         Height          =   285
         Left            =   10920
         TabIndex        =   190
         Top             =   150
         Width           =   1755
         _Version        =   851968
         _ExtentX        =   3096
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PusCAE 
         Height          =   315
         Left            =   7830
         TabIndex        =   189
         Top             =   150
         Width           =   885
         _Version        =   851968
         _ExtentX        =   1561
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Fe C.A.E."
         ForeColor       =   0
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
      End
      Begin VB.OptionButton opOtrosDocumentos 
         BackColor       =   &H80000000&
         Caption         =   "Otros"
         Height          =   315
         Left            =   4350
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   585
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Nota Débito"
         Height          =   315
         Index           =   5
         Left            =   3330
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   150
         Width           =   1005
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Remito"
         Height          =   315
         Index           =   4
         Left            =   720
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   135
         Width           =   615
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Factura"
         Height          =   315
         Index           =   0
         Left            =   60
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   120
         Width           =   645
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Presupuesto"
         Height          =   315
         Index           =   1
         Left            =   1350
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   150
         Width           =   1005
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Nota Crédito"
         Height          =   315
         Index           =   2
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Width           =   1005
      End
      Begin VB.OptionButton opTipoDoc 
         BackColor       =   &H80000000&
         Caption         =   "Documento"
         Height          =   315
         Index           =   3
         Left            =   2370
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   150
         Width           =   945
      End
      Begin VB.CommandButton cmdNotaCredito 
         BackColor       =   &H80000000&
         Caption         =   "Sel. Factura p/  NC- ND"
         Height          =   315
         Left            =   5970
         MaskColor       =   &H80000000&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1845
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000000&
         Caption         =   "Fromulario 1116"
         Height          =   315
         Left            =   6660
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   150
         UseMaskColor    =   -1  'True
         Width           =   1305
      End
      Begin XtremeSuiteControls.CheckBox chkReingresarFact 
         Height          =   195
         Left            =   9450
         TabIndex        =   195
         Top             =   210
         Width           =   1425
         _Version        =   851968
         _ExtentX        =   2514
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Reingresar Fact"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   60
      ScaleHeight     =   1935
      ScaleWidth      =   14775
      TabIndex        =   111
      Top             =   720
      Width           =   14775
      Begin XtremeSuiteControls.PushButton push_consultar_doc 
         Height          =   345
         Left            =   10260
         TabIndex        =   231
         Top             =   1560
         Width           =   2355
         _Version        =   851968
         _ExtentX        =   4154
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Consultar doc. en AFIP"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusNroComp 
         Height          =   285
         Left            =   7065
         TabIndex        =   194
         Top             =   1530
         Width           =   2025
         _Version        =   851968
         _ExtentX        =   3572
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Consultar Nro. Comp AFIP"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   315
         Left            =   9240
         TabIndex        =   187
         Top             =   1530
         Width           =   885
         _Version        =   851968
         _ExtentX        =   1561
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Actualizar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit vnroremito2 
         Height          =   315
         Left            =   11880
         TabIndex        =   167
         Top             =   960
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox vcventa 
         Height          =   315
         Left            =   11910
         TabIndex        =   165
         Top             =   630
         Width           =   2505
         _Version        =   851968
         _ExtentX        =   4419
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Cuenta Corriente"
      End
      Begin XtremeSuiteControls.FlatEdit txtNroInterno 
         Height          =   285
         Left            =   11910
         TabIndex        =   113
         Top             =   330
         Width           =   2535
         _Version        =   851968
         _ExtentX        =   4471
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   14737632
         BackColor       =   14737632
         Alignment       =   1
      End
      Begin XtremeSuiteControls.ComboBox cboPuntoDeVenta 
         Height          =   315
         Left            =   2370
         TabIndex        =   116
         Top             =   1530
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         DropDownItemCount=   5
      End
      Begin XtremeSuiteControls.FlatEdit txtNroComprobante 
         Height          =   315
         Left            =   5160
         TabIndex        =   117
         Top             =   1530
         Width           =   1785
         _Version        =   851968
         _ExtentX        =   3149
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin XtremeSuiteControls.ComboBox cboLetra 
         Height          =   315
         Left            =   600
         TabIndex        =   121
         Top             =   1530
         Width           =   915
         _Version        =   851968
         _ExtentX        =   1614
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         DropDownItemCount=   5
      End
      Begin Aplisoft_CajasDeTexto.TxF vFechaIva 
         Height          =   285
         Left            =   11880
         TabIndex        =   185
         Top             =   1320
         Width           =   2565
         _ExtentX        =   4524
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
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   255
         Left            =   10140
         TabIndex        =   193
         Top             =   390
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Actualizar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Imp. IVA :"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   10620
         TabIndex        =   186
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro. Remito:"
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   10380
         TabIndex        =   166
         Top             =   1020
         Width           =   1425
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Condición de Venta: "
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   10380
         TabIndex        =   164
         Top             =   660
         Width           =   1830
      End
      Begin VB.Label lblLetra 
         BackStyle       =   0  'Transparent
         Caption         =   "Letra:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   60
         TabIndex        =   120
         Top             =   1590
         Width           =   465
      End
      Begin VB.Label lblSucursal 
         BackStyle       =   0  'Transparent
         Caption         =   "Sucursal:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   1560
         TabIndex        =   119
         Top             =   1590
         Width           =   855
      End
      Begin VB.Label lblNroComprobante 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Comprobante:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   3555
         TabIndex        =   118
         Top             =   1665
         Width           =   1665
      End
      Begin VB.Label lblCantidadRemito 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   8550
         TabIndex        =   115
         Top             =   1155
         Width           =   135
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. Interno:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   10950
         TabIndex        =   114
         Top             =   360
         Width           =   960
      End
      Begin VB.Label lblFecha 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   11340
         TabIndex        =   112
         Top             =   60
         Width           =   570
      End
   End
   Begin VB.Frame FraAccionesDoc 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   30
      TabIndex        =   4
      Top             =   -90
      Width           =   14775
      Begin Project1.bsGradientLabel lblsaldocliente 
         Height          =   375
         Left            =   12720
         Top             =   150
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   661
         BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.CheckBox checkSaldo 
         Height          =   255
         Left            =   9720
         TabIndex        =   159
         Top             =   210
         Width           =   1245
         _Version        =   851968
         _ExtentX        =   2196
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Mostrar Saldo"
         UseVisualStyle  =   -1  'True
         Value           =   1
      End
      Begin MSComctlLib.Toolbar BarraDocumento 
         Height          =   570
         Left            =   90
         TabIndex        =   5
         Top             =   90
         Width           =   9645
         _ExtentX        =   17013
         _ExtentY        =   1005
         ButtonWidth     =   1799
         ButtonHeight    =   1005
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Nuevo"
               Object.ToolTipText     =   "Nueva Factura"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "CtaCtes"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Guardar F2"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Imprimir F5"
               ImageIndex      =   12
            EndProperty
         EndProperty
         Begin XtremeSuiteControls.PushButton PushButton9 
            Height          =   225
            Left            =   4140
            TabIndex        =   223
            Top             =   60
            Visible         =   0   'False
            Width           =   645
            _Version        =   851968
            _ExtentX        =   1138
            _ExtentY        =   397
            _StockProps     =   79
            Caption         =   "Prueba"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton8 
            Height          =   285
            Left            =   4545
            TabIndex        =   211
            Top             =   180
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   503
            _StockProps     =   79
            Caption         =   "Duplicar 2"
            Appearance      =   6
         End
         Begin XtremeSuiteControls.GroupBox GroupBox5 
            Height          =   495
            Left            =   5700
            TabIndex        =   196
            Top             =   0
            Width           =   3225
            _Version        =   851968
            _ExtentX        =   5689
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Tipo de Impresión : "
            UseVisualStyle  =   -1  'True
            Begin XtremeSuiteControls.RadioButton rdFE 
               Height          =   255
               Left            =   60
               TabIndex        =   199
               Top             =   210
               Width           =   465
               _Version        =   851968
               _ExtentX        =   820
               _ExtentY        =   450
               _StockProps     =   79
               Caption         =   "FE"
               UseVisualStyle  =   -1  'True
               Value           =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton rdIF 
               Height          =   225
               Left            =   1350
               TabIndex        =   198
               Top             =   240
               Width           =   1095
               _Version        =   851968
               _ExtentX        =   1931
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Imp. Fiscal"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton rdT 
               Height          =   225
               Left            =   570
               TabIndex        =   197
               Top             =   240
               Width           =   795
               _Version        =   851968
               _ExtentX        =   1402
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Ticket"
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.RadioButton rdotros 
               Height          =   225
               Left            =   2460
               TabIndex        =   212
               Top             =   240
               Width           =   675
               _Version        =   851968
               _ExtentX        =   1191
               _ExtentY        =   397
               _StockProps     =   79
               Caption         =   "Otros"
               UseVisualStyle  =   -1  'True
            End
         End
         Begin MSComctlLib.ImageList ImageList2 
            Left            =   9060
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483633
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   -2147483644
            UseMaskColor    =   0   'False
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   12
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":7CAC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":7DBE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":7ED0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":7FE2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":80F4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":8206
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":8318
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":842A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":853C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":864E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":8760
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmRemito.frx":8872
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin XtremeSuiteControls.PushButton PBDocAFactura 
         Height          =   525
         Index           =   0
         Left            =   4890
         TabIndex        =   37
         Top             =   180
         Visible         =   0   'False
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Doc. A Factura"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "frmRemito.frx":8984
      End
      Begin XtremeSuiteControls.PushButton PBDocAFactura 
         Height          =   525
         Index           =   1
         Left            =   3960
         TabIndex        =   38
         Top             =   180
         Visible         =   0   'False
         Width           =   945
         _Version        =   851968
         _ExtentX        =   1667
         _ExtentY        =   926
         _StockProps     =   79
         Caption         =   "Vista Previa"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "frmRemito.frx":8F86
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo de la cuenta:"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   11190
         TabIndex        =   157
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "ver ..."
      Height          =   255
      Index           =   0
      Left            =   11310
      TabIndex        =   123
      Top             =   60
      Width           =   525
   End
   Begin VB.CommandButton cmdVer 
      Caption         =   "ver ..."
      Height          =   255
      Index           =   1
      Left            =   13170
      TabIndex        =   122
      Top             =   15
      Width           =   555
   End
   Begin VB.CheckBox cf 
      Caption         =   "Dejarla Fija"
      Height          =   165
      Left            =   12480
      TabIndex        =   6
      Top             =   360
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "> Presupuesto: "
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   10020
      TabIndex        =   126
      Top             =   -30
      Width           =   1110
   End
   Begin VB.Label lblCantidadPresupuesto 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   11160
      TabIndex        =   125
      Top             =   -30
      Width           =   105
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BackStyle       =   0  'Transparent
      Caption         =   "> Remitos:"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   12090
      TabIndex        =   124
      Top             =   30
      Width           =   750
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000080FF&
      X1              =   90
      X2              =   9555
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Menu guardarOtro 
      Caption         =   "Guardar Otro"
   End
   Begin VB.Menu proponeritems 
      Caption         =   "Proponer Items"
   End
   Begin VB.Menu cambiaritems 
      Caption         =   "Cambiar un Items"
   End
   Begin VB.Menu pruebacf 
      Caption         =   "Prueba CF"
      Begin VB.Menu factahasar 
         Caption         =   "Fact A - HASAR"
      End
      Begin VB.Menu ComenzarHasar 
         Caption         =   "Comenzar - Hasar"
      End
      Begin VB.Menu nofiscalhasar 
         Caption         =   "No fiscal - Hasar"
      End
      Begin VB.Menu nc 
         Caption         =   "NC"
      End
      Begin VB.Menu nc2 
         Caption         =   "NC2"
      End
   End
   Begin VB.Menu estafiscal 
      Caption         =   "Estado Fiscal"
   End
   Begin VB.Menu mglobal 
      Caption         =   "Mensaje Global"
   End
   Begin VB.Menu cierrez 
      Caption         =   "Cierre Z"
   End
   Begin VB.Menu log1 
      Caption         =   "Log"
      Begin VB.Menu activarlog 
         Caption         =   "Activar log"
      End
      Begin VB.Menu Desactivarlog 
         Caption         =   "Desactivar log"
      End
   End
End
Attribute VB_Name = "frmRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim global_vpunto_venta, global_vtipoFactura, global_vnro_comprobante As String

Dim vnrocomprobante_control1 As String

Dim vtipoFactura, manualcae As Integer

Dim vdatos_mandante As String

Dim filaSeleccionada As Integer

Dim arr2(100, 2) As String

Dim vCodigoBarra As String

Dim vtipoDocDescript As String

Dim vultimoMensajeError As String
Dim vglblVtoCAE, vglblCAE, vetipoFactura, vnrocomprobante2 As String
Dim vcaeFecha As String
Dim vncomprobante As Long

Dim dr As New ifacta

Dim vCodigoRepartidor2 As String

Dim vestructura As String

Dim vduplicando As Boolean

Dim vnrointerno As Long

Dim vIdFactura, vIdEmpresa, vidVendedor As Long

Dim vsaldodeudor As Double

Public vsaldo11, vtotal11 As Double

Dim sCmd, sCmdExt As String
Dim bAnswer, vcancelartrans As Boolean

Dim vcbarra As String

Public vActualizaNombre As Boolean
Dim vCambioTipoDocumento As Boolean
Public vNombreNuevo, vfacturaDuplicadaMensaje As String
'Dim vIdFactura As Long
Dim vLeyendaAsiento As String, vTotalAsiento As Double
Dim vcol, c, tallefila  As Integer
Dim vtotal_real, vtotal_global, vpcosto  As Double
Dim f5 As Integer
Public vvcodigo, vvvdescrip, vvvcodigo, vcheque, vobservacion  As String
Dim vpespecial As Boolean
Public vvpdolar, vganancia As Double
Dim venvase As Boolean
'Dim vNroComprobante As Integer

' ---- Datos de la Nc seleccionada desde el formulario frmNroFacNC
Public vNroFacturaNotaC As Long
Public vLetraNotaC As String

Dim vLetra, vnrocomprobante, vPtoVta As String

Public vPuntoDeVentaNotaC As String
'-------------------------------------

Public vlistachofer As String
' 1 Modifica
Public vGrabaModo, vTipoDocumento, cargando As Integer
Dim ban As String
'----------------------------------------
'Control de Errores del Metodo Guardar
Dim vRemitoControl As Long
Dim vCantidadControl As Integer
'----------------------------------------
Dim vOpenGrilla() As Boolean
Dim vremito As Long
Dim vnlista As Integer
Dim vNoSaveDoc As Boolean
Dim checksum() As Boolean

Dim vArticuloNuevo As Boolean
'Dim vClienteNuevo As Boolean

'Const t = 8
Dim rsArticulos As ADODB.Recordset
Dim rsClientes As ADODB.Recordset
Dim rsEmpleadosGrilla As ADODB.Recordset
Dim vHabilitaDocAFactura As Boolean


'-------------------------------------------------
Dim vnroasiento, vnrobalance As Integer
'-------------------------------------------------
Dim vgTsubtotal, vgTiva105, vgTiva21, vgTiva27, vgTPdescuento, vgTimpuesto, vgTtotal As Double
Dim vgNroFactura As String

Private Enum MedioPago
    efectivoPesos = 1
    efectivoDolar = 2
    tarjeta = 3
    cheque = 4
    Deposito = 5
    NotaC = 8
    ContadoCredito = 11
    AjusteCredito = 12
End Enum

Private Sub activarlog_Click()
log.Visible = True
End Sub


Function getTipoDocDescp() As String
Dim v As String
Dim i As Integer


For i = 0 To Me.opTipoDoc.UBound
    If Me.opTipoDoc(i).Value Then v = Me.opTipoDoc(i).Caption
Next


getTipoDocDescp = v

End Function


Private Sub BarraDocumento_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next

' vestructura = Val(Me.vcodEmpresa)

vestructura = (Me.vcodEmpresa)

vtipoDocDescript = getTipoDocDescp

'If LeerXml("MostrarSaldoEnDoc") = "SI" Then Printer.PaperSize = vbPRPSLetterSmall

vnrocomprobante = Val(Me.txtNroComprobante)

vsaldo11 = Val(txtTotal.Text)

Call getTipoParaFE

vsaldo112 = Me.vsaldo11
lblsaldocliente2 = Val(Me.lblsaldocliente.Caption)

' vCodigoBarra = feCodigoBarra(Me.txtCliente(3), vtipoFactura, cboPuntoDeVenta, Me.vcae, Me.vcaeFecha2)

    Select Case Button.Index

        Case 1
            NuevoDoc

        Case 2
            CtaCte
        Case 3
            
            Me.BarraDocumento.Enabled = False
            
            manualcae = 1
            guardarinicio
        
            Me.vcae.Text = ""
            Me.vcaeFecha2.Text = ""
            
            If vcancelartrans Then Exit Sub
      
            If LeerXml("IncluyeContabilidad") = "True" Then frmAsientosAlta.SetFocus
            
           MsgBox "Presione <Enter> para continuar"
             
            Me.BarraDocumento.Enabled = True
            
            
        Case 4
        
        
            If Not validarGrabar Then
                    MsgBox "Verificar el nro de comprobante y el punto de venta", vbCritical
                    Exit Sub
            End If

            If vcaeFecha = 0 And Not Me.vcaeFecha2.Text = "" Then vcaeFecha = Val(Me.vcaeFecha2.Text)
            
            
            log.AddItem ("-------------------------------------------------------------------------------")
            
            
            fimprimirDoc
             
             ' If existeRegistro(Val(vnrointerno)) Then Exit Sub
            ' multiempresa ---------------------------------------------------
            
            vIdFactura = utltimoFactura
            vIdEmpresa = codigo2id(vcodEmpresa.Text)
            vidVendedor = codigo2id(Me.vcodRepartidor.Text)
            
            Call GuardarRel(vIdFactura, vIdEmpresa, vidVendedor, vnrointerno)
            Me.vcventa.Text = "Contado"
            If vcancelartrans Then Exit Sub
            
            Me.vcae.Text = ""
            Me.vcaeFecha2.Text = ""
            
           If vcancelartrans Then Exit Sub
            
        '   MsgBox "Presione <Enter> para continuar"
            
            
            If LeerXml("IncluyeContabilidad") = "True" Then frmAsientosAlta.SetFocus
            
        
            If Me.vcventa.Text = "Contado" And Not Me.chkReingresarFact = 1 Then
                frmCobros.SetFocus
            End If
            
            
            
    End Select

    If Err Then GrabarLog "BarraDocumento_ButtonClick (" & Button & ")", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Function validarGuardar() As Boolean
Dim vmensaje As String

vmensaje = ""
validarGuardar = True

'If vtotal > 0 Then
'    vmensaje = "El total de la factura es cero"
'End If

If Not validarCUIT2(Me.txtCliente(3).Text) = 1 Then
    vmensaje = vmensaje + "El CUIT está mal conformado. " + Chr(13) + "Quiere continuar de todas menera ?"
End If

If Not vmensaje = "" Then
   If MsgBox(vmensaje, vbYesNo) = vbYes Then
    validarGuardar = True
    
    If (Me.cboTipoIva.Text = "Consumidor Final" Or Trim(Me.cboTipoIva.Text) = "") And Val(Me.txtCliente(3).Text) > 0 And rdT.Value Then
        Me.txtCliente(3).Text = "25918294" ' tengo que ponerle un nro de doc para el controlador fiscal
        validarGuardar = True
    Else
       ' If Not UCase(LeerXml("Puesto")) = "DIEGO" Then validarGuardar = False
    End If

    End If
End If
End Function

Function obtenerCAE() As Boolean

Call getTipoParaFE


If Not UCase(LeerXml("ObtieneCAE")) = "SI" Then
 obtenerCAE = True
 Exit Function
End If



If Me.opTipoDoc(3).Value Or chkReingresarFact.Value = xtpChecked Or Me.opTipoDoc(1).Value Or Me.opTipoDoc(1).Value Then
    obtenerCAE = True
    'chkReingresarFact.Value = xtpUnchecked
    
    'If Not UCase(LeerXml("ObtieneCAE")) = "SI" Then chkReingresarFact.Value = xtpChecked
    
    Exit Function
End If

    Call PusCAE_Click
    
    
    If Val(Me.vcae.Text) = 0 Then
        obtenerCAE = False
    Else
        obtenerCAE = True
    End If
    
    
    If (vcae = "" Or vcaeFecha = "") And LeerXml("ObtieneCAE") = "SI" Then
            MsgBox "Problema al obtener el C.A.E." + Chr(10) + "Consulte al servicio técnico.", vbCritical, "No se prodrá confeccionar el documento."
            obtenerCAE = False
    End If

End Function


Private Sub verificar_nrointerno(v As String)
On Error Resume Next
Dim vv As Long
vv = 0
Dim i As Integer

i = 0
If Val(v) = 1 Then
        
        Do Until vv > 1 Or i > 10
         i = i + 1
            vv = UltimoNroInterno2 + 1
        
        Loop


Me.txtNroInterno.Text = Str(vv)

End If




If Err Then Exit Sub
End Sub

Private Sub guardarinicio()



If Not obtenerCAE Then Exit Sub

        If Me.vGrabaModo = 0 Then
            txtNroRemito.Text = NroRemitoNuevo
            txtNroInterno.Text = UltimoNroInterno2
        End If
        
Dim i As Integer

  Do Until Not existeRegistro(Val(Me.txtNroInterno))
            txtNroInterno.Text = UltimoNroInterno2
  Loop
        
       ' verificar_nrointerno (txtNroInterno.Text)
        
        
        If Not validarGuardarDocumento Then Exit Sub
        
           ' GuardarCompleto
           
            If GuardarCompleto Then Exit Sub
            
           ' vidVendedor = Val(Me.vcodRepartidor.Tag)
            
      '  vIdEmpresa = Val(Me.vcodEmpresa.Tag)
            
            
            
          vIdFactura = utltimoFactura
          vIdEmpresa = codigo2id(vcodEmpresa.Text)
          vidVendedor = codigo2id(Me.vcodRepartidor.Text)
  
            
           Call GuardarRel(vIdFactura, vIdEmpresa, vidVendedor, vnrointerno)
            
            RecargarForm
            limpiarCliente
            
            vnrocomprobante = 0
            
            If Me.vcventa.Text = "Contado" And Me.chkReingresarFact = 1 Then
                frmCobros.Show
                frmCobros.txtImporteEfectivo.SetFocus
            End If
            
            
End Sub


Function validarGuardarDocumento() As Boolean
Dim vmen As String

vmen = ""

If Val(Me.txtNroRemito.Text) = 0 Then vmen = vmen + "- Índice de comprobante" + Chr(13)
If Val(Me.txtNroInterno) = 0 Then vmen = vmen + "- Índice interno" + Chr(13)

If Not vmen = "" And Not Me.chkReingresarFact = 1 Then
    MsgBox vmen, vbCritical, "Cuidado"
    validarGuardarDocumento = False
Else
    validarGuardarDocumento = True
End If

If Me.dtpFecha.Text = "" Then
    vmen = vmen + "- Fecha inválida"
    MsgBox vmen, vbCritical, "Cuidado"
    validarGuardarDocumento = False
End If
End Function

Private Sub ImprimirNotaCenHasar()
On Error Resume Next

    '1 - Guardar Como Nota C - A o Nota C - B
    '2 - Ver a donde paga
    '3 - Imprimir

    
    If Not vNroFacturaNotaC = 0 Then
         If GuardarCompleto Then Exit Sub
        'Call GuardarCompleto
        Call ImprimirHasar(vremito, 0)
    
    Else
        
        MsgBox "Debe Seleccionar una factura para poder generar una Nota de Credito", vbInformation, "Mensaje ..."
    End If
    
If Err Then GrabarLog "ImprimirNotaCenHasar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub ImprimirEpson(vremito As Long, vMontoEnEF As Double)
On Error Resume Next

    Dim FS, vvtipodoc, vvtipodoc2 As String
    Dim vdescuento As Double
    
    
    vdescuento = 0
    
    Dim cf As New Fiscal
    
    FS = Chr$(28) '// Separador de campos del comando

    Dim rsImprimirHasar As New ADODB.Recordset, sqlImprimirHasar As String
    
    
    sqlImprimirHasar = "SELECT * FROM ImpresionFactura WHERE (Remito = " & Val(vremito) & ")"

    With rsImprimirHasar
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        Call .Open(sqlImprimirHasar, ConnDDBB, adOpenStatic, adLockReadOnly)
    End With
    
        
    Dim rsDetalleHasar As New ADODB.Recordset, sqlDetalleHasar As String
    sqlDetalleHasar = "SELECT * FROM FDetalle WHERE (Remito = " & Val(vremito) & ") ORDER BY idFDetalle ASC"

    If rsDetalleHasar.State = 1 Then rsDetalleHasar.Close
    rsDetalleHasar.CursorLocation = adUseClient
        
    Call rsDetalleHasar.Open(sqlDetalleHasar, ConnDDBB, adOpenStatic, adLockReadOnly)
    
    vvtipodoc = EsNulo(rsImprimirHasar.Fields("Tipo").Value)

    ' lógica para definir el comprobante
    vvtipodoc = Me.cboTipoIva
    
    'vvtipodoc = ""
    
    vvtipodoc2 = "Factura"
    
    If Me.opTipoDoc(3).Value Then
        vvtipodoc2 = "Documento"
    End If
    
    If Me.opTipoDoc(2).Value Then
        vvtipodoc2 = "Nota Credito"
    End If
    
    If Me.opTipoDoc(0).Value Then
        vvtipodoc2 = "Factura"
    End If
    
    vmensajeGlobal = ""
    
    With cf
        
        Select Case vvtipodoc
            
            Case "Iva Responsable Inscripto"
                                    ' Paso1 ----------------
                                    
                                    Call .DatosClientesTA2(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50), vvtipodoc2)
                                  
                                    ' Paso 2 ----------------
                                    Do Until rsDetalleHasar.EOF = True
                                        Call .ImprimirItemTA2(EsNulo(rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), Val(rsDetalleHasar.Fields("TIVa").Value), 0, , vvtipodoc2)
                                        rsDetalleHasar.MoveNext
                                    Loop
                                
                                    ' Paso 3 --------------------
                                
                                    'Call .ImprimirPago("Efectivo", vMontoEnEF)  'Val(GenerarDato("SELECT SUM(Monto) AS TotalEF FROM Recibo_Temp WHERE IdMedioPago = 1 GROUP BY idMedioPago;", "TotalEF")))
                                
                                    ' Paso 4 ---------------------
                                    
                                    Call .Cerrarcomprobate(vdescuento, vvtipodoc2)
                            
                                    ' Paso 5 ---------------------

            
            Case "Iva Exento"
                                      
                                  ' Paso1 ----------------
                                    
                                    Call .DatosClientesTA2(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, "Exento", Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50), vvtipodoc2)
                                  
                                    ' Paso 2 ----------------
                                    Do Until rsDetalleHasar.EOF = True
                                        Call .ImprimirItemTA2(EsNulo(rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), 0, 0, , vvtipodoc2)
                                        rsDetalleHasar.MoveNext
                                    Loop
                                
                                    ' Paso 3 --------------------
                                
                                    'Call .ImprimirPago("Efectivo", vMontoEnEF)  'Val(GenerarDato("SELECT SUM(Monto) AS TotalEF FROM Recibo_Temp WHERE IdMedioPago = 1 GROUP BY idMedioPago;", "TotalEF")))
                                
                                    ' Paso 4 ---------------------
                                    
                                    Call .Cerrarcomprobate(vdescuento, vvtipodoc2)
                            
                                    ' Paso 5 ---------------------

                                   
             Case "Consumidor Final"
                                   
                                    Call .DatosClientesCF(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50), vvtipodoc2)
                                    
                                   
                                      Do Until rsDetalleHasar.EOF = True
                                          Call .ImprimirItemCF(EsNulo(rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), Val(rsDetalleHasar.Fields("TIVa").Value), 0, , vvtipodoc2)
                                          rsDetalleHasar.MoveNext
                                      Loop
                                      
                                     Call .CerrarcomprobateCF(vdescuento, vvtipodoc2)
                                                     
                                   
            Case "Responsable Monotributo"
                                      
                                 
                                      Call .DatosClientesCF(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50), vvtipodoc2)
                                   '
                                      Do Until rsDetalleHasar.EOF = True
                                          Call .ImprimirItemCF(EsNulo(rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), Val(rsDetalleHasar.Fields("TIVa").Value), 0)
                                          rsDetalleHasar.MoveNext
                                      Loop
                                      
                                     Call .CerrarcomprobateCF(vdescuento)
                                   
                                   
                                   '   Call .CerrarcomprobateCFOriginal
                                      
             '   Case "NC"
                                      ' nota de credito
                                 
             '                         Call .DatosClientesNC(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                                   
             '                         Do Until rsDetalleHasar.EOF = True
             '                             Call .ImprimirItemNC(EsNulo(rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), Val(rsDetalleHasar.Fields("TIVa").Value), 0)
             '                             rsDetalleHasar.MoveNext
             '                         Loop
             '
             '                        Call .CerrarcomprobateNC(vdescuento)
                                   
                                   
                                   '   Call .CerrarcomprobateCFOriginal
                                                                
               Case "Documento"
                      
                                   ' Paso1 ----------------
                                    Call .DatosClientesTNF(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                                  
                                    ' Paso 2 -------------
                                
                                    '------------
           Dim i As Integer
           i = 1
bAnswer = True
'With frmPrincipal.FiscalEpson2
'            For i = 1 To 5
'                sCmd = Chr$(&HE) + Chr$(&H2)
'                If bAnswer Then bAnswer = .AddDataField(sCmd)
'                sCmdExt = Chr$(&H0) + Chr$(&H0)
'               If bAnswer Then bAnswer = .AddDataField(sCmdExt)
'                If bAnswer Then bAnswer = .AddDataField("Texto Texto Texto Texto Texto Texto")
'                MsgBox ""
'
'                If bAnswer Then bAnswer = .SendCommand
'                Call FPDelay
'                If .ReturnCode <> 0 Then ShowMsg
'            Next
' End With
                                    '------------
                                    
                                    
                                    Do Until rsDetalleHasar.EOF = True
                                        Call .ImprimirItemTNF(EsNulo(rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), Val(rsDetalleHasar.Fields("TIVa").Value), 0)
                                        rsDetalleHasar.MoveNext
                                    Loop
          
                        bAnswer = True
                        With frmPrincipal.FiscalEpson2
                                sCmd = Chr$(&HE) + Chr$(&H2)
                                If bAnswer Then bAnswer = .AddDataField(sCmd)
                                sCmdExt = Chr$(&H0) + Chr$(&H0)
                                If bAnswer Then bAnswer = .AddDataField(sCmdExt)
                                If bAnswer Then bAnswer = .AddDataField("Total: " + Format(txtTotal.Text, "###,###,##0.00"))
                               ' If bAnswer Then bAnswer = .AddDataField("")
                
                                If bAnswer Then bAnswer = .SendCommand
                                Call FPDelay
                                If Me.EpsonFP.ReturnCode <> 0 Then ShowMsg
                                    
                       End With
                         
                                    Call .CerrarcomprobateNF(txtTotal.Text)
                            
            Case Else
               
                                    Call .DatosClientesCF(EsNulo(rsImprimirHasar.Fields("Nombre").Value), EsNulo(rsImprimirHasar.Fields("cuit").Value), TIPO_DNI, CONSUMIDOR_FINAL, Left(rsImprimirHasar.Fields("Direccion").Value, 50))
                    
                                    Do Until rsDetalleHasar.EOF = True
                                        Call .ImprimirItemCF(EsNulo(rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), Val(rsDetalleHasar.Fields("TIVa").Value), 0)
                                        rsDetalleHasar.MoveNext
                                    Loop
                                    
        End Select
        
    End With ' fiscal
    
    
If Err Then
    Call GrabarLog("ImprimirHasar", Err.Number & " " & Err.Description, Me.Caption)
    Call MsgBox("Error Impresora:" & Err.Description, vbCritical, "Errores")
Else
    'vImpresionCorrecta = True
End If
End Sub


Private Sub ImprimirHasar(vremito As Long, vMontoEnEF As Double)  ' HASAR - FISCAL -
On Error Resume Next

    Dim FS, vmensaje As String
    
    FS = Chr$(28) '// Separador de campos del comando
    
  
  
  ' init ----------------------------------------------------------------------------------------------
  
If Not UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
            With frmPrincipal.FiscalEpson2
            
                                       .ClosePort
                                       
                                       Call FPDelay
                                       
                                       .CommPort = LeerXml("Puerto")
                                       .BaudRate = 3
                                       
                                       .ProtocolType = protocol_Extended
                                       
                                       
                                       vmensaje = " Puerto : " + Str(.CommPort) + "  - " + Str(.BaudRate)
                                       
                                       MsgBox vmensaje
                                       
                                       
                                       If (.OpenPort) Then
                                           Call FPDelay
                                       Else
                                           MsgBox "2- El controlador fiscal no está conectado. " + Chr(13) + _
                                           "Conecte el controlador y vuelva a ingresar a este módulo"
                                       End If
            
            End With
                                                     
                                                     
 End If
'--------------------------------------------------------------------------------------------------
                            
  


    Dim rsImprimirHasar As New ADODB.Recordset, sqlImprimirHasar As String
    
   ' MsgBox "Prepare la Impresora ", vbInformation, "Mensaje ..."
    
    sqlImprimirHasar = "SELECT * FROM ImpresionFactura WHERE (Remito = " & Val(vremito) & ")"

    With rsImprimirHasar
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        Call .Open(sqlImprimirHasar, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .State = 1 Then
            If Not .EOF = True Then
                'Va todo bien
            Else
                MsgBox "Remito Nro : " & remito
                Exit Sub
            End If
        Else
            MsgBox "No se pudo abrir la Factura de Venta/Nota C", vbCritical, "Mensaje ..."
        End If
    End With
    
    With frmPrincipal.FiscalHasar
        
        
        log.AddItem " Pasa por encabezado .... "
        'Call .EspecificarNombreDeFantasia(" ", " ")
        '.Encabezado(1) = EsNulo(UCase(vDatosEmpresa.Nombre))
        .Encabezado(1) = EsNulo(UCase(vDatosEmpresa.Direccion))
        .Encabezado(2) = EsNulo(UCase(vDatosEmpresa.Localidad))
        '.Encabezado(4) = EsNulo(UCase(vDatosEmpresa.CondicionIva)) & "            " & EsNulo(UCase(vDatosEmpresa.CUIT))
        .Encabezado(3) = EsNulo(UCase(vDatosEmpresa.Telefono))
        
        Select Case EsNulo(rsImprimirHasar.Fields("TipoIva").Value)
            
            Case "Iva Responsable Inscripto"
                .PrecioBase = True
                
                
                If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket Hasar") Or UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
                    
                                                '  MsgBox ".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, RESPONSABLE_INSCRIPTO," + "  " + Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50)
                    
                    Call .DatosCliente(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                    log.AddItem (".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, RESPONSABLE_INSCRIPTO," + "  " + Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                    Debug.Print (".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, RESPONSABLE_INSCRIPTO," + "  " + Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                                                'MsgBox ".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, RESPONSABLE_INSCRIPTO," + "  " + Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50)
                End If
                                                 '  Call .DatosCliente(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("CodigoPostal").Value) & "-" & EsNulo(rsImprimirHasar.Fields("Localidad").Value) & "     -     " & EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                
                
                If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket Epson") Then
                    '' todo
                    '' Fiscal.DatosClientes(EsNulo(rsImprimirHasar.Fields("Nombre").Value), Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, RESPONSABLE_INSCRIPTO, Left(EsNulo(rsImprimirHasar.Fields("CodigoPostal").Value) & "-" & EsNulo(rsImprimirHasar.Fields("Localidad").Value) & "     -     " & EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50)))
                End If
          
            Case "Responsable Monotributo"
                .PrecioBase = False
                Call .DatosCliente(rsImprimirHasar.Fields("Nombre").Value, Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, 77, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                
                log.AddItem (".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, 77," + "  " + Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))

                Debug.Print (".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, RESPONSABLE_INSCRIPTO," + "  " + Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))

            
            Case "Iva Exento"
                .PrecioBase = False
                
                Call .DatosCliente(rsImprimirHasar.Fields("Nombre").Value, Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, 69, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                'Call .DatosCliente(rsImprimirHasar.Fields("Nombre").Value, Replace(rsImprimirHasar.Fields("Cuit").Value, "-", ""), TIPO_CUIT, 77, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))

                'Debug.Print (".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, RESPONSABLE_EXENTO, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                log.AddItem (".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, 69," + "  " + Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
            
            
            Case "Consumidor Final"
                .PrecioBase = False
                Call .DatosCliente(EsNulo(rsImprimirHasar.Fields("Nombre").Value), EsNulo(rsImprimirHasar.Fields("cuit").Value), TIPO_DNI, CONSUMIDOR_FINAL, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))
                
                Debug.Print (".DatosCliente(" + EsNulo(rsImprimirHasar.Fields("Nombre").Value) + "  " + Replace(rsImprimirHasar.Fields("Cuit").Value, "-", "") + "  " + "TIPO_CUIT, RESPONSABLE_INSCRIPTO," + "  " + Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))

                'Debug.Print EsNulo(rsImprimirHasar.Fields("Nombre").Value), EsNulo(rsImprimirHasar.Fields("NroDocumento").Value), TIPO_DNI, CONSUMIDOR_FINAL, Left(EsNulo(rsImprimirHasar.Fields("Direccion").Value), 50))r.Fields("CodigoPostal").Value) & "-" & EsNulo(rsImprimirHasar.Fields("Localidad").Value) & " - " & EsNuloGuion(rsImprimirHasar.Fields("Direccion").Value), 50)
        End Select
        
   
        Select Case EsNulo(rsImprimirHasar.Fields("Tipo").Value)
        
            Case "Fact A"
            
                'Call .AbrirComprobanteFiscal(FACTURA_A)   ' todoing machete
                 
                 If UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
                 
                    Call .AbrirComprobanteFiscal(FACTURA_A)   ' todoing machete
                 
                 Else
                 
                    Call .AbrirComprobanteFiscal(TICKET_FACTURA_A)
                
                End If
                
                
                log.AddItem ".AbrirComprobanteFiscal(TICKET_FACTURA_A)"
              
            Case "Ticket-Factura"
                Call .AbrirComprobanteFiscal(TICKET_FACTURA_A)
              
            Case "Fact B"
                'Call .AbrirComprobanteFiscal(FACTURA_B)
                 
                If UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
                 
                    .AbrirComprobanteFiscal (FACTURA_B)
                    log.AddItem (".AbrirComprobanteFiscal (FACTURA_B)")
                 
                Else
                
                    Call .AbrirComprobanteFiscal(TICKET_FACTURA_B)
                    log.AddItem (".AbrirComprobanteFiscal (TICKET_FACTURA_B)")
                
                End If
            
            Case "Documento"
                Call .AbrirComprobanteNoFiscal

            Case "Nota C"
                
                .DocumentoDeReferencia(1) = "0003-" & FormatoNroFactura(vNroFacturaNotaC)
                
                If EsNulo(rsImprimirHasar.Fields("idTipoIva").Value) = "001" Then
                   
                    If UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
                        .AbrirDNFH (NOTA_CREDITO_A)
                         log.AddItem (".AbrirDNFH NOTA_CREDITO_A")
                    Else
                        .AbrirComprobanteFiscal (TICKET_NOTA_CREDITO_A)
                        
                        log.AddItem (".AbrirComprobanteNoFiscalHomologado TICKET_NOTA_CREDITO_A")
                    End If
                
                Else
                   
                   If UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
                        .AbrirDNFH (NOTA_CREDITO_B)
                        
                        log.AddItem (".AbrirDNFH NOTA_CREDITO_B")
                    Else
                        .AbrirComprobanteFiscal (TICKET_NOTA_CREDITO_B)
                        log.AddItem (".AbrirComprobanteFiscal TICKET_NOTA_CREDITO_B")
                    End If
                
                End If
            
            
            Case "Nota D"
                
                .InformacionRemito(1) = "0003-" & FormatoNroFactura(vNroFacturaNotaC)
                
                If EsNulo(rsImprimirHasar.Fields("idTipoIva").Value) = "001" Then
                   
                    If UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
                        
                        .AbrirDNFH (NOTA_DEBITO_A)
                        
                        log.AddItem (".AbrirDNFH NOTA_CREDITO_A")
                         
                    Else
                        
                        .AbrirDNFH (TICKET_NOTA_DEBITO_A)
                        
                        log.AddItem (".AbrirDNFH TICKET_NOTA_CREDITO_A")
                    
                    End If
                
                Else
                   
                   If UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
                        .AbrirDNFH (NOTA_DEBITO_B)
                        
                        log.AddItem (".AbrirDNFH NOTA_CREDITO_B")
                    Else
                        .AbrirDNFH (TICKET_NOTA_DEBITO_B)
                        log.AddItem (".AbrirDNFH TICKET_NOTA_DEBITO_B")
                    End If
                
                End If
            
            Case "Remito"
                '
        End Select
        
        Dim rsDetalleHasar As New ADODB.Recordset, sqlDetalleHasar As String
        
        sqlDetalleHasar = "SELECT * FROM FDetalle WHERE (Remito = " & Val(vremito) & ") ORDER BY idFDetalle ASC"
        
        If rsDetalleHasar.State = 1 Then rsDetalleHasar.Close
        rsDetalleHasar.CursorLocation = adUseClient
        
        Call rsDetalleHasar.Open(sqlDetalleHasar, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not rsDetalleHasar.State = 1 Then
            If rsDetalleHasar.EOF = True Then
                'Panic : Paso algo con el detalle
            Else
                'Va todo bien
            End If
        End If
    
        Do Until rsDetalleHasar.EOF = True
            
            Select Case EsNulo(rsImprimirHasar.Fields("Tipo").Value)
        
                Case "Fact A", "Ticket-Factura", "Fact B", "Nota C"
                    Call .ImprimirItem((rsDetalleHasar.Fields("Detalle").Value), Val(rsDetalleHasar.Fields("Cantidad").Value), Val(rsDetalleHasar.Fields("Precio").Value), Val(rsDetalleHasar.Fields("TIVa").Value), 0)
                    
                    log.AddItem ".ImprimirItem" + " " + EsNulo(rsDetalleHasar.Fields("Detalle").Value) + "  " + Str(Val(rsDetalleHasar.Fields("Cantidad").Value)) + "  " + Str(Val(rsDetalleHasar.Fields("Precio").Value)) + "  " + Str(Val(rsDetalleHasar.Fields("TIVa").Value))
                    
                    Debug.Print ".ImprimirItem" + " " + EsNulo(rsDetalleHasar.Fields("Detalle").Value) + "  " + Str(Val(rsDetalleHasar.Fields("Cantidad").Value)) + "  " + Str(Val(rsDetalleHasar.Fields("Precio").Value)) + "  " + Str(Val(rsDetalleHasar.Fields("TIVa").Value))
        
                    
                Case "Documento"
                    Call .ImprimirTextoNoFiscal(rsDetalleHasar.Fields("Detalle").Value)
        
            End Select
            
            rsDetalleHasar.MoveNext
        Loop

        '.DescuentoUltimoItem "Oferta del Dia", 5, True
        '.DescuentoGeneral "Oferta Pago Efectivo", 25, True
        '.EspecificarPercepcionPorIVA "Percep IVA21", 100, 21
        '.EspecificarPercepcionGlobal "Percep. RG 0000", 125#

        'Imprimir Comentarios
       ' Call .ImprimirPago("Efectivo", vMontoEnEF)  'Val(GenerarDato("SELECT SUM(Monto) AS TotalEF FROM Recibo_Temp WHERE IdMedioPago = 1 GROUP BY idMedioPago;", "TotalEF")))
        
        'Call ImprimirComentariosFacturaHasar
        
        
        Select Case EsNulo(rsImprimirHasar.Fields("Tipo").Value)
        
            Case "Fact A", "Ticket-Factura", "Fact B"
            
                Dim vdescuento As Double
                Dim vdescuento_string As String
                
                vdescuento = Val(Me.txtDescuento.Text)
                
                vdescuento_string = " Descuento General del (% " + Me.txtPDescuento.Text + " )"
                
                If vdescuento > 0 Then
                    .DescuentoGeneral vdescuento_string, vdescuento, True
                End If
                
                
                Call .CerrarComprobanteFiscal
                log.AddItem ".CerrarComprobanteFiscal"
            
            Case "Documento"
                Call .CerrarComprobanteNoFiscal
                log.AddItem ".CerrarComprobanteNoFiscal"
                
            Case "Nota C", "Nota D"
                'Call .CerrarComprobanteNoFiscalHomologado
                
                Call .CerrarDNFH
                
                'Call .CerrarComprobanteFiscal
                log.AddItem ".CerrarDNFH"
                
                'log.AddItem ".CerrarComprobanteNoFiscalHomologado"
        
        End Select
        
       ' .Finalizar
    End With
    
    MsgBox "El documento fue enviado a impresión correctamente", vbInformation
    
    
If Err Then
    Call GrabarLog("ImprimirHasar", Str(Err.Number) & " " & Err.Description, Me.Caption)
    Call MsgBox("Error Impresora:" & (Err.Description), vbCritical, "Errores")
Else
    'vImpresionCorrecta = True
End If
End Sub
Private Sub ImprimirComentariosFacturaHasar()
On Error Resume Next

      Dim rsComentariosFactura As New ADODB.Recordset, sqlComentariosFactura As String, l As Integer
    
    sqlComentariosFactura = "SELECT * FROM ComentariosFactura LIMIT 0,4"
    
    'No Tocar esto
    l = 11
    With rsComentariosFactura
        Call .Open(sqlComentariosFactura, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then .MoveFirst
        
        For l = 11 To 14
            If .Fields("Imprimir").Value = "S" Then
                frmPrincipal.FiscalHasar.Encabezado(l) = EsNulo(Left(.Fields("Comentario").Value, 50))
            Else
                frmPrincipal.FiscalHasar.Encabezado(l) = EsNulo(" ")
            End If
            
            .MoveNext
        Next
    
    End With

    sqlComentariosFactura = ""

    If rsComentariosFactura.State = 1 Then
        rsComentariosFactura.Close
        Set rsComentariosFactura = Nothing
    End If
    
    
If Err Then GrabarLog "ImprimirComentariosFacturaHasar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Function controlNroFactura() As Boolean ' controla si el documento fue cargado
On Error Resume Next
Dim c1 As String
Dim sql As String
vfacturaDuplicadaMensaje = ""

sql = "NComprobante=" + Str(Val(Me.txtNroComprobante.Text)) + " and puntodeventa='" + Trim(Me.cboPuntoDeVenta.Text) + "' and letra='" + Trim(Me.cboLetra) + "' and TipoMovimiento='" + TipoDocumento + "'"
c1 = TraerDato("factura", sql, "ncomprobante")
If Not Trim(c1) = "" Then
    controlNroFactura = True
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Cli/Provee :" + TraerDato("factura", sql, "Codigo")
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Fecha :" + TraerDato("factura", sql, "Fecha")
    vfacturaDuplicadaMensaje = vfacturaDuplicadaMensaje + Chr(13) + "> Tipo :" + TraerDato("factura", sql, "TipoMovimiento")
    Me.txtNroInterno = Val(txtNroInterno) - 1
Else
    controlNroFactura = False
End If



If Err Then
    MsgBox "No se puede controlar si este documento fue grabado anteriormente." + Chr(13) + "Consulte con el servicio técnico.", vbCritical
    Exit Function
End If
End Function

Function validarGrabar() As Boolean
Dim vmensaje As String
validarGrabar = True
validarGrabar = True
If Not Val(Me.txtNroComprobante) > 0 Then
    vmensaje = "Val(Me.txtNroComprobante) > 0"
    validarGrabar = False
   ' Grabar = False
End If


If Not Val(Me.cboPuntoDeVenta) > 0 Then
    vmensaje = "Val(Me.cboPuntoDeVenta) > 0"
    validarGrabar = False
   ' Grabar = False
End If

If Val(vtotal_global) - Val(Me.txtSubtotal) > 0.1 Then
    vmensaje = "vtotal_global - txtSubtotal"
    validarGrabar = False
End If

If Me.vcodEmpresa.Text = "" And LeerXml("Puesto") = "EMPRESAS" Then
    
    validarGrabar = False
    
    vmensaje = "Debe seleccionar una empresa. Tenés el xml setrado como empresa"

End If



If Not validarGrabar Then

    MsgBox "Verificar el nro de comprobante y el punto de venta" + Chr(13) + vmensaje, vbCritical
    validarGrabar = False
   ' Grabar = False
    
End If
End Function

Function GuardarCompleto() As Boolean
'Private Sub GuardarCompleto()
On Error Resume Next
 
 
GuardarCompleto = False

If Not validarGrabar Then
    GuardarCompleto = True
    Exit Function
End If

 If Me.vGrabaModo = 0 Then
 
        ' ------------ verifica nro interno ----------------------
        'If existeRegistro(Val(Me.txtNroInterno)) Then Exit Function
        '----------------------------------------------------------
        
        Do Until Not existeRegistro(Val(Me.txtNroInterno))
            txtNroInterno.Text = UltimoNroInterno2
        Loop
        
        '--------
        
 
 End If
 
 ' ----- pongos las componentes de pago visibles --------------
 frmRemito.pbCarga(3).Enabled = True
 frmRemito.txtCaja(2).Enabled = True
 frmRemito.txtCaja(3).Enabled = True
 '-------------------------------------------------------------
                

    If controlNroFactura And vGrabaModo = 0 Then
        MsgBox "Este documento ya fue grabado anteriormente." + Chr(13) + "Datos del documento guardado: " + Chr(13) + vfacturaDuplicadaMensaje, vbCritical, "Documento duplicado..."
        Exit Function
    End If

    Dim vPeriodoFactura As String, codClienteCobro As String, remi As Long
    
    vPeriodoFactura = Year(dtpFecha.Value) & Mid(dtpFecha.Value, 4, 2)
    
    If Val(txtNroInterno.Text) = 0 And (opOtrosDocumentos.Value = True) And vDatosEmpresa.UsarNroInterno = "SI" Then
        MsgBox "Debe Ingresar un Nro Interno para el Documento!!", vbExclamation, "Mensaje ..."
        txtNroInterno.SetFocus
        Exit Function
    End If
    
    If vPeriodoFactura = TraerDato("IvaVentaCerrado", "Periodo = '" & vPeriodoFactura & "'", "Periodo") Then
        MsgBox "La Factura pertenece a un periodo de Iva Venta Ya Cerrado!!!", vbInformation, "Mensaje ..."
        Exit Function
    End If
    
    With txtEmpleados(0)
        If ConfigRemito(4) = True And Trim(.Text) = "" Then
            .BackColor = vbRed
            .SetFocus
            Exit Function
        End If
    End With
    
    With cbolista
        If opOtrosDocumentos.Value = False Then
            If (ConfigRemito(6) = False) And (Val(cbolista.Text) = 0) And (vGrabaModo = 0) Then
                .BackColor = vbRed
                .SetFocus
                
              '  MsgBox "No se puede realizar la operación. " + Chr(13) + "  If (ConfigRemito(6) = False) And (Val(cbolista.Text) = 0) And (vGrabaModo = 0) Then"
                'Exit Sub
            End If
        End If
    End With

    If Trim(cboLetra.Text) = "" Or Trim(cboPuntoDeVenta.Text) = "" Or Val(txtNroComprobante.Text) = 0 Then
        MsgBox "Debe definir los datos Fiscales de este comprobante !!!", vbExclamation, "Mensaje ..."
        cboLetra.SetFocus
        Exit Function
    End If
    
    If Trim(txtCliente(0).Text) = "" Then
        MsgBox "Tiene campos obligatorios vacios, complete la factura y vuelva a intentarlo", vbInformation, "Mensaje"
        txtCliente(0).SetFocus
        Exit Function
    End If
    
    If opOtrosDocumentos.Value = False Then
        If Not opTipoDoc(2).Value = True Then
            
            codClienteCobro = EsNulo(txtCliente(0).Tag)
            remi = Val(txtNroRemito.Text)
        
            'If vConfigGral.vIncluyeCobros = True Then
             If 1 = 2 Then
                
                With frmCobros
                    .txtNroComprobante.Text = EsNulo(txtNroComprobante.Text)
                    .NroComprobante = Val(txtNroComprobante.Text)
                    .txtTipoComp.Text = TipoDocumento
                    .tipoComprobante = TipoDocumento
                    .total = Val(txtTotal.Text)
                    .pendiente = Val(txtTotal.Text)
                    .remito = Val(txtNroRemito.Text)
                    .fechaDocumento = Me.dtpFecha.Value
                    .esComprobanteAutomatico = False
                    .esFacturacion = True
                    .codCliente = txtCliente(0).Tag
    
                    
                    GuardarDoc 'Guardo la Factura y el Detalle
                
                    Call .BuscarDatosOperacionesCliente(codClienteCobro, remi)
    
                    .txtImporteEfectivoPesos.SetFocus
                    .Height = 9465
                    .Width = 9285
            
                    Load frmCobros
    
                    .HabilitarControles (True)
                End With
            
                lblsaldocliente.Caption = "0.000"
            Else
                'NOOOO Es una Nota C y No esta habilitado COBROS
                GuardarDoc
            End If
        
        Else
            'Es una Nota C y No esta habilitado COBROS
            GuardarDoc
        End If
    Else
        
        Select Case UCase(txtTipoMovimiento(0).Text)
        
            Case "CD"
                GBCaja.Visible = True
                txtCaja(7).Text = Trim(txtIB(10).Text)
                txtCaja(8).Text = Trim(txtIB(7).Text)
                txtCaja(0).SetFocus
                
                       ' desactivo componentes nocontado
                   
                        frmRemito.pbCarga(3).Enabled = False
                        frmRemito.txtCaja(4).Text = "EF"
                        frmRemito.txtCaja(5).Text = "EFECTIVO"
                        frmRemito.txtCaja(2).Enabled = False
                        frmRemito.txtCaja(3).Enabled = False
                           
                
            
            Case Else
                GBCaja.Visible = False
        
                vLeyendaAsiento = Trim(txtIB(7).Text)
                If Val(txtIB(10).Text) < 0 Then
                    vTotalAsiento = Val(Format(txtIB(10).Text * (-1), "#####0.00"))
                Else
                    vTotalAsiento = Format(Val(txtIB(10).Text), "#####0.00")
                End If
        
                GuardarDoc
                Me.txtNroInterno = Val(txtNroInterno) + 1
        
        End Select

    End If
    
    If Not LeerConfig(17) = "Otros" Then BorrarArticulosNoGuardados

If Err > 0 Then
    GrabarLog "GuardarCompleto (" & 1 & ")", Err.Number & " " & Err.Description, Me.Name
   ' MsgBox "El comprobante de venta ha sido guardado" + Chr(13) + "Error: " + Err.Description, vbInformation, "Error al guardar"
Else
    
    If Not txtTipoMovimiento(0).Text = "CD" Then
        'MsgBox "El comprobante de venta ha sido guardado" + Chr(13) + "Ahora debe ingresar el asiento correspondiente", vbInformation, "Guardando documento de venta .."
    End If
End If

End Function
Private Sub BorrarArticulosNoGuardados()
    On Error Resume Next

    MousePointer = vbHourglass
    
    Call BorrarBase("Articulos WHERE Observaciones = 'CargadoPorRemito'", pathDBMySQL)

    MousePointer = vbDefault

    If Err Then GrabarLog "BorrarArticulosNoGuardados", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub BuscarArticulo()
    On Error Resume Next

    MousePointer = vbHourglass
    barticulo.Refresh
    
    frmBuscarArticulo.o = 1
    frmBuscarArticulo.vlista = Val(cbolista.Text)

    frmBuscarArticulo.txtArticulo.SetFocus
    frmBuscarArticulo.Visible = True

    If frmBuscarArticulo.busca = 2 Then
        frmBuscarArticulo.txtArticulo = vvvcodigo
        ' frmBuscarArticulo.txtArticulo_KeyPress (10)
    Else
        frmBuscarArticulo.txtArticulo.Text = vvvdescrip
        ' frmBuscarArticulo.txtArticulo_KeyPress (13)
    
    End If

    frmBuscarArticulo.txtArticulo.SetFocus
    MousePointer = vbDefault

    If Err Then GrabarLog "BuscarArticulo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function BuscarCliente() As Boolean
    On Error Resume Next
    
    Dim rsCliente As New ADODB.Recordset, sqlCliente As String

    sqlCliente = "SELECT * FROM Clientes WHERE ((Nombre = '" & Trim(txtCliente(0).Text) & "') OR (Codigo = '" & Trim(txtCliente(0).Text) & "'))"
    
    With rsCliente
        .CursorLocation = adUseClient
        Call .Open(sqlCliente, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then

            Select Case ValidarCliente(.Fields("Codigo").Value)
            
                Case "CreditoMax"
                    Call Habilitar(Not True)
                    Exit Function
                
                Case "Estado"
                    'MsgBox "El Estado del Cliente No permite que pueda Facturarle!!", vbExclamation, "Mensaje ..." ' panic
                    Call Habilitar(Not True)
                    Exit Function
                
                Case Else
                
            End Select
            
            BuscarCliente = True
            txtCliente(0).Tag = EsNulo(.Fields("Codigo").Value)
            txtCliente(0).Text = EsNulo(.Fields("Nombre").Value)
            txtCliente(1).Text = EsNulo(.Fields("Direccion").Value)
            txtCliente(2).Text = EsNulo(.Fields("Localidad").Value)
            txtCliente(3).Text = EsNulo(.Fields("Cuit").Value)
            Me.cbolista = EsNulo(.Fields("idListas").Value)

            cboTipoIva.Tag = EsNulo(.Fields("idTipoIva").Value)
            cboTipoIva.Text = TraerDato("TipoIva", "idTipoIva =  '" & (EsNulo(.Fields("idTipoIva").Value)) & "'", "TipoIva")
            
            If Not .Fields("u_pago").Value = "" Then
                gupago = .Fields("u_pago").Value
            Else
                gupago = "No encontrado"
            End If

            If Not EsNulo(.Fields("u_venta").Value) = "" Then
                guventa = .Fields("u_venta").Value
            Else
                guventa = "No encontrado"
            End If

            gsaldo = Format(.Fields("saldo").Value, "#######0.000")
            gcredito = Format(.Fields("creditoMax").Value, "#######0.000")

            cbolista.Text = Val(.Fields("idListas").Value)
        
            'Ver info de cliente
            'If ConfigRemito(0) = True Then frmClienteInfo.foco
                       
            'Cargar numero de comprobante al inicio
            'If ConfigRemito(1) = True Then NroComprobante

            Call Habilitar(True)
        
            'txtNroRemito.Text = NroRemito ' busca el nro deremito
            
            
            If Me.checkSaldo Then lblsaldocliente.Caption = SaldoCliente
        
        
            If Not LeerConfig(17) = "Otros" Then
               ' opTipoDoc(0).Value = True
               ' opTipoDoc_Click (0)
            Else
                opOtrosDocumentos.Value = True
                opOtrosDocumentos_Click
                txtNroInterno.SetFocus
            End If
            'BuscarCliente (1)
            
        Else
            
            'vClienteNuevo = True
            'txtCliente(0).Tag = Val(GenerarDato("SELECT MAX(Codigo) AS UltimoCodigo FROM Clientes", "UltimoCodigo")) + 1
            'txtCliente(0).Tag = FormatoUltimoCodigo(4, txtCliente(0).Tag)
            'Habilitar (True)
            'txtCliente(1).SetFocus
            BuscarCliente = False
        End If
        
    End With

    sqlCliente = ""
    
    If rsCliente.State = 1 Then
        rsCliente.Close
        Set rsCliente = Nothing
    End If
    
If Err Then GrabarLog "BuscarCliente", Err.Number & " " & Err.Description, Me.Name
End Function
Private Function ValidarCliente(vCodigoCliente As String) As String
On Error Resume Next

    'ValidarCliente = vClienteNuevo
    
    'If vClienteNuevo = True Then Exit Function
    
    If vCodigoCliente = "" Then
        ValidarCliente = Not True
        MsgBox "Debe ingresar un cliente !!!!", vbExclamation, "Mensaje ..."
        Exit Function
    End If

    Dim vSaldoCliente As Double, vCreditoMax As Double, i As Integer
    
    ValidarCliente = ""
    
   ' vSaldoCliente = Format(TraerDato("SaldoClientesSimple", "Codigo = '" & Trim(vCodigoCliente) & "'", "SaldoCliente"), "#######0.000")
    vCreditoMax = Val(TraerDato("Clientes", "Codigo = '" & Trim(vCodigoCliente) & "'", "CreditoMax"))

    'Controlo que El Estado lo deje facturar

  Exit Function
    If (TraerDato("Estados", "idEstados = '" & EsNulo(rsClientes.Fields("idEstados").Value) & "'", "SePuedeFacturar") = "N") Then


  
        
        For i = 0 To txtCliente.Count - 1
            txtCliente(i).Text = ""
            txtCliente(i).Tag = ""
        Next
        
        cboTipoIva.Text = ""
        cboTipoIva.Tag = ""
        
        ValidarCliente = "Estado"
        
        Exit Function
    End If

    If (vSaldoCliente > vCreditoMax) And Not (vCreditoMax = 0) Then
        If Not MsgBox("El Saldo Actual del Cliente Supera el Limite de Crédito ¿ Permitir Movimiento de todas maneras ?", vbExclamation + vbYesNo, "Mensaje ...") = vbYes Then
                        
            ValidarCliente = "CreditoMax"
            
            For i = 0 To txtCliente.Count - 1
                txtCliente(i).Text = ""
                txtCliente(i).Tag = ""
            Next

            Exit Function
        End If
                    
    End If
        
If Err Then GrabarLog "ValidarCliente", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub BuscaDoc()
    On Error Resume Next
    
    frmBuscarFactura.Show

    If Err Then GrabarLog "BuscaDoc", Err.Number & " " & Err.Description, Me.Name
End Sub
Function CalcularIva(vTipoIva As String, vvtotal As Double) As Double
    On Error Resume Next
    Dim ivatotal As Double
    

    Select Case Trim(cboTipoIva.Text)

        Case "Responsable Inscripto", "Resp.Inscripto"
            ivatotal = vvtotal * Val(vTipoIva) / 100

        Case "Fact B"
            ivatotal = vvtotal * Val(vTipoIva) / 100

        Case "Consumidor Final"
            ivatotal = vvtotal * Val(vTipoIva) / 100
    End Select

    ivatotal = vvtotal * Val(vTipoIva) / 100
    CalcularIva = ivatotal

    If Err Then GrabarLog "caliva", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub CalcularTotales()
    On Error Resume Next
    Dim iva21, iva105, iva27 As Integer
    Dim vTotalParcial, vTotal105, vTotal210, vTotal270, vdescuento, vImpuesto, vnegro, vIva As Double
    Dim i As Integer
    
    LimpiarTotales

    vTotalParcial = 0
    vTotal105 = 0
    vTotal210 = 0
    vTotal270 = 0
   ' vdescuento = 0
    vnegro = 0

    vLeyendaAsiento = ""
    
    
    With KlexDetalle
        
       ' If chkTotalManual.Value = 0 Then
       Dim vhasta As Integer
       
       vhasta = .Rows
            
            For i = 1 To vhasta - 1
                
                
                
                If Not .TextMatrix(i, 5) = "0.000" Then
                    
                Else
                    .RemoveItem (i)
                    i = i - 1
                    vhasta = vhasta - 1
                End If
                
                If Not Trim(.TextMatrix(i, 5)) = "" Then
                    
                   ' Call SeleccionarColor(.TextMatrix(i, 25), i)
                    
                    
                    vtotal_global = 0
                    
                    vTotalParcial = vTotalParcial + Val(KlexDetalle.TextMatrix(i, 11))

                    If opTipoDoc(5).Value = True Or opTipoDoc(0).Value = True Or opTipoDoc(2).Value = True Or opTipoDoc(4).Value = True Or opTipoDoc(1).Value = True Then      ' tipos de documentos (factura, reimito, documentos) ' todo
                        
                        'If (Trim(cboTipoIva.Text) = "Iva Responsable Inscripto" Or Trim(cboTipoIva.Text) = "Iva Resp.Inscripto") And vDatosEmpresa.CondicionIva = "Responsable Inscripto" Then ' hago este control porque no tengo id para el tipo de iva (cod:id-TipoIVA) (cod:TipoIva-Empresa)
                            
                        If True Then  ' hago este control porque no tengo id para el tipo de iva (cod:id-TipoIVA) (cod:TipoIva-Empresa)
                            
                            'iva21 = 0
                            'iva105 = 0
                            'iva27 = 0
                            
                            
                            If Val(.TextMatrix(i, 9)) = 10.5 Then vTotal105 = vTotal105 + Val(.TextMatrix(i, 11))   ' * 0.105)
                            If Val(.TextMatrix(i, 9)) = 21 Then vTotal210 = vTotal210 + Val(.TextMatrix(i, 11)) '  * 0.21)
                            If Val(.TextMatrix(i, 9)) = 27 Then vTotal270 = vTotal270 + Val(.TextMatrix(i, 11)) ' * 0.27)
                            
                            txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtImpuesto.Text)
                        Else
                            
                            iva21 = 0
                            iva105 = 0
                            iva27 = 0
                            
                        End If
                    Else
                        
                            vTotal105 = 0
                            vTotal210 = 0
                            vTotal270 = 0
                        
                    End If
                
                    'If Val(.Recordset("Tiva").Value) = 0 Then vnegro = vnegro + .Recordset("total").Value
    
                    vLeyendaAsiento = vLeyendaAsiento & Trim(.TextMatrix(i, 6)) & " - "
                    
                End If
            Next
            
        
        'Else
        '    txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtImpuesto.Text)
        '   vTotalParcial = Val(txtTotal.Text)
        'End If
        
    End With
    
    vLeyendaAsiento = Trim(vLeyendaAsiento)
    vLeyendaAsiento = Trim(Mid(vLeyendaAsiento, 1, Len(vLeyendaAsiento) - 1))
    
    vdescuento = Val(txtPDescuento)
    
    'vTotalParcial = descuento(vTotalParcial, vdescuento) ' le aplico el descuento a todo el subtotal
    vtotal_global = vTotalParcial
    
    vnegro = descuento(vnegro, vdescuento)
    
    If opTipoDoc(5).Value = True Or opTipoDoc(0).Value = True Or opTipoDoc(0).Value = True Or opTipoDoc(2).Value = True Or opTipoDoc(4).Value = True Or opTipoDoc(1).Value = True Then
        If Trim(cboTipoIva.Tag) = "001" Then
            
            txtSubtotal.Text = vTotalParcial
    
            ' le aplico el descuento a cada uno de los ivas
            txtIva(0).Text = descuento(vTotal105, vdescuento) * 0.105
            txtIva(1).Text = descuento(vTotal210, vdescuento) * 0.21
            txtIva(2).Text = descuento(vTotal270, vdescuento) * 0.27
    
        Else
            txtSubtotal.Text = vTotalParcial + Val(Me.txtIva(0).Text) + Val(Me.txtIva(1).Text) + Val(Me.txtIva(2).Text) + vnegro
        End If
    Else
        txtSubtotal.Text = vTotalParcial + Val(Me.txtIva(0).Text) + Val(Me.txtIva(1).Text) + Val(Me.txtIva(2).Text) + vnegro
    End If
    
    vtotal_real = Val(txtSubtotal.Text + Val(Me.txtIva(0).Text) + Val(Me.txtIva(1).Text) + Val(Me.txtIva(2).Text) + vnegro)
    
    vdescuento = Str((vTotalParcial + vnegro) * Val(txtPDescuento.Text) / 100)
    txtDescuento.Text = vdescuento
    
    'vTotalParcial = vTotalParcial - vdescuento + txtImpuesto

    vImpuesto = Val(txtImpuesto.Text) * (vTotalParcial + vnegro) / 100
    
    If opTipoDoc(5).Value = True Or opTipoDoc(0).Value = True Or opTipoDoc(2).Value = True Or opTipoDoc(1).Value = True Then
        
        If Trim(cboTipoIva.Tag) = "001" Then
            txtTotal.Text = vtotal_real + Val(vImpuesto) - Val(txtDescuento.Text)
        Else
            txtTotal.Text = Val(vTotalParcial) + Val(vnegro) + Val(vImpuesto) - Val(txtDescuento.Text)
        
        End If
    Else
        
        txtTotal.Text = Val(vTotalParcial) + Val(vnegro) + Val(vImpuesto) - Val(txtDescuento.Text)
    
    End If
    
    vTotalAsiento = Val(txtTotal.Text) + Val(txtIB(10).Text)
    

' pasa totales a variables globales
vgTsubtotal = Me.txtSubtotal
vgTiva105 = Me.txtIva(0)
vgTiva21 = Me.txtIva(1)
vgTiva27 = Me.txtIva(2)
vgTPdescuento = Me.txtPDescuento
vgTimpuesto = Me.txtImpuesto
vgTtotal = Me.txtTotal
    
    'DecorarTalles
    
    If Err Then GrabarLog "CalcularTotales", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function Maximo(n1 As Double, n2 As Double) As Double

    If (n1 > n2) Then
        Maximo = n1
    Else
        Maximo = n2
    End If

End Function
Private Sub SeleccionarColor(vColor As String, vFila As Integer)
On Error Resume Next

    With KlexDetalle
        .Col = 25
        .Row = vFila
        
        Select Case Left(vColor, 1)
                    
            Case "N"
                .CellBackColor = vbRed
                
            Case "B"
                .CellBackColor = vbGreen
                
            Case ""
                .CellBackColor = vbWhite
                    
        End Select

    End With

If Err Then GrabarLog "SeleccionarColor", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarComentario_OLD()
On Error Resume Next
    
    If Trim(txtCliente(0).Tag) = "" Then
        MsgBox "Debe cargar un cliente previamente", vbExclamation, "Mensaje..."
        Exit Sub
    End If
    
    'With frmComentario
    '    .txtCliente.Text = Trim(txtCliente(0).Tag)
    '    .txtCliente_Keypress 13
    '    .TabComentarios.Tab = 0
    'End With

If Err Then GrabarLog "CargarComentario", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub ElegirTipoPrecio()
    On Error Resume Next
        
    If Trim(tprecio.Text) = "Preguntar" Then
        If MsgBox("¿ Desea seleccionar el precio en Pesos ($) ?", vbYesNo) = vbYes Then
            txtDetalle(2).Text = Val(txtDetalle(2).Text)
            tprecio.Text = "Pesos ($)"
        Else
            txtDetalle(2).Text = inulo(vvpdolar) * gdolar
            tprecio.Text = "Dolar (u$s)"
        End If
            
    End If
        
    If Trim(tprecio.Text) = "Dolar (u$s)" Then txtDetalle(2) = vvpdolar
    If Trim(tprecio.Text) = "Pesos ($)" Then txtDetalle(2).Text = Val(txtDetalle(2).Text)
    
    If Err Then GrabarLog "ElegirTipoMoneda", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function CargarFDetalle(vRemitoDetalle As Long) As Boolean
    On Error Resume Next

    Dim rsCargaFDetalle As New ADODB.Recordset, sqlCargaFDetalle As String
    Dim j As Integer
    
   Call EjecutarScript("delete from fdetalle where remito = " & vRemitoDetalle + " and codigo is null", pathDBMySQL)
    
    sqlCargaFDetalle = "SELECT * FROM FDetalle WHERE (remito = " & vRemitoDetalle & ") and cantidad > 0 ORDER BY idFDetalle ASC"
    
    With rsCargaFDetalle
        .CursorLocation = adUseClient
        
        Call .Open(sqlCargaFDetalle, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .RecordCount > 0 Then
        
        Else
        
            Exit Function
        End If
        
        FormatoGrillaDetalle (.RecordCount)
            
        If Not .EOF = True Then
            .MoveFirst
            
            j = 1
            
            
            Call Json.initDic(arr2)
            
            Do Until .EOF = True
                                
                                
                                
                               Call Json.setDic(EsNulo(.Fields("codigo").Value), EsNulo(.Fields("Cantidad").Value), arr2())
                               
                                
                                KlexDetalle.TextMatrix(j, 1) = EsNulo(.Fields("idFDetalle").Value)
                                KlexDetalle.TextMatrix(j, 2) = EsNulo(.Fields("Fecha").Value)
                                KlexDetalle.TextMatrix(j, 3) = EsNulo(.Fields("Remito").Value)
                                KlexDetalle.TextMatrix(j, 4) = "[" + .Fields("codigo").Value + "]"
                                KlexDetalle.TextMatrix(j, 5) = EsNulo(.Fields("Cantidad").Value)
                                KlexDetalle.TextMatrix(j, 6) = EsNulo(.Fields("Detalle").Value)
                                KlexDetalle.TextMatrix(j, 9) = EsNulo(.Fields("tiva").Value)
                
                                If opTipoDoc(1).Value = True Or Trim(cboTipoIva.Text) = "Consumidor Final" Or Trim(cboTipoIva.Text) = "Responsable Monotributo" Then
                                    
                                    KlexDetalle.TextMatrix(j, 7) = Format(Val(.Fields("Precio").Value), "######0.000") '(precio)
                                    'KlexDetalle.TextMatrix(j, 9) = "0"
                                
                                Else
                                    
                                    If bienes.Value = 1 Then
                                        KlexDetalle.TextMatrix(j, 7) = Format(Val(txtDetalle(2).Text) - (Val(txtDetalle(2).Text) * 9.5 / 100), "######0.000") ' (precio)
                                        KlexDetalle.TextMatrix(j, 11) = "10.5" '(tiva)
                                    Else
                                        KlexDetalle.TextMatrix(j, 7) = Format(.Fields("Precio").Value, "######0.000")
                                        KlexDetalle.TextMatrix(j, 8) = ""
                                       ' KlexDetalle.TextMatrix(j, 9) = EsNulo(TraerDato("Articulos", "Codigo = '" & .Fields("codigo").Value & "'", "TipoIva"))
                                        
                                        If KlexDetalle.TextMatrix(j, 9) = "" Then KlexDetalle.TextMatrix(j, 9) = 21
                                        
                                    End If
                                
                                End If
                                KlexDetalle.TextMatrix(j, 8) = EsNulo(.Fields("Descuento").Value)
                                KlexDetalle.TextMatrix(j, 10) = EsNulo(.Fields("Impuesto").Value)
                                
                                KlexDetalle.TextMatrix(j, 11) = EsNulo(.Fields("Total").Value)
                    
                    
                                'KlexDetalle.TextMatrix(j, 13) = EsNulo(.Fields("Envase").value)
                                'KlexDetalle.TextMatrix(j, 15) = EsNulo(.Fields("Pago").value)
                                'KlexDetalle.TextMatrix(j, 16) = EsNulo(.Fields("Resta").value)
                
                                'KlexDetalle.TextMatrix(j, 17) = EsNulo(.Fields("TotalIva").value)    'Totaliva
                                'KlexDetalle.TextMatrix(j, 18) = EsNulo(.Fields("ganancia").value)    'Ganancia
                                'KlexDetalle.TextMatrix(j, 19) = EsNulo(.Fields("Sueldo").value)      'Sueldo
                                'KlexDetalle.TextMatrix(j, 20) = EsNulo(.Fields("repartidor").value)  'Repartidor
                                'KlexDetalle.TextMatrix(j, 21) = EsNulo(.Fields("Confirmado").value)
                                
                                'KlexDetalle.TextMatrix(j, 22) = EsNulo(.Fields("IdFDetalle").value)
                                
                                KlexDetalle.Row = KlexDetalle.Row + 1
                        
                                vRemitoControl = Val(.Fields("Remito").Value)
                                vCantidadControl = vCantidadControl + 1
                                .MoveNext
                
                                j = j + 1
                
            Loop

            CalcularTotales  ' Diagrama de módulos

            CargarFDetalle = True
        Else
        
            CargarFDetalle = Not True
        
        End If
            
    End With
    
    Call LastKlexRow(Me.KlexDetalle)
    
    sqlCargaFDetalle = ""
    
    If rsCargaFDetalle.State = 1 Then
        rsCargaFDetalle.Close
        Set rsCargaFDetalle = Nothing
    End If
    
    If Err Then GrabarLog "CargarFDetalle (" & vRemitoDetalle & ")", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub CargarFactura(vRemitoDetalle As Long)
    On Error Resume Next

    With bfactura

        txtCliente(0).Tag = EsNulo(.Recordset("Codigo").Value)
        txtCliente(0).Text = EsNulo(.Recordset("Nombre").Value)
        txtCliente(1).Text = EsNulo(.Recordset("Domicilio").Value)
        txtCliente(2).Text = EsNulo(.Recordset("Localidad").Value)
        'txtCliente(3).Text = EsNulo(.Recordset("Telefono").Value)
        txtCliente(3).Text = EsNulo(.Recordset("Cuit").Value)
                
        
        cboTipoIva.Text = EsNulo(.Recordset("Iva").Value)
        cboTipoIva.Tag = TraerDato("Tipoiva", "tipoiva = '" & EsNulo(.Recordset("Iva").Value) & "'", "idTipoIva")
        
        txtNroInterno.Text = EsNulo(.Recordset("NroInterno").Value)
        
        Me.vnroremito2 = EsNulo(.Recordset("nroremito").Value)
        Me.vcventa = EsNulo(.Recordset("cventa").Value)
        
        
        dtpFecha.Value = EsNulo(.Recordset("fecha").Value)
        vFechaIva.Value = EsNulo(.Recordset("fechaIVA").Value)
        
        txtNroRemito.Text = Val(.Recordset("remito").Value)
        txtSubtotal.Text = Format(.Recordset("subtotal").Value, "###########0.000")
        txtIva(1).Text = Format(.Recordset("Tiva").Value, "###########0.000")
        txtIva(0).Text = Format(.Recordset("tiva2").Value, "###########0.000")
        txtTotal.Text = Format(.Recordset("total").Value, "###########0.000")
        txtDescuento.Text = Format(.Recordset("descuento").Value, "###########0.000")
        txtImpuesto.Text = Format(.Recordset("Impuesto").Value, "###########0.000")
        
        CargarTipoDocumento (.Recordset("tipo").Value)
          
        cboLetra.Text = EsNulo(.Recordset("LeNComprobantetra").Value)
        cboPuntoDeVenta.Text = EsNulo(.Recordset("PuntoDeVenta").Value)
        txtNroComprobante.Text = EsNulo(.Recordset("NComprobante").Value)
        
        vqrnombre = Replace(Trim(txtCliente(3).Text), "-", "") + Trim(txtNroComprobante.Text)
        
        vnrocomprobante_control1 = txtNroComprobante.Text
        
        txt_vcantidadVolquete = EsNulo(.Recordset("cantidadvolquetes").Value)
        
        ' ----- cargar datos de remito
        Me.vRemitoRecibio = EsNulo(.Recordset("Recibio").Value)
        Me.vTransportistaNombre = .Recordset("TransportistaNombre").Value
        Me.vTransportistaCuit = .Recordset("TransportistaCuit").Value
        Me.vTransportistaDomicilio = .Recordset("TransportistaDomicilio").Value
        Me.vlentrega = .Recordset("lentrega").Value
        Me.vobservacion = .Recordset("comentario").Value
        
        lblsaldocliente.Caption = Str(.Recordset("saldos").Value)
        
        Me.vcae.Text = .Recordset("cae").Value
        Me.vcaeFecha2 = .Recordset("caevto").Value
        vcaeFecha = .Recordset("caevto").Value

        '------------------------------=
        
        
        
        Me.vdescEmpresa.Text = getvdesEmpresa(.Recordset("idfactura"))
        
        Me.vDesRepartidor.Text = getvdesRepartidor(.Recordset("idfactura"))
        Me.vcodEmpresa.Text = getvcodEmpresa(.Recordset("idfactura"))
        Me.vcodRepartidor.Text = getvcodRepartidor(.Recordset("idfactura"))
    End With
    
    If Err Then GrabarLog "CargarFactura (" & vRemitoDetalle & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub CargarRemito(vRemitoModif As Long)
    On Error Resume Next
    
    With bfactura
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Factura WHERE (remito = " & vRemitoModif & ") order by nrointerno desc limit 1"
        .Refresh

        If .Recordset.RecordCount > 1 Then
            MsgBox "Cuidado. Hay un problema al querere cargar este documento. Consulte servicio técnico"
            Exit Sub
        End If
        
        If Not .Recordset.EOF Then
            
            CargarFactura (vRemitoModif)
            CargarFDetalle (vRemitoModif)
            Me.CargarChoferAGrilla (Me.vlistachofer)
            
            Me.chkReingresarFact.Value = xtpChecked
            
            'vncomprobante = .Recordset("NComprobante").value
            'txtNroComprobante.Text = vncomprobante
            
            'cboLetra.Text = EsNulo(.Recordset("Letra").value)
            'cboPuntoDeVenta.Text = EsNulo(.Recordset("PuntoDeVenta").value)
        
        End If
        
    End With
    
    If Err Then GrabarLog "CargarRemito (" & vRemitoModif & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub HabilitarDocAFactura(vHabilita As Boolean)
On Error Resume Next

    vHabilitaDocAFactura = vHabilita
    
    PBDocAFactura(0).Enabled = vHabilita
    PBDocAFactura(1).Enabled = vHabilita
    
    RBIva(0).Enabled = vHabilita
    RBIva(1).Enabled = vHabilita
    
    chkBorrarDocOriginal.Enabled = vHabilita
    
    GBDocAFactura(0).Enabled = vHabilita
    GBDocAFactura(1).Enabled = vHabilita

If Err Then GrabarLog "HabilitarDocAFactura", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarTipoDocumento(vtipo)
    On Error Resume Next

    Select Case vtipo

        Case "Fact A"
            opTipoDoc(0).Value = True

        Case "Fact B"
            opTipoDoc(0).Value = True

        Case "Presupuesto"
            opTipoDoc(1).Value = True

        Case "Nota C"
            opTipoDoc(2).Value = True

        Case "Documento"
            opTipoDoc(3).Value = True

        Case "Remito"
            opTipoDoc(4).Value = True
    
        Case "Nota D"
            opTipoDoc(5).Value = True
            
    End Select

    If Err Then GrabarLog "cargartipodoc (" & vtipo & ")", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub btn_CargarChoferes_Click(Index As Integer)
' completar fSeleccionChoferes
'Call fbuscarGrilla("empleados", "Nombre", "idEmpleados", Me.vchofer.Name, Me)
Call fbuscarGrilla("proveedores", "Nombre", "idEmpleados", Me.vchofer.Name, Me)
End Sub

Private Sub btn_pasar_Click()
' ema:
Me.grd_Choferes.Cols = 3
Me.grd_Choferes.Rows = Me.grd_Choferes.Rows + 1
Me.grd_Choferes.TextMatrix(Me.grd_Choferes.Rows - 1, 1) = vchofer.Tag
Me.grd_Choferes.TextMatrix(Me.grd_Choferes.Rows - 1, 2) = vchofer.Text
End Sub

Private Sub cboBienesServicios_GotFocus()
On Error Resume Next

    With cboBienesServicios
        .Clear
        .AddItem ("Bienes")
        .AddItem ("Servicios")
    End With

If Err Then GrabarLog "cboBienesServicios_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboLetra_Click()
On Error Resume Next

    If opOtrosDocumentos.Value = True Then
        If Not Trim(cboPuntoDeVenta.Text) = "" And Not Trim(cboLetra.Text) = "" Then
    
            txtNroComprobante.Text = Val(EsNulo(GenerarDato("SELECT MAX(NComprobante) AS NComp, TipoMovimiento FROM Factura WHERE (Letra = '" & Trim(cboLetra.Text) & "') AND (PuntoDeVenta = '" & Trim(cboPuntoDeVenta.Text) & "') AND (TipoMovimiento = 'FC')  GROUP BY TipoMovimiento", "NComp"))) + 1
                
        Else
            'No Puede traer el Ultimo Codigo
        End If
    Else
        If Not Trim(cboPuntoDeVenta.Text) = "" And Not Trim(cboLetra.Text) = "" Then
            txtNroComprobante.Text = Val(EsNulo(GenerarDato("SELECT MAX(NComprobante) AS NComp FROM Factura WHERE (Tipo = '" & TipoDocumento & "') AND (Letra = '" & Trim(cboLetra.Text) & "') AND (PuntoDeVenta = '" & Trim(cboPuntoDeVenta.Text) & "')", "NComp"))) + 1
        
        
        Else
            'No Puede traer el Ultimo Codigo
        End If
    End If

If Err Then GrabarLog "cboLetra_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cboLetra_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = 13 Then cboPuntoDeVenta.SetFocus
    
    If Err Then GrabarLog "cboLetra_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub



Private Sub cboLetra_LostFocus()

vncomprobante = NroComprobanteNuevo(TipoDocumento, Trim(cboLetra.Text), Trim(cboPuntoDeVenta.Text))
Me.txtNroComprobante = vncomprobante

End Sub

Private Sub cboLista_GotFocus()
    On Error Resume Next
    
    Call CargarCombo("Listas", "Lista", cbolista, False)
   ' cbolista.Text = 1
    
    If Err Then GrabarLog "cboLista_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboPuntoDeVenta_Click2()
On Error Resume Next

    If opOtrosDocumentos.Value = True Then
        If Not Trim(cboPuntoDeVenta.Text) = "" And Not Trim(cboLetra.Text) = "" Then
    
            txtNroComprobante.Text = Val(EsNulo(GenerarDato("SELECT MAX(NComprobante) AS NComp, TipoMovimiento FROM Factura WHERE (Letra = '" & Trim(cboLetra.Text) & "') AND (PuntoDeVenta = '" & Trim(cboPuntoDeVenta.Text) & "') AND (TipoMovimiento = 'FC')  GROUP BY TipoMovimiento;", "NComp"))) + 1
            
            'txtNroComprobante.Text = frmPrincipal.FiscalHasar.UltimaFactura(1)

                
        Else
            'No Puede traer el Ultimo Codigo
        End If
    Else
        If Not Trim(cboPuntoDeVenta.Text) = "" And Not Trim(cboLetra.Text) = "" Then
    
            txtNroComprobante.Text = Val(EsNulo(GenerarDato("SELECT MAX(NComprobante) AS NComp FROM Factura WHERE (Tipo = '" & TipoDocumento & "') AND (Letra = '" & Trim(cboLetra.Text) & "') AND (PuntoDeVenta = '" & Trim(cboPuntoDeVenta.Text) & "')", "NComp"))) + 1
                
        Else
            'No Puede traer el Ultimo Codigo
        End If
    End If
    
If Err Then GrabarLog "cboPuntoDeVenta_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboPuntoDeVenta_GotFocus()
On Error Resume Next
Exit Sub
    With cboPuntoDeVenta
        .Clear
        .AddItem ("0001")
        .AddItem ("0002")
        .AddItem ("0003")
        .AddItem ("0004")
        .AddItem ("0005")
        .AddItem ("0006")
        .AddItem ("0007")
        .AddItem ("0008")
        .AddItem ("0009")
        .AddItem ("0010")
        .AddItem ("0011")
        .AddItem ("0012")
        .AddItem ("0013")
        .AddItem ("0014")
        .AddItem ("0015")
    End With

If Err Then GrabarLog "cboPuntoDeVenta_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboLetra_GotFocus()
On Error Resume Next

    With cboLetra
        .Clear
        .AddItem ("A")
        .AddItem ("B")
        .AddItem ("C")
    End With

If Err Then GrabarLog "cboLetra_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboPuntoDeVenta_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = 13 Then txtNroComprobante.SetFocus
    
    If Err Then GrabarLog "cboPuntoDeVenta_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboPuntoDeVenta_LostFocus()
If Not Val(Me.vcodEmpresa.Text) > 0 Then
    vnrocomprobante2 = NroComprobanteNuevo(TipoDocumento, Trim(cboLetra.Text), Trim(cboPuntoDeVenta.Text))
    Me.txtNroComprobante = vnrocomprobante2
End If

End Sub

Private Sub cboTipoIva_Click()
On Error Resume Next

    cboTipoIva.Tag = TraerDato("TipoIva", "TipoIva = '" & Trim(cboTipoIva.Text) & "'", "idTipoIva")

If Err Then GrabarLog "cboTipoIva_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboTipoIva_GotFocus()
On Error Resume Next

    Call CargarComboNew("TipoIva", "TipoIva", cboTipoIva, True)

If Err Then GrabarLog "cboTipoIva_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cierrez_Click()
    Call frmPrincipal.cierrez_Click
End Sub

Private Sub cmdAct_Click()
Dim vsql As String
vsql = "update Configuracion set LeyendaFactura='" + Trim(Me.vleyenda.Text) + "' where Id=1"
Call EjecutarScript(vsql, PathDBConfig)

End Sub

Public Sub cmdActualizarTotal_Click()
    On Error Resume Next
    
    If chkTotalManual.Value = 0 Then CalcularTotales

    If Err Then GrabarLog "cmdActualizarTotal_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdBuscar_Click()
On Error Resume Next
    Call MostrarCoincidencias("Articulos", txtDetalle(1).Text)
If Err Then Exit Sub
End Sub

Private Sub cmdCambiarNombre_Click()
On Error Resume Next

    vNombreNuevo = ""
    vNombreNuevo = InputBox("Ingrese el nombre que desea actualizar", "Mensaje ...", "")
    
    If Not Trim(vNombreNuevo) = "" Then
        vActualizaNombre = True
        txtCliente(0).Text = vNombreNuevo
    Else
        vActualizaNombre = False
    End If
    
If Err Then GrabarLog "cmdCambiarNombre_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdCerrarPago_Click()
On Error Resume Next
Me.GBCaja.Visible = False
If Err Then GrabarLog "cmdCerrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdComentarios_Click()
On Error Resume Next

    frmComentariosFactura.Show

If Err Then GrabarLog "cmdComentarios_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdCommand1_Click()

   With Me
        .EpsonFP.ClosePort
       
        Call FPDelay
        .EpsonFP.CommPort = Val(Me.vpuerto)
        .EpsonFP.BaudRate = Val(Me.vvelocidad)
        .EpsonFP.ProtocolType = protocol_Extended
        
        MsgBox "Puerto " + Str(EpsonFP.CommPort)
        
        
        MsgBox "Velocidad " + Str(.EpsonFP.BaudRate)
  End With
End Sub

Private Sub cmdCommand2_Click()
'---------------------------
        ' Ticket
        '---------------------------
            'Open
            sCmd = Chr$(&HA) + Chr$(&H1)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H0)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmdExt)
            If bAnswer Then bAnswer = Me.EpsonFP.SendCommand
            Call FPDelay
            If Me.EpsonFP.ReturnCode <> 0 Then ShowMsg
            
            'Item
            sCmd = Chr$(&HA) + Chr$(&H2)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H0)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmdExt)
            
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("Descripci¢n Extra #1")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("Descripci¢n Extra #2")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("Descripci¢n Extra #3")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("Descripci¢n Extra #4")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("Descripci¢n Item")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("10000")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("1000")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("2100")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("")
            If bAnswer Then bAnswer = Me.EpsonFP.SendCommand
            Call FPDelay
            If Me.EpsonFP.ReturnCode <> 0 Then ShowMsg
            
            'Payment
            sCmd = Chr$(&HA) + Chr$(&H5)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H0)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmdExt)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("EFECTIVO")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("500")
            If bAnswer Then bAnswer = Me.EpsonFP.SendCommand
            Call FPDelay
            If Me.EpsonFP.ReturnCode <> 0 Then ShowMsg
            
            'Close
            sCmd = Chr$(&HA) + Chr$(&H6)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H1)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmdExt)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(1)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("Cola de reemplazo #1")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(2)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("Cola de reemplazo #2")
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(3)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField("Cola de reemplazo #3")
            If bAnswer Then bAnswer = Me.EpsonFP.SendCommand
            Call FPDelay
            If Me.EpsonFP.ReturnCode <> 0 Then ShowMsg
End Sub

Private Sub cmdEstado_Click()
           sCmd = Chr$(&H0) + Chr$(&H1)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H0)
            If bAnswer Then bAnswer = Me.EpsonFP.AddDataField(sCmdExt)
            If bAnswer Then bAnswer = Me.EpsonFP.SendCommand
            Call FPDelay
End Sub

Private Sub cmdGuardarPago_Click()
On Error Resume Next

 If controlNroFactura Then
                MsgBox "Este documento ya fue grabado anteriormente.", vbCritical, "Documento duplicado..."
                Exit Sub
End If

vnrointerno = Val(Me.txtNroInterno)

 ' ------------ verifica nro interno ----------------------
 If existeRegistro(Val(Me.txtNroInterno)) Then
 
    Exit Sub

 End If
 '----------------------------------------------------------



    'Panic :Faltan controles de todo tipo
    Dim i As Integer
    
    If Not Val(txtCaja(7).Text) = Val(txtIB(10).Text) Then
        MsgBox "Los Pagos no coinciden con el monto total del Comprobante !!", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    

    vLeyendaAsiento = Trim(txtIB(7).Text)
    
    If Val(txtIB(10).Text) < 0 Then
        vTotalAsiento = Val(Format(txtIB(10).Text * (-1), "#####0.00"))
    Else
        vTotalAsiento = Format(Val(txtIB(10).Text), "#####0.00")
    End If

    GuardarDoc

    GBCaja.Visible = False
    
    Me.txtNroInterno = vnrointerno + 1
    

    For i = 0 To Val(txtCaja.Count - 1)
        txtCaja(i).Text = ""
    Next
        
        
If Err Then GrabarLog "cmdGuardarPago_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub checkSaldo_Click()
If Me.checkSaldo Then lblsaldocliente.Caption = SaldoCliente
End Sub

Private Sub chkIncCod_LostFocus()
Call fijarparametro(Str(Me.chkIncCod.Value), Me.chkIncCod.Tag)
End Sub

Private Sub Desactivarlog_Click()
log.Visible = False
End Sub

Private Sub dgClientes_Click()
MsgBox "1!"
End Sub

Private Sub dtpFecha_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        If Not LeerConfig(17) = "Otros" Then
            focoEnLinea
            'txtDetalle(0).SetFocus
            
        Else
        
            Me.txtNroInterno.SetFocus
        
        End If
    
    End If

If Err Then GrabarLog "dtpFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub FlatEdit2_Change()

End Sub

Private Sub dtpFecha_LostFocus()
If dtpFecha.Text = "" Then
 MsgBox "Debe ingresar una fecha válida"
 dtpFecha.SetFocus
End If
End Sub

Private Sub estafiscal_Click()
Dim vdatos, vmensaje  As String
    Dim sCmd As String
    Dim sCmdExt As String
    Dim bAnswer As Boolean
    vdatos = UCase(LeerXml("Impresora"))

With frmPrincipal.FiscalEpson2



    If vdatos = UCase("Fiscal Ticket epson") Then
                sCmd = Chr$(&H0) + Chr$(&H1)
                bAnswer = .AddDataField(sCmd)
                sCmdExt = Chr$(&H0) + Chr$(&H0)
                If bAnswer Then bAnswer = .AddDataField(sCmdExt)
                If bAnswer Then bAnswer = .SendCommand
                Call FPDelay
                
                vmensaje = "Impresora: " + Format(Hex(.PrinterStatus), "0000") + Chr(13) + _
                "Fiscal: " + Format(Hex(.FiscalStatus), "0000") + Chr(13) + _
                "RC: " + Format(Hex(.ReturnCode), "0000") + Chr(13)

                MsgBox vmensaje
    
    End If



    If UCase(vdatos) = UCase("Fiscal Hasar") Then
                
                With frmPrincipal.FiscalHasar
                   ' vmensaje = .PedidoDeStatus
                End With
                
                MsgBox vmensaje
    End If


End With

End Sub

Private Sub factahasar_Click()
Dim msg As String
Dim comando As String
Dim FS As String

FS = Chr$(28)                                            '// Separador de campos del comando

On Error GoTo impresora_apag
Procesar:


With frmPrincipal.FiscalHasar

    .DatosCliente "Cliente...", "20061346326", TIPO_CUIT, RESPONSABLE_INSCRIPTO, _
                        "Domicilio..."
    .AbrirComprobanteFiscal FACTURA_A
    .ImprimirTextoFiscal "Texto Fiscal..."
    .ImprimirItem "Producto Uno", 1, 0.1, 21, 0
    .DescuentoUltimoItem "Oferta del Dia", 0.1, True
    .DescuentoGeneral "Oferta Pago Efectivo", 0.1, True
    .EspecificarPercepcionPorIVA "Percep IVA21", 0.1, 21
    .EspecificarPercepcionGlobal "Percep. RG 0000", 125#
    .ImprimirPago "Efectivo", 295#
    .CerrarComprobanteFiscal


End With
    
    
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

Private Sub Form_Initialize()
Me.txtCliente(0).SetFocus

Me.dgClientes.Left = 1080
Me.dgClientes.Top = 1560
Me.dgClientes.Width = 7935
Me.dgClientes.Height = 3730

End Sub

Private Sub guardarOtro_Click()

Exit Sub

vGrabaModo = 1
 
manualcae = 1
guardarinicio

Me.vcae.Text = ""
Me.vcaeFecha2.Text = ""

End Sub

Private Sub KlexDetalle_Click()
filaSeleccionada = Me.KlexDetalle.Row

End Sub

Private Sub mglobal_Click()
MsgBox vmensajeGlobal
End Sub

Private Sub nc_Click()
Dim msg As String
Dim comando As String
Dim FS As String



On Error GoTo impresora_apag
Procesar:

With frmPrincipal.FiscalHasar

    .CerrarDNFH

    .DatosCliente "Cliente...", "20061346326", TIPO_CUIT, RESPONSABLE_INSCRIPTO, _
                        "Domicilio..."
    .AbrirDNFH (NOTA_CREDITO_A)
    .DocumentoDeReferencia(1) = "0003-00000001"
    .ImprimirTextoFiscal "Texto Fiscal..."
    .ImprimirItem "Producto Uno", 1, 0.1, 21, 0
    .CerrarDNFH
    
    
End With
    
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If



End Sub

Private Sub nc2_Click()
Dim msg As String
Dim comando As String
Dim FS As String



On Error GoTo impresora_apag
Procesar:

With frmPrincipal.FiscalHasar

    .ReComenzar
    
End With
    
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If


End Sub

Private Sub nofiscalhasar_Click()
Dim msg As String
Dim j As Integer

On Error GoTo impresora_apag
       
Procesar:
    
    With frmPrincipal.FiscalHasar
    
        
   
        .AbrirComprobanteNoFiscal
    
            For j = 1 To 10
                .ImprimirTextoNoFiscal "Linea Texto No Fiscal..."
            Next j
            
        .CerrarComprobanteNoFiscal
    
    
    
    End With
    
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
    
End Sub

Private Sub opOtrosDocumentos_Click()
On Error Resume Next
    
    GBOtrosDocumentos.Visible = True
    GBOtrosDocumentos.Left = 60
    GBOtrosDocumentos.Top = fraTipoDocumento.Top + 500
    
    fraDetalle.Visible = Not True
    fraTotales.Visible = Not True
    PbAcciones(0).Visible = Not True
    PbAcciones(1).Visible = Not True
    PbAcciones(2).Visible = Not True
    
    'lblComentario.Visible = Not True
    txtObservaciones.Visible = Not True
    
   ' FraAccionesDoc.Top = GBOtrosDocumentos.Top + GBOtrosDocumentos.Height
   ' FraAccionesDoc.Width = GBOtrosDocumentos.Width
   ' FraAccionesDoc.Left = GBOtrosDocumentos.Left
    
   ' Me.Height = FraAccionesDoc.Top + FraAccionesDoc.Height + 500
    
    txtTipoMovimiento(0).SetFocus
    
    txtIB(1).Text = 21

If Err Then GrabarLog "opOtrosDocumentos_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdVer_Click(Index As Integer)
On Error Resume Next
    
    If Index = 0 Then
        With frmBuscarFactura
            .Show
            '.c1.Value = 0
            '.c2.Value = 0
            '.c3.Value = 0
            '.c4.Value = 0
            '.c5.Value = 0
            '.c6.Value = 1
            '.c7.Value = 0
            '.opFecha_Click (0)
            .TabDocumentos.tab = 0
            .cmdFiltrar_Click
        End With
        
        Unload Me

    Else
        With frmBuscarFactura
            '.c1.Value = 0
            '.c2.Value = 0
            '.c3.Value = 0
            '.c4.Value = 1
            '.c5.Value = 0
            '.c6.Value = 0
            '.c7.Value = 0
            '.opFecha_Click (0)
            .TabDocumentos.tab = 0
            .cmdFiltrar_Click
        End With
        Unload Me
    
    End If
    
If Err Then GrabarLog "cmdVer_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdBuscarArticulo_Click()
    On Error Resume Next
    
    MousePointer = vbHourglass
    BuscarArticulo
    MousePointer = vbDefault

    If Err Then GrabarLog "cmdBuscarArticulo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdGuardarPrecio_Click()
    On Error Resume Next
    
    If Not Trim(txtDetalle(1).Tag) = "" And Not Val(cbolista.Text) = 0 Then
        Call EjecutarScript("UPDATE Articulos SET PVenta" & Val(cbolista.Text) & " = " & Val(txtDetalle(2).Text) & " WHERE (Codigo = '" + Trim(txtDetalle(1).Tag) + "')")
        
        txtDetalle(6).SetFocus
    
        MsgBox "El Precio del Articulo fue Actualizado para la Lista Nº " & Val(cbolista.Text), vbInformation, "Mensaje..."
    Else
        MsgBox "Los datos del artículo no pudieron ser Actualizados", vbExclamation, "Mensaje..."
    End If
    

    If Err Then GrabarLog "cmdGuardarPrecio_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub BorrarFDetalleRemito(vremito As Long)
On Error Resume Next
Dim vsql As String

' borro los registros del stock
vsql = "delete from stock where idFDetalle in (select idFDetalle  from fdetalle  where Remito=" + Str(vremito) + ")"
Call EjecutarScript(vsql, pathDBMySQL)


' borro el fdetalle
vsql = "delete from fdetalle where remito=" + Str(vremito)
Call EjecutarScript(vsql, pathDBMySQL)


If Err Then
    MsgBox "Error"
    Exit Sub
End If
End Sub
Private Sub ConfirmarDetalle()
    On Error Resume Next
    
    Dim vcant_actual As Double
    
    Dim vcodigo As String
    
    'FormatoGrillaDetalle (1)
    
    Dim vdife1 As Double, i As Integer, vConceptoCaja
  
    Dim rsFDetalle As New ADODB.Recordset, sqlFDetalle As String
    
    
    
    BorrarFDetalleRemito (vremito) ' borro el fdetalle si es que existe
    
        
    sqlFDetalle = "SELECT * FROM fdetalle WHERE (Remito = " & Val(txtNroRemito.Text) & ") ORDER BY idFDetalle"
    
    With rsFDetalle
        Call .Open(sqlFDetalle, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then .MoveFirst
        
        For i = 1 To KlexDetalle.Rows - 1
        
            vcodigo = Replace(Replace(KlexDetalle.TextMatrix(i, 4), "[", ""), "]", "")
            vcant_actual = Val(KlexDetalle.TextMatrix(i, 5)) - Val(Json.getDic(vcodigo, arr2))
            

            
          '  If Not Trim(KlexDetalle.TextMatrix(i, 1)) = "" Then
          '      .Filter = "idFDetalle = " & Trim(KlexDetalle.TextMatrix(i, 1)) & ""
          '      'Se Borro, Algo malo Paso
          '      If .EOF = True Then .AddNew
          '  Else
            If Val(Val(KlexDetalle.TextMatrix(i, 5))) > 0 Then
                .AddNew
          '  End If
          
                
            
                .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value) ' KlexDetalle.TextMatrix(i, 0)
                               
                .Fields("Remito").Value = vremito
                .Fields("Cantidad").Value = Val(KlexDetalle.TextMatrix(i, 5))
                .Fields("Codigo").Value = Replace(Replace(KlexDetalle.TextMatrix(i, 4), "[", ""), "]", "")
                .Fields("Detalle").Value = EsNulo(KlexDetalle.TextMatrix(i, 6))
                .Fields("Precio").Value = Val(KlexDetalle.TextMatrix(i, 7))         '(precio)
                .Fields("tiva").Value = Val(KlexDetalle.TextMatrix(i, 9))           '(tiva)
                '.fields("devolucion").Value = KlexDetalle.TextMatrix(i, 7)         '(devolucion)
                .Fields("Total").Value = Val(KlexDetalle.TextMatrix(i, 11))         '(total)
                .Fields("pcosto").Value = Val(KlexDetalle.TextMatrix(i, 12))
                .Fields("Descuento").Value = Val(KlexDetalle.TextMatrix(i, 8))
       
                If TipoDocumento = "Fact A" And Trim(cboTipoIva.Text) = "Iva Responsable Inscripto" Then
                    .Fields("totaliva").Value = Val(KlexDetalle.TextMatrix(i, 17))  '(totaliva)
                Else
                    .Fields("TotalIva").Value = Val(KlexDetalle.TextMatrix(i, 11))
                End If
                     
                .Fields("confirmado").Value = "S"
                .Fields("repartidor").Value = Trim(txtEmpleados(0).Text)

                ' bdetalle.Recordset("totaliva") = Talles.TextMatrix(Talles.Row, 8)
                
                If Not KlexDetalle.TextMatrix(i, 21) = "S" Then
                    .Update
                    KlexDetalle.TextMatrix(i, 1) = .Fields("idFDetalle").Value

                    If TipoDocumento = "Nota C" Then
                        Call GuardarEnStock("Devolucion", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha.Value), vcant_actual, "Devolucion de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                    Else
                        If Not TipoDocumento = "Presupuesto" Then
                            If Not KlexDetalle.TextMatrix(i, 24) = "S" Then
                                Call GuardarEnStock("Remito-Nuevo", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha.Value), vcant_actual, "Salida de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                            Else
                                Call GuardarEnStock("Remito-Nuevo", EsNulo(KlexDetalle.TextMatrix(i, 4)), strfechaMySQL(dtpFecha.Value), vcant_actual, "Actualizacion de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                            End If
                        End If
                    End If
            
                Else
                    Call GuardarEnStock("Remito-Modificar", EsNulo(.Fields("Codigo").Value), strfechaMySQL(dtpFecha.Value), vcant_actual, "Entrada de Mercaderia", KlexDetalle.TextMatrix(i, 1), 0)
                    .MoveNext
                End If
        
        End If
        
        Next
    
    
    End With

    Call frellenar_renglones(vremito, i)

    sqlFDetalle = ""
    
    If rsFDetalle.State = 1 Then
        rsFDetalle.Close
        Set rsFDetalle = Nothing
    End If
    
If Err < 0 Then
    GrabarLog "ConfirmarDetalle", Left(Err.Number & " " & Err.Description, 99), Me.Name
    MsgBox "Revise si el documento fue guardado correctamente", vbCritical, "Cuidado"
Else
    checksum(2) = True
End If
End Sub

Private Sub guardarFdetalleTemp()
Dim vValor, vcampos, vsql As String
Dim i, vrow As Integer

 Call BorrarBase("Documentos", PathDBListados)

vcampos = ""
vValor = ""

With Me.KlexDetalle
vrow = .Rows

For i = 1 To vrow - 1
        vValor = .TextMatrix(i, 5) + ",'" + Replace(Replace(.TextMatrix(i, 4), "[", ""), "]", "") + "','" + .TextMatrix(i, 6) + "'," + Str(Val(.TextMatrix(i, 7))) + "," + Str(Val(.TextMatrix(i, 8))) + "," + Str(Val(.TextMatrix(i, 9))) + "," + Str(Val(.TextMatrix(i, 11)))
        vcampos = "cantidad,codigo,descripcion,pventa,descuento,iva,total"
        
        vsql = "insert into Documentos (" + vcampos + ") values (" + vValor + ")"
        
        Call EjecutarScript(vsql, PathDBListados)
Next
End With

End Sub

Private Sub CtaCte()
    On Error Resume Next
    
    With frmCtaCteC
        .Show
        frmCtaCteC.WindowState = vmaximizar
        
        '.txtCliente.Text = Trim(v(0).Text)
        '.txtCliente_Keypress (13)
    End With
    
    If Err Then GrabarLog "CtaCte", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdNotaCredito_Click()
On Error Resume Next

    If opTipoDoc(2).Value Then
        'Antes de guardar tengo que pedir los datos de las facturas
        frmNroFactNC.Show
    End If

If Err Then GrabarLog "cmdNotaCredito_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgArticulos_DblClick()
On Error Resume Next

If UCase(LeerXml("Puesto")) = UCase("Empresas") Then
            txtDetalle(1).Text = rsArticulos.Fields("Detalle").Value
            txtDetalle(2).Text = rsArticulos.Fields("Precio").Value
           ' Call txtDetalle_KeyPress(1, 13)
            txtDetalle(1).Tag = 99
Else

    With rsArticulos
        If Not .EOF = True And Not .BOF = True Then
            txtDetalle(1).Text = .Fields("Codigo").Value
            txtDetalle(1).Tag = .Fields("Codigo").Value
            
            If .Fields("stock") <= 0 Then
                txtDetalle(0).BackColor = vbRed
            End If
                
            
            Call txtDetalle_KeyPress(1, 13)
        End If
    End With
    
End If

If Err Then GrabarLog "DgArticulos_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgClientes_DblClick()
On Error Resume Next

    'vClienteNuevo = False
    
    txtCliente(0).Text = EsNulo(rsClientes.Fields("Codigo").Value)
    Me.txtcodigoCliente.Text = txtCliente(0).Text
    
    Call txtCliente_KeyPress(0, 13)
    dgClientes.Visible = False
    
If Err Then GrabarLog "dgClientes_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgClientes_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        dgClientes_DblClick
    End If
With Me.dgClientes

.RowBookmark (.Row)

End With
If Err Then GrabarLog "dgClientes_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgEmpleados_DblClick()
On Error Resume Next

    With rsEmpleadosGrilla
        If .ActiveConnection = "" Then
            Exit Sub
        End If
        
        If Not .EOF = True Then
            txtEmpleados(0).Text = .Fields("Codigo").Value
            txtEmpleados(1).Text = .Fields("Nombre").Value
            dgEmpleados.Visible = Not True
            focoEnLinea
            'txtDetalle(0).SetFocus
        End If
    End With
    
If Err Then GrabarLog "dgEmpleados_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgEmpleados_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        dgEmpleados_DblClick
    End If

If Err Then GrabarLog "dgEmpleados_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgArticulos_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        dgArticulos_DblClick
    End If

If Err Then GrabarLog "dgArticulos_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Option1_Click()
'frmFormulario1116.Show
End Sub

Function getTipoParaFE() As Integer  ' doing Ale

If Me.opTipoDoc(0).Value Then ' FACTURA
            If Me.cboLetra.Text = "A" Then
                vtipoFactura = 1
            End If
            
            If Me.cboLetra.Text = "B" Then
                
                If cboTipoIva.Text = "Consumidor Final" Then
                    vtipoFactura = 6
                Else
                    vtipoFactura = 6
                End If
            
            
            
            End If
            
            If Me.cboLetra.Text = "C" Then
                vtipoFactura = 11
            End If
            
            vetipoFactura = "FACTURA"
End If


If Me.opTipoDoc(5).Value Then ' ND
vetipoFactura = "NOTA DEBITO"
            If Me.cboLetra.Text = "A" Then
                vtipoFactura = 2
            End If
            
            If Me.cboLetra.Text = "B" Then
                vtipoFactura = 7
            End If
End If


If Me.opTipoDoc(2).Value Then ' NC
vetipoFactura = "NOTA CREDITO"

            If Me.cboLetra.Text = "A" Then
                vtipoFactura = 3
            End If
            
            If Me.cboLetra.Text = "B" Then
                vtipoFactura = 8
            End If
End If
 
End Function


Private Sub PusCAE_Click()
On Error Resume Next

Dim vcuit As String
Dim vmodotest As Integer
Dim vc1, vc2 As Long

Call getTipoParaFE

If vGrabaModo = 1 Or chkCaeTest.Value = xtpChecked Then Exit Sub

If manualcae = 0 Then
           ' If Not InputBox("Ing. Clave: ") = "dalas.2015" Then
           '     Exit Sub
           ' End If
End If

manualcae = 0

vcuit = Replace(Me.txtCliente(3).Text, "-", "")

If Me.chkCaeTest.Value Then
    vmodotest = 0
Else
    vmodotest = 0
End If

' nro desde afip

Debug.Print "Ultimo Nro de comprobante + 1 " + txtNroComprobante

'txtNroInterno = traerDatos2("select max(numero) as c from t_nrointerno", "c", pathDBMySQL) + 1


'Call fecae(fe, vtipoFactura, Str(Val(Me.txtNroComprobante)), Me.txtSubtotal, Me.txtTotal, vCuit, _
Format(Me.dtpFecha, "yyyymmdd"), Val(Me.cboPuntoDeVenta.Text), Me.txtNroInterno, vc1, vc2, Val(Me.txtIva(0).Text), Val(Me.txtIva(1).Text), Val(Me.txtIva(2).Text), vmodotest, Val(txtNroInterno))

vnroempresa = Val(Me.vcodEmpresa.Text)


Dim vtD As Integer
Dim vcuit_afip As String


vtD = 0

vcuit_afip = vcuit

If Me.cboTipoIva = "Consumidor Final" Then
    vcuit_afip = 0
    vtD = 99
End If


' inicializo la variable global que uso para control
vnrocomprobante_control1 = ""


'2021 facturación electrónica



Call fecae2(fe, vtipoFactura, Str(Val(Me.txtNroComprobante)), Me.txtSubtotal, Me.txtTotal, vcuit_afip, _
Format(Me.dtpFecha, "yyyymmdd"), Val(Me.cboPuntoDeVenta.Text), Me.txtNroInterno, vc1, vc2, Val(Me.txtIva(0).Text), Val(Me.txtIva(1).Text), Val(Me.txtIva(2).Text), vmodotest, Val(txtNroInterno), , vtD)


Me.vcae.Text = Str(vc1)
vcaeFecha = Str(vc2)
vcaeFecha2 = Str(vc2)

global_vpunto_venta = Me.cboPuntoDeVenta.Text
global_vtipoFactura = vtipoFactura
global_vnro_comprobante = Me.txtNroComprobante

vnrocomprobante_control1 = Me.txtNroComprobante

'Call control_datos_afip_vs_datos_factura


If Err < 0 Then
    MsgBox "Ocurrió el siguiente error: " + Err.Description
End If

End Sub

Private Sub control_datos_afip_vs_datos_factura()

' 2021 ------ todo nuevo

' chequeamos si la factura con el mismo
Dim bresult, bresult2 As Boolean
Dim vtotal_consultado As Variant

' hago un llamado para consultar
bresult = fe.F1CompConsultarS(global_vpunto_venta, global_vtipoFactura, global_vnro_comprobante)
' consulto el total de la factura que acabo de generar
vtotal_consultado = fe.F1DetalleImpTotal()
' controlo si el valor que me devuelve AFIP es el mismo del monto de la factura
bresult2 = Val(vtotal_consultado) = Val(Me.txtTotal)

If Not bresult2 Then
    MsgBox ("No coinciden los valores de la factura con los registado por AFIP " + _
    Chr(13) + "AFIP informa un total de: " + vtotal_consultado + _
    Chr(13) + "El total de la factura es:" + Me.txtTotal)
End If

End Sub




Private Sub setTicketAFIP()

Dim bResultado As Boolean
Dim cIdentificador As String
Dim v As Variant

Dim vvcertificado, vWSAFIPFE As String


' doing empresas wsfe FE
vvcertificado = Trim(LeerXml("vcertificado"))
vWSAFIPFE = "WSAFIPFE.lic"
  
If Val(Me.vcodEmpresa.Text) > 0 Then
    vWSAFIPFE = Trim$(Me.vcodEmpresa.Text) + ".lic"
End If

' -------------------------------------

  
  If Not LeerXml("ObtieneCAE") = "SI" Then Exit Sub
  
  If Trim(LeerXml("modoFiscal")) = "1" Then
     'bResultado = fe.iniciar(1, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + "WSAFIPFE.lic") ' Paso 1
    
    bResultado = fe.iniciar(1, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))   ' Paso 1
    

    'bResultado = fe.iniciar(1, "30707384316", App.Path + "\PoliCertificadoProduccion11.pfx", App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))
  Else
   
    bResultado = fe.iniciar(0, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))   ' Paso 1

   ' bResultado = fe.iniciar(0, Trim(LeerXml("vcuit")), Trim(LeerXml("vcertificado")), "WSAFIPFE.lic")    ' Paso 1
   '  bResultado = fe.iniciar(0, Trim(LeerXml("vcuit")), Trim(LeerXml("vcertificado")), vWSAFIPFE)    ' Paso 1
  End If
  
  If bResultado Then
     'If Not fe.f1TicketEsValido Then bResultado = fe.ObtenerTicketAcceso()
     bResultado = fe.ObtenerTicketAcceso()
     vultimoMensajeError = Str(fe.f1TicketEsValido) + " "
  End If

If bResultado Then
 
strTicket = fe.f1GuardarTicketAcceso()
Debug.Print strTicket
'Me.Caption = fe.f1TicketValido



MsgBox "Ultimo comprobante 1: " + Str(fe.f1CompUltimoAutorizado(2, 1))

MsgBox "Ultimo comprobante 6: " + Str(fe.f1CompUltimoAutorizado(2, 6))



 Else
 
 
 Debug.Print " Hubo un problema con los datos generados por AFIP " + Chr(13) + "El documento no se puede generar" + Chr(13) _
         + "Motivo: " + fe.FERespuestaDetalleMotivo + Chr(13) + " Detalle: " + fe.UltimoMensajeError
         Chr (13) + "El sistema se cerrará "
 
  MsgBox " Hubo un problema con los datos generados por AFIP " + Chr(13) + "El documento no se puede generar" + Chr(13) _
         + "Motivo: " + fe.FERespuestaDetalleMotivo + Chr(13) + " Detalle: " + fe.UltimoMensajeError
         Chr (13) + "El sistema se cerrará "
  End

End If


End Sub


Public Function getNroCompAfip() As Integer
On Error Resume Next
Dim vNroComprobanteAnterior As Long
Dim vcantiIVA, vultimoNroComprobante As Integer


Screen.MousePointer = vbHourglass


Call getTipoParaFE


' Me.vcodEmpresa.Text = Val(Me.vcodEmpresa.Text)


If UCase(LeerXml("ObtieneCAE")) = "NO" Then
    vcae = ""
    vcaeFecha = ""
    Exit Function
End If
  
' Documentación en: https://sites.google.com/site/facturaelectronicax/documentacion-wsfev1/wsfev1/wsfev1-metodos
  
  Dim bResultado As Boolean
  Dim cIdentificador As String
  Dim v As Variant
  
  v = Test
  
  vdatos_mandante = ""
  
 If Trim(LeerXml("modoFiscal")) = "1" Then
  '  bResultado = fe.iniciar(1, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))
    bResultado = fe.iniciar(1, Trim(Str(getCuitFE(Me.vcodEmpresa.Text))), App.Path + "\" + Trim(getCertificadoFE(Me.vcodEmpresa.Text)), App.Path + "\" + Trim(getLicenciaFE(Me.vcodEmpresa.Text)))
  
   Debug.Print Trim(getCuitFE(Me.vcodEmpresa.Text)) + " " + App.Path + "\" + Trim(getCertificadoFE(Me.vcodEmpresa.Text)) + " " + Trim(getLicenciaFE(Me.vcodEmpresa.Text))
 
  vdatos_mandante = "cuit: " + Str(getCuitFE(Me.vcodEmpresa.Text)) + " certificado: " + getCertificadoFE(Me.vcodEmpresa.Text) + " licencia: " + getLicenciaFE(Me.vcodEmpresa.Text)
    
  
  Else
   'bResultado = fe.iniciar(0, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))
    'bResultado = fe.iniciar(0, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), "")
     
     bResultado = fe.iniciar(0, "20249182940", App.Path + "\" + "sartorio22.pfx", "")
     fe.ArchivoCertificadoPassWord = ""
     
     
     'bResultado = fe.iniciar(0, Trim(Str(getCuitFE(Me.vcodEmpresa.Text))), App.Path + "\" + Trim(getCertificadoFE(Me.vcodEmpresa.Text)), "")

    ' bResultado = fe.iniciar(0, Trim(getCuitFE(Me.vcodEmpresa.Text)), App.Path + "\" + Trim(getCertificadoFE(Me.vcodEmpresa.Text)), "")
    Debug.Print Trim(getCuitFE(Me.vcodEmpresa.Text)) + " " + App.Path + "\" + Trim(getCertificadoFE(Me.vcodEmpresa.Text))



End If

'bResultado = fe.iniciar(1, "30707384316", App.Path + "\PoliCertificadoProduccion11.pfx", App.Path + "\WSAFIPFE.lic")
   
bResultado = fe.f1ObtenerTicketAcceso()

MsgBox fe.UltimoMensajeError + "  " + Str(fe.UltimoNumeroError)

           ' Next

Dim vnomr As String

vnomr = "recibido" + Format(Date, "yymmdd") + Replace(Str(Time()), ":", "") + ".xml"


'MsgBox ("Despues de f1ObtenerTicket: " + fe.UltimoMensajeError)

'fe.ArchivoXMLRecibido = App.Path + "\Log\recibido" + Replace(Str(Time()), ":", "") + ".xml"

fe.ArchivoXMLRecibido = App.Path + "\Log\" + vnomr

fe.ArchivoXMLEnviado = App.Path + "\Log\enviado" + Replace(Str(Time()), ":", "") + ".xml"

Dim vvtemp As Integer
vvtemp = fe.f1CompUltimoAutorizadoS(Val(Me.cboPuntoDeVenta.Text), Val(vtipoFactura))

txtNroComprobante = LeerXmlRecibido(vnomr)

vultimoMensajeError = fe.UltimoMensajeError

vNroComprobanteAnterior = txtNroComprobante

vultimoMensajeError = fe.UltimoMensajeError
Debug.Print "Nro de comprobante devuelto por el WS " + vNroComprobanteAnterior


txtNroComprobante = vNroComprobanteAnterior + 1 'le sumo uno para que siga correlativo
'MsgBox " Lo paro acá "
'Exit Sub
vnrocomprobante2 = txtNroComprobante

getNroCompAfip = Val(vnrocomprobante2)

Screen.MousePointer = vbDefault

If Err < 0 Then
    MsgBox Err.Description

    getNroCompAfip = 0
    Exit Function
End If
End Function

Private Sub push_consultar_doc_Click()
    Dim bResultado As Boolean
    Dim punto_venta, tipo_doc As Integer
    Dim nro_comp As String
    
    
    tipo_doc = InputBox("Ingresar el nro del tipo de comprobante" + Chr(13) + "1- Fact A, 6- Fact B, 3- NC A")
    nro_comp = Trim(Str(InputBox("Ingresar el nro de comprobante: ")))
    
    punto_venta = CInt(cboPuntoDeVenta.Text)
    
    
    
           bResultado = fe.F1CompConsultarS(CInt(punto_venta), tipo_doc, nro_comp)
           
           
           If fe.UltimoMensajeError = "" Then
              
            
            ' genero el qr del documento que estoy consultando
            vqrnombre = Trim(Replace(txtCliente(3), "-", "")) + Trim(nro_comp)
            fe.F1Detalleqrarchivo = App.Path + "\" + vqrnombre + ".jpg"
            fe.F1Detalleqrformato = 6
            fe.F1Detalleqrtipocodigo = "E"
            fe.F1Detalleqrtolerancia = 1
            fe.F1Detalleqrresolucion = 2
              
              
              
              MsgBox "- CAE consultado: " + fe.F1RespuestaDetalleCae + Chr(13) + "- Fecha Vto : " + fe.F1DetalleCbteFch + Chr(13) _
              + "- Total: " + Str(fe.F1DetalleImpTotal) + Chr(13) _
              + "- Nro compronate: " + nro_comp + Chr(13) _
              + "- Tipo comprobante: " + Str(tipo_doc) + Chr(13) _
              + "- CUIT " + fe.F1DetalleDocNro + Chr(13) _
              + Chr(13) _
              + "Si este comprobante no está ingresado en el sistema, UD. debe REINGRESAR FACTURA poniendo de modo manual el CAE y la Fecha que aquì se indica"
           
            If MsgBox("Desea imprimir un compronate con estos datos fiscales", vbYesNo) = vbYes Then
                      
                Me.vcae.Text = fe.F1RespuestaDetalleCae
                Me.vcaeFecha2.Text = fe.F1DetalleCbteFch
                Me.chkReingresarFact.Value = xtpChecked
                
                MsgBox "Debe completar todos los datos del encabezado del documento correspondientes a este documento"
                
            End If
            
           
           Else
              MsgBox ("fallo consulta: " + fe.UltimoMensajeError)
           End If
End Sub

Private Sub PushButton1_Click()
MsgBox fcomoponerlistachoferes
End Sub

Private Sub PushButton2_Click()
Me.grd_Choferes.Rows = 0
End Sub

Private Sub PushButton3_Click()
vnrocomprobante2 = NroComprobanteNuevo(TipoDocumento, Trim(cboLetra.Text), Trim(cboPuntoDeVenta.Text))
Me.txtNroComprobante = vnrocomprobante2
End Sub



Private Sub PushButton5_Click()
If log.Visible Then
    log.Visible = False
Else
    log.Visible = True
End If
End Sub

Private Sub PushButton6_Click()
Dim vsql, vc1, vc2 As String

vsql = "(Select * from proveedores where tipocliente  = 'Empresa') t"
vc1 = "Nombre"
vc2 = "Codigo"

Call fbuscarGrilla(vsql, vc1, vc2, Me.vdescEmpresa.Name, Me)

End Sub

Private Sub PushButton7_Click()
Dim vsql, vc1, vc2 As String

If UCase(LeerXml("Login")) = "MANUAL" Then
    MsgBox "No tiene permiso"
    Exit Sub
End If

vsql = "(Select * from proveedores where tipocliente  = 'Vendedor') t"
vc1 = "Nombre"
vc2 = "Codigo"

Call fbuscarGrilla(vsql, vc1, vc2, Me.vDesRepartidor.Name, Me)
End Sub

Private Sub PushButton8_Click()
        'vGrabaModo = 0
        
        
        If Not MsgBox("Esta seguro que quiere duplicar este documento ?", vbYesNo) = vbYes Then Exit Sub
        
        vduplicando = True
        
        manualcae = 1
        guardarinicio
        
        Me.vcae.Text = ""
        Me.vcaeFecha2.Text = ""
        
         vGrabaModo = 0
End Sub

Private Sub PushButton9_Click()
ifacta1.Show
ifacta2.Show
ifacta3.Show
ifacta4.Show
ifacta5.Show
ifacta6.Show
ifacta7.Show
ifacta8.Show

ifacta12.Show


End Sub

Private Sub PusNroComp_Click()
On Error Resume Next
Dim a As Boolean

            
MsgBox "Nro de comprobante AFIP: " + Str(getNroCompAfip) + _
Chr(13) + "Debería continuar con el número siguiente " + _
Chr(13) + Chr(13) + " Mensaje del sistema al programador : " + Err.Description _
+ Chr(13) + " Mensaje de AFIP al programador: " + vultimoMensajeError _
+ Chr(13) + " - Ticket hora de vencimiento: - Es Válido: " + fe.f1TicketHoraVencimiento + " " + Str(fe.f1TicketValido) _
+ Chr(13) + vdatos_mandante

vultimoMensajeError = ""


If Err Then Exit Sub
End Sub

Private Sub PusNuevo_Click()
limpiarCliente
Me.txtCliente(0).SetFocus
End Sub

Private Sub PusTextil_Click()
frmTextilArt.Show
End Sub

Private Sub txtCaja_GotFocus(Index As Integer)
On Error Resume Next

    txtCaja(Index).SelStart = 0
    txtCaja(Index).SelLength = Len(txtCaja(Index).Text)

If Err Then GrabarLog "txtCaja_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtCaja_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        txtCaja(Index).Text = UCase(txtCaja(Index).Text)
        
        Select Case Index
        
            Case 0
                txtCaja(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtCaja(Index).Text) & "'", "Descripcion")
                txtCaja(Index + 2).SetFocus
            
            Case 2
                txtCaja(Index + 1).Text = TraerDato("BancosCuentas", "idBancosCuentas = " & Trim(txtCaja(Index).Text) & "", "Cuenta")
                txtCaja(Index + 2).SetFocus
                
            
            Case 4
                txtCaja(Index + 1).Text = TraerDato("TipoValor", "idTipoValor = '" & Trim(txtCaja(Index).Text) & "'", "TipoValor")
                
                If Not Trim(txtCaja(Index + 1).Text) = "" Then
                    If Not UCase(Trim(txtCaja(Index + 1).Text)) = "CH" Then
                        txtCaja(7).Text = ""
                        txtCaja(7).SetFocus
                    Else
                        txtCaja(8).SetFocus
                    End If
                Else
                    txtCaja(Index).Text = ""
                    txtCaja(Index + 1).Text = ""
                End If
                
            Case 6, 7
                txtCaja(Index + 1).SetFocus
            
            Case 8
                cmdGuardarPago.SetFocus
        
        End Select
    
    
        'If txtCaja(Index).Text = "" Then txtIB(0).SetFocus
    
    End If
    
If Err Then GrabarLog "txtCaja_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtCaja_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF3 Then
        Select Case Index
        
            Case 0
                pbCarga_Click (2)
            Case 1
                
            Case 2
                pbCarga_Click (3)
            
            Case 3
                
            
            Case 4
                pbCarga_Click (4)
            Case 5
                
            
            Case 6
        
        End Select
        
        
        
    End If

If Err Then GrabarLog "txtCaja_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtCliente_DblClick(Index As Integer)
limpiarCliente
End Sub

Private Sub txtCliente_GotFocus(Index As Integer)
Me.txtCliente(Index).BackColor = vbYellow
End Sub

Private Sub txtCliente_LostFocus(Index As Integer)
On Error Resume Next
    
    Me.txtCliente(Index).BackColor = vbWhite
    
    Select Case Index

        Case 0
            vOpenGrilla(0) = False
            dgClientes.Visible = Not True
            
        Case 1
        
        Case 2
        
        Case 3
        
        Case 4
        
        Case 5
    
    End Select

If Err Then GrabarLog "txtCliente_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtDetalle_Click(Index As Integer)
On Error Resume Next

    With txtDetalle(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

If Err Then GrabarLog "txtDetalle_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtDetalle_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If Index = 1 Then
        If KeyCode = 38 Then
            With rsArticulos
                If Not .EOF = True And Not .BOF = True Then
                    .MovePrevious
                Else
                    .MoveLast
                End If
            End With
        End If

        If KeyCode = 40 Then
            With rsArticulos
                If Not .EOF = True And Not .BOF = True Then
                    .MoveNext
                Else
                    .MoveFirst
                End If
            End With
        End If
    
        If KeyCode = 13 And Not Trim(txtDetalle(Index).Text) = "" Then
            dgArticulos_DblClick
        End If
    End If
    
If Err Then GrabarLog "f_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub KlexDetalle_DblClick()
On Error Resume Next

    With KlexDetalle
        .Col = 25
        Select Case .TextMatrix(.Row, 25)
        
            Case "Blanco"
                .CellBackColor = vbRed
                .TextMatrix(.Row, 25) = "Negro"

            Case "Negro"
                .TextMatrix(.Row, 25) = "Blanco"
                .CellBackColor = vbGreen
                

            Case ""
                .TextMatrix(.Row, 25) = "Negro"
                .CellBackColor = vbRed
        
        End Select
        
    End With

If Err Then GrabarLog "KlexDetalle_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub KlexDetalle_LeaveCell()
On Error Resume Next

   ' Exit Sub
    
    'No se ejecuta
    
    
    
    
    With KlexDetalle
    
        
    
        Select Case .Col
        
            Case 0, 4
        
            Case 5 'Cantidad
                .TextMatrix(.Row, 11) = .TextMatrix(.Row, 5) * .TextMatrix(.Row, 7)
                CalcularTotales
                
        
            Case 7 'Precio
                .TextMatrix(.Row, 11) = .TextMatrix(.Row, 5) * .TextMatrix(.Row, 7)
                CalcularTotales
                
            
            Case 8
                .TextMatrix(.Row, 11) = DescuentoImpuesto
                CalcularTotales
             
            
            Case 9
            
            Case 10
            
            Case 11
            
            Case 25
            
        End Select



    End With

If Err Then GrabarLog "KlexDetalle_LeaveCell", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function DescuentoImpuesto() As Double
On Error Resume Next

    Dim vdescuento As Double, vImpuesto As Double
    
    Dim vAuxiliar As Double
    
    With KlexDetalle
        vAuxiliar = .TextMatrix(.Row, 5) * .TextMatrix(.Row, 7)
        
        If Not .TextMatrix(.Row, 8) = "" Then
            vdescuento = vAuxiliar - (vAuxiliar * Val(.TextMatrix(.Row, 8)) / 100)
                    
                    
        Else
        
        End If
        

    End With
    
    DescuentoImpuesto = vdescuento + vImpuesto

    
If Err Then GrabarLog "DescuentoImpuesto", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub PbAcciones_Click(Index As Integer)

On Error Resume Next
    
    Select Case Index
    
        Case 0
            'If vGrabaModo = 1 Then
              '  If MsgBox("Desea Restaurar el stock despues de borrar el Registro?", vbInformation + vbYesNoCancel, "Mensaje ...") = vbYes Then

            '        Call ModificarStock(1, KlexDetalle.TextMatrix(KlexDetalle.Row, 2), KlexDetalle.TextMatrix(KlexDetalle.Row, 3))
                
            '    End If
                
            '    Call BorrarBase("FDetalle WHERE (idFDetalle = " & KlexDetalle.TextMatrix(KlexDetalle.Row, 22) & ")", pathDBMySQL)
            'End If
                
           ' KlexDetalle.RemoveItem KlexDetalle.RowSel
           
           ' KlexDetalle.RowSel = filaSeleccionada
            
            KlexDetalle.RemoveItem KlexDetalle.RowSel

            
            If vCantidadControl >= 1 Then
                vCantidadControl = vCantidadControl - 1
            End If
            
            If vCantidadControl < 1 Then
                        Call FormatoGrillaDetalle(1)
            
                        vCantidadControl = 0
                        
                        Me.KlexDetalle.Tag = 0
            End If
            
            CalcularTotales
            
        Case 1
            Dim i As Integer, j As Integer
            
            Call FormatoGrillaDetalle(1)
            
            Me.KlexDetalle.Tag = 0  ' para que arranque en el primero de la fila
            
            vCantidadControl = 0
    
    
    
        Case 2
            frmExcel.Show
            frmExcel.vVieneDesdeExcel = "Remito"
        
    End Select

If Err Then
    ' GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Name
    Exit Sub
End If
End Sub
Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

Dim vsql, vc1, vc2 As String

vsql = "(Select * from proveedores where tipocliente  = 'Empleado') t"
vc1 = "Nombre"
vc2 = "Codigo"

Call fbuscarGrilla(vsql, vc1, vc2, Me.vdescEmpresa.Name, Me)
            
If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PBDocAFactura_Click(Index As Integer)
On Error Resume Next

    Dim i As Integer, j As Integer
    
    '0- Validar
        'A- Controlar Todos los Detalles Con IVA (10,21,27)
        'B- Controlar Condicion de Iva del Cliente (Monotributo, Responsable Inscripto, Exento)
        'C- Otros (No Implementado)
    
    '1- Cambiar el Tipo de Documento
    '2- Cambiar los Detalles (Sumar el Iva o Restar el Iva)
    '3- ReCalcular Totales
    '4- Borrar el Documento Original (chkBorrarDocOriginal)
        
    '0
    If ValidarDocAFactura = False Then
        MsgBox "No estan bien Algunos Parametros del Documento o del Cliente", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    '1
    opTipoDoc(0).Value = True
    
    '2
    Call CambiarIvaEnDetalles
    
If Err Then GrabarLog "PBDocAFactura_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function ValidarDocAFactura() As Boolean
On Error Resume Next

    Dim i As Integer, j As Integer
    '0
    
    'A
    With KlexDetalle
        For i = 1 To .Rows - 2
            
            If Val(.TextMatrix(i, 11)) = 10.5 Or Val(.TextMatrix(i, 11)) = 21 Or .TextMatrix(i, 11) = 27 Then
                ValidarDocAFactura = True
            Else
                ValidarDocAFactura = False
                Exit Function
            End If
        
        Next
    
    End With

    'B
    If Trim(cboTipoIva.Text) = "Responsable Inscripto" Or Trim(cboTipoIva.Text) = "Resp. Inscripto" Or Trim(cboTipoIva.Text) = "Exento" Or Trim(cboTipoIva.Text) = "Responsable Monotributo" Then
        ValidarDocAFactura = True
    End If
    
If Err Then GrabarLog "ValidarDocAFactura", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub CambiarIvaEnDetalles()
On Error Resume Next

    Dim i As Integer, j As Integer
    
    '2
    With KlexDetalle
        For i = 1 To .Rows - 2
            
            'Cantidad
            .TextMatrix(i, 2) = .TextMatrix(i, 2)
                 
            'Tiva
            .TextMatrix(i, 11) = .TextMatrix(i, 11)
            
            If RBIva(0).Value = True Then
                'Resta el Iva al Detalle
                
                'Precio
                .TextMatrix(i, 5) = .TextMatrix(i, 5) / Val(1 & "." & Val(Replace(.TextMatrix(i, 11), ".", "")))
                
                'Total
                .TextMatrix(i, 8) = Val(.TextMatrix(i, 2)) * Val(.TextMatrix(i, 5))
                
                'Total_CtaCte
                .TextMatrix(i, 10) = ""
                
                'Pago
                .TextMatrix(i, 15) = ""
                
                'Resta
                .TextMatrix(i, 16) = ""
                
                'TotalIva
                .TextMatrix(i, 17) = ""
            
            Else
            
                'Suma el Iva al Detalle
                .TextMatrix(i, 2) = ""      'Cantidad
                .TextMatrix(i, 5) = ""      'Precio
                .TextMatrix(i, 8) = ""      'Total
                .TextMatrix(i, 10) = ""     'Total_CtaCte
                .TextMatrix(i, 11) = ""     'T-IVA (tiva)
                .TextMatrix(i, 15) = ""     'Pago
                .TextMatrix(i, 16) = ""     'Resta
                .TextMatrix(i, 17) = ""     'TotalIva
            End If
            
        Next
    
    End With

If Err Then GrabarLog "CambiarIvaEnDetalles", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub TabTipoDetalle_Click(PreviousTab As Integer)
On Error Resume Next

    Dim i As Integer
    
    Select Case TabTipoDetalle.tab

        Case 0
            
            For i = 0 To txtDetalle.Count - 1
                txtDetalle(i).Visible = True
            Next
            
            txtDetalle(1).Left = 1020
            txtDetalle(1).Width = 4000
            txtDetalle(1).BackColor = &HFFFFFF
            
            cmdBuscarArticulo.Visible = True
            cmdGuardarPrecio.Visible = True
            focoEnLinea
            'txtDetalle(0).SetFocus
        Case 1
        
            For i = 0 To txtDetalle.Count - 1
                txtDetalle(i).Visible = Not True
            Next

            cmdBuscarArticulo.Visible = Not True
            cmdGuardarPrecio.Visible = Not True

            txtDetalle(1).Visible = True
            txtDetalle(1).Left = 60
            txtDetalle(1).Width = 9500
            txtDetalle(1).BackColor = &HFFFFC0
    
            txtDetalle(1).SetFocus
    End Select
    
    ConfigurarGrilla

If Err Then GrabarLog "TabTipoDetalle_Click", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub txtDetalle_LostFocus(Index As Integer)
Me.txtDetalle(Index).BackColor = vbWhite
End Sub

Private Sub txtEmpleados_Change(Index As Integer)
On Error Resume Next

    Exit Sub
    If Not vGrabaModo = 1 Then
        Call MostrarCoincidencias("Empleados", txtEmpleados(Index).Text)
        vOpenGrilla(2) = True
    End If
    
If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtIB_GotFocus(Index As Integer)
Me.txtIB(7) = "[Doc: " + Me.txtTipoMovimiento(0) + " " + cboLetra + " " + txtNroComprobante + "]"
End Sub

Private Sub txtIB_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        txtIB(Index + 1).SetFocus
    End If

If Err Then GrabarLog "txtIB_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtImpuesto_Change()
On Error Resume Next

    txtImpuesto.Text = Format(txtImpuesto.Text, "########0.000")


Call cmdActualizarTotal_Click


If Err Then GrabarLog "txtImpuesto_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtIB_Change(Index As Integer)
On Error Resume Next
    
    Dim vCalculoIva As Double
    Select Case Index
    
        Case 0
            txtIB(2).Text = Val(Format(txtIB(0).Text * Val(txtIB(1).Text / 100), "#######0.00"))
            
        
        Case 2, 3, 4, 5, 6
            txtIB(10).Text = Val(txtIB(0).Text) + Val(txtIB(2).Text) - Val(txtIB(3).Text) - Val(txtIB(4).Text) + Val(txtIB(5).Text) + Val(txtIB(6).Text)
        
        Case 1
            If (Val(txtIB(1).Text) = 21) Or (Val(txtIB(1).Text) = 10.5) Or (Val(txtIB(1).Text) = 27) Then
                txtIB(2).Text = Val(Format(txtIB(0).Text * Val(txtIB(1).Text / 100), "#######0.00"))
            End If
    End Select
    

If Err Then GrabarLog "txtIB_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtNroComprobante_Click()
On Error Resume Next

    With txtNroComprobante
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

If Err Then GrabarLog "txtNroComprobante_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtNroComprobante_GotFocus()
On Error Resume Next

    With txtNroComprobante
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

If Err Then GrabarLog "txtNroComprobante_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtNroComprobante_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    
    If KeyAscii = 13 Then
        If opOtrosDocumentos.Value = True Then
        
            Me.txtTipoMovimiento(0).SetFocus
            'dtpFecha.SetFocus
        
        End If
    End If
    
    If Err Then GrabarLog "txtNroComprobante_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtNroInterno_Click()
On Error Resume Next

    With txtNroInterno
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

If Err Then GrabarLog "txtNroInterno_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtNroInterno_GotFocus()
On Error Resume Next

    With txtNroInterno
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

If Err Then GrabarLog "txtNroInterno_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtNroInterno_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        cboLetra.SetFocus
    End If
    

If Err Then GrabarLog "txtNroInterno_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtPDescuento_Change()
    On Error Resume Next
    Dim vauxi As Double
    
    vauxi = Val(txtSubtotal) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text)
    vauxi = (vauxi * Val(txtPDescuento.Text) / 100)
    txtDescuento.Text = vauxi
    Call cmdActualizarTotal_Click


    If Err Then GrabarLog "txtPDescuento_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtImpuesto_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        If (chkTotalManual = 0) Then
            'CalcularTotales
        Else
             txtTotal.SetFocus
        End If
    End If

    If Err Then GrabarLog "txtImpuesto_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtDetalle_Change(Index As Integer)
    On Error Resume Next
    
    Dim descuento, impuesto, vtotal  As Double
    
    
  '  If escodigodebarra(vcbarra) Then
  '      txtDetalle(1).Text = vcbarra
  '     ' Me.txtDetalle(1).Text = Me.txtDetalle(0).Text
  '      Me.txtDetalle(0).Text = 1
  '       Call Me.txtDetalle_KeyPress(1, 13)
  '  End If
    

    If Index = 1 Then
        
        If chkfijo Then Exit Sub
        
        If Len(txtDetalle(1).Text) > 4 Or Left(Me.txtDetalle(1).Text, 2) = "**" Or Not UCase(LeerXml("Cliente")) = "KIOSCO" Then
                Call MostrarCoincidencias("Articulos", txtDetalle(Index).Text)
                vOpenGrilla(1) = True
        End If
    
    Else
        
        If (ConfigRemito(5) = False) And (Val(cbolista.Text) = 0) Then
           cbolista.Text = 1
            'MsgBox "Debe cargar un número de lista para poder facturar ", vbInformation, "Mensaje ..."
            'cboLista.BackColor = vbRed
            'cboLista.SetFocus
            'Exit Sub
        End If
    
        cbolista.BackColor = vbWhite
    
        vtotal = Val(txtDetalle(0).Text) * Val(txtDetalle(2).Text)
        
        descuento = Val(txtDetalle(3).Text) * vtotal / 100
        impuesto = Val(txtDetalle(5).Text) * vtotal / 100

        If (Val(txtDetalle(0).Text) * Val(txtDetalle(2).Text)) > 0 Then
            txtDetalle(6).Text = vtotal - descuento + impuesto
        End If
    End If
    
    
    If Err Then GrabarLog "f_change", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub MostrarCoincidencias(vTipoBusqueda As String, vBusqueda As String)
On Error Resume Next
    
    
    If UCase(LeerXml("Textil")) = "ADBA" And vTipoBusqueda = "Articulos" Then
        vTipoBusqueda = vTipoBusqueda + "-" + Me.txtcodigoCliente.Text
    End If
    
    ' poner acá la exclusion por codigo de barra
    
    'If Not vcbarra = "" Then
    '    vcbarra = ""
    '    Exit Sub
    'End If
    
    Select Case vTipoBusqueda
    
        Case "Articulos"
            Dim sqlArticulos As String, sqlTipoDetalle As String
    
            Set rsArticulos = New ADODB.Recordset
    
            If Trim(txtDetalle(1).Text) = "" Then
                sqlArticulos = "SELECT * FROM Articulos WHERE 1=2"
            Else
        
        
        
                    If UCase(LeerXml("Puesto")) = UCase("Empresas") Then
                                               
                                       ' If Val(vBusqueda) > 0 Then
                                            sqlArticulos = "SELECT  '-',codigo, detalle, detalle, detalle, '-', '-',Precio, Precio FROM fdetalle WHERE (codigo LIKE '%" & Trim(vBusqueda) & "%') OR (detalle LIKE '%" & Trim(vBusqueda) & "%') group by detalle limit 50 "
                                       ' End If
                                    
                    Else
        
        
                                        If Val(vBusqueda) > 0 And Not UCase(LeerXml("Puesto")) = "KIOSCO" Then
                                            sqlArticulos = "SELECT * FROM Articulos WHERE (Codigo like '%" & Trim(vBusqueda) & "%')"
                                        Else
                                            sqlArticulos = "SELECT * FROM Articulos WHERE (Codigo LIKE '%" & Trim(vBusqueda) & "%') OR (Descrip LIKE '%" & Trim(vBusqueda) & "%')"
                                        End If
                   End If
            
            End If
    
            With rsArticulos
                If .State = 1 Then .Close
        
                .CursorLocation = adUseClient
            
                Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
            
                dgArticulos.Visible = Not .EOF
            
                If Not .EOF = True Then
                    Set dgArticulos.DataSource = rsArticulos
                    
                    Call FormatoGrilla("Articulos")
                Else
                    Set dgArticulos.DataSource = Nothing
                End If
            
            End With
    
            sqlArticulos = ""

        Case "Clientes"
            Dim sqlClientes As String
            
            Set rsClientes = New ADODB.Recordset
            
            If Trim(vBusqueda) = "" Then
                sqlClientes = "SELECT * FROM Clientes WHERE 1=2"
            Else
                
                
             If Not vBusqueda = "" Then
             
            If Val(vBusqueda) > 0 Then
                sqlClientes = "SELECT * FROM Clientes WHERE (Codigo = '" & Trim(vBusqueda) & "')"

            Else
                sqlClientes = "SELECT * FROM Clientes WHERE (Codigo LIKE '%" & Trim(vBusqueda) & "%') OR (Nombre LIKE '%" & Trim(vBusqueda) & "%')"
            End If
            End If
             
                
                
            
            End If
            
            With rsClientes
                If .State = 1 Then .Close
            
                .CursorLocation = adUseClient
            
                Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
            
                dgClientes.Visible = Not .EOF
                
                BarraCliente.Buttons(1).Enabled = .EOF
                Dim i As Integer
                
                For i = 1 To txtCliente.Count - 1
                    txtCliente(i).Enabled = .EOF
                Next
                cboTipoIva.Enabled = .EOF
                                        
                'vClienteNuevo = .EOF
        
                If Not .EOF = True Then
                    Set dgClientes.DataSource = rsClientes
                    Call FormatoGrilla("Clientes")
                Else
                    Set dgClientes.DataSource = Nothing
                End If
            
            End With
            
            sqlClientes = ""
    
        Case "Empleados"
            Dim sqlEmpleados As String
            
            Set rsEmpleadosGrilla = New ADODB.Recordset
            
            If Trim(vBusqueda) = "" Then
                sqlEmpleados = "SELECT * FROM Empleados WHERE 1=2"
            Else
                sqlEmpleados = "SELECT * FROM Empleados WHERE (Codigo LIKE '%" & Trim(vBusqueda) & "%') OR (Nombre LIKE '%" & Trim(vBusqueda) & "%')"
            End If
            
            With rsEmpleadosGrilla
                If .State = 1 Then .Close
            
                .CursorLocation = adUseClient
            
                Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockReadOnly)
            
                dgEmpleados.Visible = Not .EOF

                If Not .EOF = True Then
                    Set dgEmpleados.DataSource = rsEmpleadosGrilla
                    Call FormatoGrilla("Empleados")
                Else
                    Set dgEmpleados.DataSource = Nothing
                End If
            
            End With
            
            sqlEmpleados = ""
            
    End Select


If Err Then GrabarLog "MostrarCoincidencias", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(vtipo As String)
On Error Resume Next
    
    Dim i As Integer
    
    Select Case vtipo
    
        Case "Articulos"
    
            If Val(cbolista.Text) = 0 Then cbolista.Text = 1
            
            With dgArticulos
                'Lo Paso al Frente
                .ZOrder (0)
        
                'Lo Ubico justo debajo de donde escribo
                '.Top = Me.fraDetalle.Height + Me.fraDetalle.Top - 125 'fraCargaDetalle.Top + fraCargaDetalle.height
        
                '.Left = txtDetalle(1).Left
              '  .Width = txtDetalle(1).Width + txtDetalle(2).Width + txtDetalle(3).Width + txtDetalle(4).Width 'txtDetalle(1).Width
        
                .HeadLines = 1.2
            
                Dim indicePVenta As Integer
                indicePVenta = ElegirColumnaSegunListaPrecio
                
                For i = 0 To .Columns.Count - 1
                    
                    Select Case i
                        Case 1
                            .Columns(i).Width = 1000
                            .Columns(i).Caption = "Código"
                        Case 4
                            .Columns(i).Width = txtDetalle(1).Width
                            .Columns(i).Caption = "Descripción"
                        Case indicePVenta
                            .Columns(indicePVenta).Width = 750
                            .Columns(indicePVenta).Caption = "Precio"
                    Case 9
                            .Columns(i).Caption = "Proveedor"
                            .Columns(i).Width = 1000
                    Case 22
                            .Columns(i).Width = 750
                    Case 23
                            .Columns(i).Caption = "Stock"
                            .Columns(i).Width = 1000
                        Case Else
                            .Columns(i).Width = 0
                    End Select
                Next

        End With
    
        Case "Clientes"
            With dgClientes
        
                .ZOrder (0)
               ' .Top = 1155
                '.Left = 1260

                .HeadLines = 1.2
        
                For i = 0 To .Columns.Count - 1
        
                    Select Case i
        
                        Case 3
                            .Columns("Nombre").Width = .Width - 1000
                        
                        Case Else
                            .Columns(i).Width = 0
                
                    End Select
                Next

            End With
    
        Case "Empleados"
            With dgEmpleados
        
                .ZOrder (0)
                .Top = 3200
                .Left = 6150

                .HeadLines = 1.2
        
                For i = 0 To .Columns.Count - 1
        
                    Select Case i
        
                        Case 1
                            .Columns("Nombre").Width = .Width - 1000
                        
                        Case Else
                            .Columns(i).Width = 0
                
                    End Select
                Next

            End With
    
    End Select
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function ElegirColumnaSegunListaPrecio()
On Error Resume Next

    Select Case Val(cbolista.Text)

        Case 1
            ElegirColumnaSegunListaPrecio = 12
        Case 2
            ElegirColumnaSegunListaPrecio = 13
        Case 3
            ElegirColumnaSegunListaPrecio = 14
        Case 4
            ElegirColumnaSegunListaPrecio = 15
        Case 5
            ElegirColumnaSegunListaPrecio = 16
        
    End Select

If Err Then GrabarLog "ElegirColumnaSegunListaPrecio", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub txtDetalle_GotFocus(Index As Integer)
    On Error Resume Next
    
    Me.txtDetalle(Index).BackColor = vbYellow

    'If Me.txtDetalle(1).Text = "" And UCase(LeerXml("Puesto")) = "KIOSCO" Then
    '    Me.txtDetalle(1).Text = "**"
    'End If
    If Index = 0 Then Exit Sub
    
    With bdetalle
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM fdetalle WHERE (remito = " & Val(txtNroRemito.Text) & ") ORDER BY idFDetalle ASC"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveLast
    End With
    
    If txtDetalle(4).Text = "" Then txtDetalle(4).Text = "21"
    
    With txtDetalle(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

    If Err Then GrabarLog "txtDetalle_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub


Public Sub txtDetalle_KeyPress(Index As Integer, _
                      KeyAscii As Integer)

    On Error Resume Next
    
    
        If KeyAscii = 13 Then
       
        Select Case Index

            Case 1
                
                'Pasar (Index)
                'Exit Sub
                
                If Not vOpenGrilla(1) = True Then Pasar (Index)
                
                With barticulo
                    .ConnectionString = pathDBMySQL
                    
                    If vOpenGrilla(1) = True Then
                        With rsArticulos
                            If Not (.EOF = True) And Not (.BOF = True) Then
                            
                                If UCase(LeerXml("Puesto")) = UCase("Empresas") Then
                                           Call dgArticulos_DblClick
                                Else
                                            barticulo.RecordSource = "SELECT * FROM Articulos WHERE (codigo =  '" & Trim(.Fields("Codigo").Value) & "')"
                                            barticulo.Refresh
                                End If
                            
                            Else
                            
                                Dim vlong As Long
                            
                                vlong = CLng(Me.txtDetalle(1).Text)
                            
                              If Val(Trim(Me.txtDetalle(1).Text)) > 11111111 Then
                                    vcbarra = Me.txtDetalle(1).Text
                                    barticulo.RecordSource = "SELECT * FROM Articulos WHERE (CodigoBarra = '" & Trim((Trim(Me.txtDetalle(1).Text))) & "') order by pventa1 desc limit 1"
                                    barticulo.Refresh
                                Else
                                    barticulo.RecordSource = "SELECT * FROM Articulos WHERE (Descrip LIKE '%" & Trim(txtDetalle(1).Text) & "%') OR (codigo LIKE '%" & Trim(txtDetalle(1).Text) & "%')"
                                    barticulo.Refresh
                                End If
                                
                            
                            End If
                            
                        End With
                    Else
                        .RecordSource = "SELECT * FROM Articulos WHERE (codigo =  '" & Trim(txtDetalle(1).Text) & "')"
                        .Refresh
                    End If
                    
        
                    
                    If .Recordset.EOF = True Then
                        
                        If Not .Recordset.EOF Then
                        
                            txtDetalle(1).Tag = EsNulo(.Recordset("codigo").Value)
                        
                            vganancia = TraerDato("Articulos_Ganancia", "(CodEmp = '" & Trim(txtEmpleados(0).Text) & "') AND (CodCli = '" & Trim(txtCliente(0).Tag) & "') AND (CodRub = '" & barticulo.Recordset("rubro").Value & "')", "Porcentaje")
                            venvase = .Recordset("Envase").Value
                        
                            If vganancia = 0 Then
                                vganancia = Val(Format(.Recordset("Ganancia").Value, "#######0.000"))
                            End If
                        
                            MostrarDetalle
                            ElegirTipoPrecio

                        Else
                            
                            'doy True si lo carga desde remito
                            vArticuloNuevo = True
                            
                            ' verificar si es un código de barra nuevo
                            
                            If escodigodebarra(Me.txtDetalle(1).Text) Then
                                        frmArticulosAlta.txtAlta(6).Text = Me.txtDetalle(1).Text
                                                    
                                        ' si es cero lo pongo en 1
                                        If Val(Me.txtDetalle(0)) = 0 Then
                                            Me.txtDetalle(0).Text = 1
                                        End If
                            
                            Else
                                    frmArticulosAlta.txtAlta(1).Text = Me.txtDetalle(1).Text
                            End If
                            
                            frmArticulosAlta.vViene = "frmRemito"
                            frmArticulosAlta.Show
                            Exit Sub
                        
                        End If
                    Else
                        MostrarDetalle
                        'ElegirTipoPrecio
                    End If
                
                End With
                
                dgArticulos.Visible = False

                
                Pasar (Index)

            Case Else
                Pasar (Index)
        End Select

    End If

    If KeyAscii = 10 Then
        
        Select Case Index

            Case 1
                Pasar (Index)
                frmBuscarArticulo.busca = 1
                barticulo.Refresh
                'barticulo.Recordset.Sort = "descrip"
                barticulo.Recordset.Find ("descrip like '" + Trim(txtDetalle(1)) + "%'")
                
                vvvdescrip = txtDetalle(1).Text  ' para ver q codigo usa
                MostrarDetalle
                ElegirTipoPrecio
                
                txtDetalle(1).Tag = EsNulo(barticulo.Recordset("codigo").Value)
                
                'buscaart
                Pasar (Index)

            Case Else
                Pasar (Index)
        End Select

    End If

    If Err Then GrabarLog "f_keypress (" & Index & "-" & KeyAscii & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarArticuloNuevo(vDetalle As String, vPrecio As Double, vPorcenjeIva As String)
On Error Resume Next

    Dim rsArticuloNuevo As New ADODB.Recordset, sqlArticuloNuevo As String, i As Integer
    
    sqlArticuloNuevo = "SELECT * FROM Articulos WHERE 1=2"
    
    With rsArticuloNuevo
        Call .Open(sqlArticuloNuevo, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then
        
            .AddNew
            
            txtDetalle(1).Tag = Val(GenerarDato("SELECT idArticulos, Codigo FROM Articulos ORDER BY idArticulos DESC", "Codigo")) + 1
            txtDetalle(1).Tag = FormatoUltimoCodigo(5, txtDetalle(1).Tag)
            
            .Fields("Codigo").Value = txtDetalle(1).Tag
            .Fields("CodigoNum").Value = Val(.Fields("Codigo").Value)
            .Fields("Descrip").Value = vDetalle
            .Fields("idPorcentajeIva").Value = TraerDato("PorcentajeIva", "Porcentaje = " & vPorcenjeIva & "", "idPorcentajeIva")

            .Fields("idRubros").Value = ""
            .Fields("idSubRubros").Value = ""
            .Fields("CodigoBarra").Value = ""
            
            .Fields("idProveedores").Value = ""
            .Fields("idFabricantes").Value = ""
            
            .Fields("PCosto").Value = 0
            
            For i = 1 To 6
                If i = Val(cbolista.Text) Then
                    .Fields("PVenta" & Val(cbolista.Text) & "").Value = vPrecio
                Else
                    .Fields("PVenta" & Val(i) & "").Value = 0
                End If
            Next
            
            .Fields("Stock").Value = Val(txtDetalle(0).Text) * (-1)
            
            .Fields("FechaAlta").Value = strfechaMySQL(Date)
            .Fields("Observaciones").Value = "CargadoPorRemito"
        
            .Update
    
        End If
        
    End With
    
    sqlArticuloNuevo = ""

    If rsArticuloNuevo.State = 1 Then
        rsArticuloNuevo.Close
        Set rsArticuloNuevo = Nothing
    End If
    
If Err Then GrabarLog "GuardarArticuloNuevo", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, _
                       Shift As Integer)
                    
On Error Resume Next

    
    If KeyCode = 13 Then
    
        ' Call Keytab
        Call Form_KeyUp(vbTab, 1)
    
    End If

    If KeyCode = vbKeyF10 Then Me.txtDetalle(0).SetFocus
    

    If KeyCode = vbKeyF1 Then
       
      ' Call Me.limpiarCliente
       Me.txtCliente(0).Text = ""
       Me.txtCliente(0).Tag = 0
       
       Me.txtCliente(0).SetFocus
       
       
       ' If MsgBox("¿ Desea ver la Ayuda del Formulario " & Me.Caption & "?", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
       '     'Call VerAyuda(Me.Name)
       ' End If
        
        'venta.SetFocus
    End If
    
    If KeyCode = vbKeyF2 Then
                
            If Not validarGrabar Then
                    MsgBox "Verificar el nro de comprobante y el punto de venta", vbCritical
                    Exit Sub
            End If

        
            
            manualcae = 1
            guardarinicio
        
            Me.vcae.Text = ""
            Me.vcaeFecha2.Text = ""
            
            
            
            If Me.vcventa.Text = "Contado" Then
                frmCobros.SetFocus
            End If
            
            
           'MsgBox "Presione <Enter> para continuar"
            
            
    End If
    
    If KeyCode = vbKeyF3 Then
        
    End If
    
    If KeyCode = vbKeyF4 Then
        'Call ImprimirFacturaHasar("", 1)
    End If
    
If KeyCode = vbKeyF5 Then
            
            If Not validarGrabar Then
                    MsgBox "Verificar el nro de comprobante y el punto de venta", vbCritical
                    Exit Sub
            End If

        
            If vcaeFecha = 0 And Not Me.vcaeFecha2.Text = "" Then vcaeFecha = Val(Me.vcaeFecha2.Text)
            manualcae = 1
            fimprimirDoc
            Me.vcae.Text = ""
            Me.vcaeFecha2.Text = ""
            
            If Me.vcventa.Text = "Contado" Then
                frmCobros.SetFocus
            End If
            
            'MsgBox "Presione <Enter> para continuar "
            
    End If
    
    If KeyCode = vbKeyF6 Then
        Call PushButton6_Click
    End If
        
    If KeyCode = vbKeyF7 Then
        Call PushButton7_Click
    End If
    
    If KeyCode = vbKeyF8 Then
    End If

    If KeyCode = vbKeyF9 Then
        cmdCambiarNombre_Click
    End If
    
    If KeyCode = vbKeyF12 Then
        
        If Me.chkfijo.Value = xtpChecked Then
            Me.chkfijo.Value = xtpUnchecked
        Else
        
             Me.chkfijo.Value = xtpChecked
        End If
    
    End If
    
    'Pendiente: Que hace esto
    'If KeyCode = vbKeyF10 Then
    '    If Not cboVenta.ListCount - 1 = cboVenta.ListIndex Then
    '        cboVenta.ListIndex = cboVenta.ListIndex + 1
    '    Else
    '        cboVenta.ListIndex = 0
    '    End If
    'End If
    
    'configurarGrilla

    If KeyCode = vbKeyF11 Then
        txtCliente(0).SetFocus
    End If



 
    
    'CargoDatosDelClienteSeleccionado

    
    If Err Then GrabarLog "Form_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub fimprimirDoc()

txtNroRemito.Text = 0
  
If UCase(LeerXml("Puesto")) = "EMPRESAS2" Then
    Me.rdotros.Value = True
    vestructura = "12"
    If Me.vcodEmpresa.Text = "" Then Me.vcodEmpresa.Text = "12"
End If


If Not obtenerCAE Or abortarFactura Then Exit Sub


vCodigoBarra = feCodigoBarra(Me.txtCliente(3), vtipoFactura, cboPuntoDeVenta, Me.vcae, Me.vcaeFecha2)


If abortarFactura Then Exit Sub

vLetra = Me.cboLetra
vPtoVta = cboPuntoDeVenta
vnrocomprobante2 = txtNroComprobante.Text
vsaldodeudor = Val(lblsaldocliente.Caption)
        
        
        
         If Me.vGrabaModo = 0 Then
            txtNroRemito.Text = NroRemitoNuevo
            txtNroInterno = UltimoNroInterno2
            
            Do Until Not existeRegistro(Val(Me.txtNroInterno))
            txtNroInterno.Text = UltimoNroInterno2
        Loop
        
         End If
        
        ' verificar_nrointerno (txtNroInterno.Text)
        
         If Not validarGuardarDocumento Then Exit Sub
    
        
        If Me.rdT.Value Then
             
            If GuardarCompleto Then
                vcancelartrans = True
                Exit Sub
            End If
 
            If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket HASAR") Or UCase(LeerXml("Impresora")) = UCase("Fiscal HASAR") _
            Then
                Call ImprimirHasar(vremito, 0)
            End If
            
            If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket EPSON") Then Call ImprimirEpson(vremito, 0)
            
            RecargarForm
            limpiarCliente
        
        Else
                Call Imprimir
        End If
        
        vIdFactura = utltimoFactura
        vIdEmpresa = codigo2id(vcodEmpresa.Text)
        vidVendedor = codigo2id(Me.vcodRepartidor.Text)
        Call GuardarRel(vIdFactura, vIdEmpresa, vidVendedor, vnrointerno)
               
        limpiarCliente
        vnrocomprobante = 0
        
End Sub
Public Sub ConfigurarGrilla()
On Error Resume Next

    Dim i As Integer, j As Integer

    Select Case TabTipoDetalle.tab
    
        Case 0
            With KlexDetalle
                .Cols = 23
                
                .FixedRows = 1
                .FixedCols = 1

                .TextMatrix(0, 2) = "Cantidad"
                .TextMatrix(0, 3) = "Código"
                .TextMatrix(0, 4) = "Detalle"
                .TextMatrix(0, 5) = "Precio"
                .TextMatrix(0, 6) = "Desc."
                .TextMatrix(0, 8) = "Total"
        
                .ColWidth(0) = 100
                .ColWidth(1) = 0
                .ColWidth(2) = 1000
                .ColWidth(3) = 1300
                .ColWidth(4) = 9000
                .ColWidth(5) = 900
                .ColWidth(6) = 900
                .ColWidth(7) = 0
                .ColWidth(8) = 1000
                .ColWidth(9) = 0

                .Row = KlexDetalle.Rows - 1
            End With
        
        Case 1
            
            With KlexDetalle
                .Cols = 2
                .FixedRows = 1
                .FixedCols = 1

                .TextMatrix(0, 1) = "Detalle del Movimiento"
                '.ColAlignmentFixed(1) = 3
                
                .ColWidth(1) = 9000
                

            End With
    
    End Select
    
If Err Then GrabarLog "ConfigurarGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub Habilitar(vHabilita As Boolean)
On Error Resume Next
    
    Dim i As Integer
    
    vOpenGrilla(0) = False
    
    With Me
       ' .fraTipoDocumento.Enabled = vHabilita

        For i = 1 To txtCliente.Count - 1
            
            txtCliente(i).Enabled = vHabilita
            
        Next
        
        .FraAccionesDoc.Enabled = vHabilita
        .fraCargaDetalle.Enabled = vHabilita
        .fraConfig.Enabled = vHabilita
        .fraDetalle.Enabled = vHabilita
        .fraPrecio.Enabled = vHabilita
        .fraTotales.Enabled = vHabilita
        .KlexDetalle.Enabled = vHabilita
        .PbAcciones(0).Enabled = vHabilita
        .PbAcciones(1).Enabled = vHabilita
        .PbAcciones(2).Enabled = vHabilita
        .fraDocAbrir.Enabled = vHabilita
        .BarraCliente.Enabled = vHabilita
        .dtpFecha.Enabled = vHabilita
        
        .TabTipoDetalle.Enabled = vHabilita
        
        .cboTipoIva.Enabled = vHabilita
        
        .cboLetra.Enabled = vHabilita
        .cboPuntoDeVenta.Enabled = vHabilita
        .txtNroComprobante.Enabled = vHabilita
    
    End With

If Err Then GrabarLog "Habilitar", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub Form_Load()
    On Error Resume Next
    
    'setTicketAFIP
    

    With Me
       ' .Show
        .Top = 0
        .Left = 0
        .Width = 15090
        If Not LeerConfig(17) = "Otros" = True Then .Height = 9465
        .KeyPreview = True
        .vGrabaModo = 0
    End With

   

    ReDim vOpenGrilla(2)
    
    vOpenGrilla(0) = False
    vOpenGrilla(1) = False
    vOpenGrilla(2) = False
    
    dtpFecha.Value = Date

    FormatoGrillaDetalle (1)
    
    chkTotalManual.Value = 0
    
    'vNroFacturaNotaC = 0

    Dim i As Integer
    For i = 0 To 5
        opTipoDoc(i).Value = False
    Next
    
    txtCliente(0).SetFocus
    
    'Me.txtNroComprobante.Text = EsNulo(frmPrincipal.FiscalHasar.UltimaFactura)
    If LeerConfig(17) = "Otros" Then
        opOtrosDocumentos_Click
    End If
    
    Call CentrarFormulario(Me)
    
    txtCliente(0).SetFocus
    
    txtNroRemito.Text = NroRemitoNuevo ' busca el nro deremito
    
    
    ' cargo el mismo cliente que estaba
    
     If Me.txtCliente(0).Text = "" Then
        Call Habilitar(False)
    
    Else
        ' mantiene el cliente
        BuscarCliente
        txtNroInterno.SetFocus
    End If
    
Me.vcventa.Clear
Me.vcventa.Text = "Contado"
Me.vcventa.AddItem ("Contado")
Me.vcventa.AddItem ("Cuenta Corriente")

    init
    
    If Err > 0 Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub init()
On Error Resume Next

vcancelartrans = False


Me.vcodEmpresa.Text = "0"

If UCase(LeerXml("Login")) = "MANUAL" Then
    'Me.vDesRepartidor.Tag = vConfigGral.vIdUsuario
    
    Dim vsql5 As String
    
    vsql5 = "select Nombre,Codigo from proveedores where codigo = '" + Trim(vConfigGral.vIdUsuario) + "'"
    
    Me.vDesRepartidor.Text = traerDatos2(vsql5, "Nombre", pathDBMySQL)
    Me.vcodRepartidor.Text = traerDatos2(vsql5, "Codigo", pathDBMySQL)
    
    vCodigoRepartidor2 = vcodRepartidor.Text
    
  '  Call vDesRepartidor_Change
End If


If UCase(LeerXml("ObtieneCAE")) = "PRUEBA" Then
    chkReingresarFact.Value = xtpChecked
End If

'vduplicando = False
            
 If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket Hasar") Then
 
              Me.rdT.Value = True
              frmPrincipal.FiscalHasar.Modelo = 32 'LeerXml("Modelo")
              frmPrincipal.FiscalHasar.Puerto = 1  'LeerXml("Puerto")
             ' frmremito.Caption = frmremito.Caption +  " Modelo: "  +
              frmPrincipal.FiscalHasar.Comenzar
 End If
 
 
 If UCase(LeerXml("Impresora")) = UCase("Fiscal Hasar") Then
 
              Me.rdT.Value = True
           '   frmPrincipal.FiscalHasar.Modelo = 32 'LeerXml("Modelo")
           '   frmPrincipal.FiscalHasar.Puerto = 1  'LeerXml("Puerto")
             ' frmremito.Caption = frmremito.Caption +  " Modelo: "  +
            '  frmPrincipal.FiscalHasar.Comenzar
            
            
 End If
    
    
    
    
 If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket Epson") Then
                   
                    Me.rdT.Value = True
                    
                    With frmPrincipal.FiscalEpson2
                            
                            .ClosePort
                            
                            Call FPDelay
                            
                            .CommPort = LeerXml("Puerto")
                            .BaudRate = 3
                            
                            .ProtocolType = protocol_Extended
                            
                            If (.OpenPort) Then
                                Call FPDelay
                            Else
                                MsgBox "2- El controlador fiscal no está conectado. " + Chr(13) + _
                                "Conecte el controlador y vuelva a ingresar a este módulo"
                            End If

                     End With
 
 
 End If
 
manualcae = 0

TabTotales.SelectedItem = 0


If UCase$(LeerXml("Puesto")) = "ASOCIAL" Then
    Me.opTipoDoc(1).Caption = "Ficha"
End If


If UCase$(LeerXml("Puesto")) = "DIEGO" Or UCase$(LeerXml("Puesto")) = "EMPRESAS2" Then
    Me.rdotros.Value = True
End If


'txtNroInterno = UltimoNroInterno2 + 1
MousePointer = vbDefault
vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)

Me.Caption = "Documento de venta "
Me.Caption = Me.Caption + "      [Nro. de Balance: " + Str(vnrobalance) + "]"
vnroasiento = Val(GenerarDato("SELECT MAX(Numero) AS UAsiento FROM Asientos WHERE NroBalance = " + Str(vnrobalance), "UAsiento")) + 1
    
Me.Caption = Me.Caption + "    [Nro. Asiento: " + Str(vnroasiento) + "]"
    
Me.grd_Choferes.Rows = 0

If Not LeerConfig(17) = "Otros" Then
    Call opTipoDoc_Click(0)
End If

If LeerConfig(30) = "volquete" Then
'......
End If
MousePointer = vbDefault

'txtCliente(0).SetFocus



' vctacorriente
Me.vcventa.Clear
Me.vcventa.Text = "Contado"
Me.vcventa.AddItem ("Contado")
Me.vcventa.AddItem ("Cuenta Corriente")


Me.cboPuntoDeVenta.Clear

Me.cboPuntoDeVenta.AddItem ("0001")
Me.cboPuntoDeVenta.AddItem ("0002")
Me.cboPuntoDeVenta.AddItem ("0003")
Me.cboPuntoDeVenta.AddItem ("0004")
Me.cboPuntoDeVenta.AddItem ("0005")
Me.cboPuntoDeVenta.AddItem ("0006")

Me.cboLetra.Clear

Me.cboLetra.AddItem ("A")
Me.cboLetra.AddItem ("B")
Me.cboLetra.AddItem ("X")



If LeerXml("Puesto") = "CajaComuna" Then
    cboLetra.Text = "X"
    
    Me.vcventa.Text = "Contado"
    Call opTipoDoc_Click(3)
     opTipoDoc(3).Value = True

Else
    'Call opTipoDoc_Click(0)
End If



Dim vsql As String
vsql = "select (select * from configuracion where id = " + Str(Val(Me.vcodEmpresa.Text) + 1) + ") a "


Me.cboPuntoDeVenta.Text = traerDatos2("select * from configuracion", "SucursalDocVenta", PathDBConfig)


Me.vcventa.Text = traerDatos2("select * from configuracion", "cventaDocVenta", PathDBConfig)
'Me.chkfijo.Value = xtpChecked

chkfijo.Value = Val(traerDatos2("select * from configuracion", "FijoDocVenta", PathDBConfig))

vleyenda.Text = (traerDatos2("select * from configuracion", "LeyendaFactura", PathDBConfig))


vncomprobante = NroComprobanteNuevo(TipoDocumento, Trim(cboLetra.Text), Trim(cboPuntoDeVenta.Text))
Me.txtNroComprobante = vncomprobante


Me.chkIncCod.Value = IIf(traerDatos2("select * from configuracion", "IncluyeCodEnDoc", PathDBConfig) = "1", 1, 0)

Me.KlexDetalle.Rows = 1


Me.txtCliente(0).SetFocus


cbolista.Text = 1


On Error GoTo impresora_apag  ' si hay un error en este módulo debe ser de la impresora apagada
    
Procesar:


If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket Hasar") Then
    frmPrincipal.FiscalHasar.ReporteZIndividualPorNumero 13
End If
    
If UCase(LeerXml("Puesto")) = "KIOSCO" Then
    Me.txtCliente(0).Text = "GENERICO"
    Me.txtCliente(0).Tag = 1
    dgClientes.Visible = False
    Call Habilitar(True)
    Me.opTipoDoc(0).Value = True
    Me.txtCliente(3) = "20249182940"
    Me.cboTipoIva = "Consumidor Final"
    Me.txtDetalle(0).SetFocus
    
    Call opTipoDoc_Click(0)
End If
      
   
    
    Exit Sub

impresora_apag:

    'If MsgBox("Error F:" & Err.Description + " -- " + Str(Err.Number), vbRetryCancel, "Errores") = vbRetry Then
        'Resume Procesar
          Me.txtDetalle(0).SetFocus
        Exit Sub
   ' End If
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    vGrabaModo = 0
    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)
    
    'frmPrincipal.CargarMenu
    

    If Err Then GrabarLog "Form_unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GrabarCondicion() ' para el tema de factura
    On Error Resume Next

    Select Case vGrabaModo ' esta variable contiene 1 si se está modificando una factura
    
        Case 1
            'bfactura.Refresh
            'bfactura.Recordset.EditMode
            ' modifica iva venta
        
            ' ----- cristian
            ' Arollback(Index, 1) = "update"
            'bfactura.Refresh
            'bfactura.Recordset.MoveFirst
            With bfactura
                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "SELECT * FROM Factura WHERE (remito = " & Trim(txtNroRemito.Text) & ")"
                .Refresh
            
                If .Recordset.EOF = True Then
                    MsgBox "Error al querer modificar la factura seleccionada", vbInformation
                    GrabarLog "GrabarCondicion", Err.Number & " " & Err.Description, Me.Name
                    Exit Sub
                End If
            
            End With
            '---------------------------------------
            ' bivaventa.Refresh
            ' bivaventa.Recordset.Find("clave = " + v(6))

            'If bivaventa.Recordset.EOF Then Exit Sub
            
            '----------------------------------------
            
            'grabaivaventa
        
        Case Else
    
            ' ----- cristian
            ' Arollback(Index, 1) = "delete"
        
            ' actualiza iva venta
            
            '-----------------01/08/2007
            'bivaventa.Refresh
            'bivaventa.Recordset.AddNew
            '-----------------
            'grabaivaventa
            With bfactura
                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "SELECT * FROM Factura WHERE 1=2"
                .Refresh
                .Recordset.AddNew
            End With
    End Select

    If Err Then GrabarLog "GrabarCondicion", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function GuardarFactura() As Boolean  ' ema: guardar factura
    On Error Resume Next
    
    With bfactura
        If vGrabaModo = 1 Then
            .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM Factura WHERE (remito = " & vremito & ")"
            .Refresh

            If Not .Recordset.EOF = True Then
                .Recordset.MoveFirst
            Else
              '  MsgBox "La factura no fue Actualizada!!", vbExclamation, "Mensaje ..."
                Exit Function
            End If
            '.Recordset.Edit
        Else
            .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM Factura WHERE 1=2"
            .Refresh

            .Recordset.AddNew

            
        End If
        
        
        
        Dim vcoment2 As String
        
        
        If opTipoDoc(2).Value Then
    
            vcoment2 = " > Saldo Anterior: " + Format((lblsaldocliente2), "###,###,##0.00") + "        > Total Doc. : " + Format(Val(Me.txtTotal), "###,###,##0.00") + "       > Saldo Actual: " + Format(Val(Me.lblsaldocliente.Caption) - Val(Me.txtTotal), "###,###,##0.0")
 
        Else
    
            vcoment2 = " > Saldo Anterior: " + Format((lblsaldocliente2), "###,###,##0.00") + "        > Total Doc. : " + Format(Val(Me.txtTotal), "###,###,##0.00") + "       > Saldo Actual: " + Format(Val(Me.lblsaldocliente.Caption) + Val(Me.txtTotal), "###,###,##0.0")
        End If
        
        .Recordset("comentario").Value = vcoment2
        
        
        
        .Recordset("Letra").Value = UCase(Trim(cboLetra.Text))
        
        
        .Recordset("Tipo").Value = TipoDocumento
        .Recordset("TipoMovimiento").Value = Trim(txtTipoMovimiento(0).Text)
        .Recordset("NComprobante").Value = (Val(Trim(txtNroComprobante.Text)))
        
        
        .Recordset("PuntoDeVenta").Value = Trim(cboPuntoDeVenta.Text)
        .Recordset("cae").Value = Trim(Me.vcae.Text)
        .Recordset("caevto").Value = Trim(vcaeFecha)
        
        .Recordset("saldos").Value = Val(lblsaldocliente.Caption)
        
        .Recordset("NroFactNC").Value = Format(Me.vLetraNotaC + "-" + Me.vPuntoDeVentaNotaC + "-" + Me.vNroFacturaNotaC, "#################")

        .Recordset("Fecha").Value = dtpFecha.Value
        .Recordset("FechaIVA").Value = vFechaIva.Value
        
        .Recordset("Hora").Value = Time
        .Recordset("codigo_Num").Value = Val(txtCliente(0).Tag)
        .Recordset("codigo").Value = Trim(txtCliente(0).Tag)
        .Recordset("nombre").Value = txtCliente(0).Text
        .Recordset("domicilio").Value = txtCliente(1).Text
        .Recordset("Cod_Repartidor").Value = txtEmpleados(0).Text
        .Recordset("repartidor").Value = txtEmpleados(1).Text
        .Recordset("Localidad").Value = txtCliente(2).Text
        .Recordset("Telefono").Value = "" 'txtCliente(3).Text
        .Recordset("Iva").Value = Trim(cboTipoIva.Text)
        .Recordset("Cuit").Value = txtCliente(3).Text
        
        .Recordset("nroremito").Value = Me.vnroremito2
        .Recordset("cventa").Value = Me.vcventa
        
        
        
        '----------- ema: campos nuevos para volquete
        .Recordset("cantidadvolquetes").Value = Val(Me.txt_vcantidadVolquete)
        .Recordset("idlistachoferes").Value = Trim(fcomoponerlistachoferes)  ' ema:ale: se encarga de armar la lista de la id de choferes a partir de un grid

        'Dim vestadodocumento As EstadoDocumento

        .Recordset("estadodocumento").Value = "Adeudado"
        
        
        .Recordset("tipopedido").Value = "No Retirado"
        
        '-----------------
                
                
        ' ------------ datos del remito -----------------
        
        .Recordset("Recibio").Value = Me.vRemitoRecibio
        .Recordset("TransportistaNombre").Value = Me.vTransportistaNombre
        .Recordset("TransportistaCuit").Value = Me.vTransportistaCuit
        .Recordset("TransportistaDomicilio").Value = Me.vTransportistaDomicilio
        .Recordset("lentrega").Value = Me.vlentrega
        '.Recordset("comentario").Value = Me.vobservacion
            
        '-------------------------------------------------
                
        vobservacion = Me.txtObservaciones
                
        If Not vGrabaModo = 1 Then .Recordset("remito").Value = Val(txtNroRemito.Text)
        
        If opOtrosDocumentos.Value = True Then
            .Recordset("Subtotal").Value = Val(txtIB(0).Text)
            .Recordset("total").Value = Val(txtIB(10).Text)
            .Recordset("Comentario").Value = Left(Trim(txtIB(7).Text), 100)
        Else
            .Recordset("subtotal").Value = Val(txtSubtotal.Text)
            .Recordset("total").Value = Val(txtTotal.Text)
            .Recordset("descuento").Value = Val(txtDescuento.Text)
            .Recordset("Impuesto").Value = Val(txtImpuesto.Text)
        
            If Trim(txtObservaciones.Text) = "" Then
            '    .Recordset("Comentario").Value = "" 'Trim(lblTipoDocumento.Caption) & " " & Trim(lblNroDocumento.Caption)
            Else
             '   .Recordset("Comentario").Value = Left(Trim(txtObservaciones.Text), 100)
            End If
         
        End If
        
        'Pasaron a la Tabla 'IvaFacturaVenta
        '.Recordset("tiva").Value = Val(txtIva(1).Text)
        '.Recordset("tiva2").Value = Val(txtIva(0).Text)
         
         .Recordset("NroInterno").Value = Val(txtNroInterno.Text)
         
         vnrointerno = Val(txtNroInterno.Text) ' ojo
         
         If Trim(cboBienesServicios.Text) = "Bienes" Then
            .Recordset("BienesServicios").Value = "B"
         Else
            .Recordset("BienesServicios").Value = "S"
         End If
         
         'PANIC
         '.Recordset("nrofactnc").Value = vnrofactnc
         
        
        vTipoDocumento = TipoDocumento
        
        vIdFactura = Val(.Recordset("idFactura").Value)
        
        'Me.bfactura.Recordset.Update
        
        
        
        ' Dim vcoment2 As String
        
        Dim vmarcaCantidad As String
        
        If Me.KlexDetalle.Rows - 1 < 18 Then
        .Recordset("Cod_repartidor").Value = "1"
        Else
        .Recordset("Cod_repartidor").Value = "2"
        End If
        
        
        
        'vcoment2 = " > Saldo Anterior: " + Format(Val(Me.lblsaldocliente.Caption), "###,###,##0.00") + "        > Total Fact. : " + Format(Val(Me.txtTotal), "###,###,##0.00") + "       > Saldo Actual: " + Format(Val(Me.lblsaldocliente.Caption) + Val(Me.txtTotal), "###,###,##0.0")

           
        If opTipoDoc(2).Value Then
    
            vcoment2 = " > Saldo Anterior: " + Format(Val(Me.lblsaldocliente.Caption), "###,###,##0.00") + "        > Total Doc. : " + Format(Val(Me.txtTotal), "###,###,##0.00") + "       > Saldo Actual: " + Format(Val(Me.lblsaldocliente.Caption) - Val(Me.txtTotal), "###,###,##0.00")
 
        Else
    
            vcoment2 = " > Saldo Anterior: " + Format(Val(Me.lblsaldocliente.Caption), "###,###,##0.00") + "        > Total Doc. : " + Format(Val(Me.txtTotal), "###,###,##0.00") + "       > Saldo Actual: " + Format(Val(Me.lblsaldocliente.Caption) + Val(Me.txtTotal), "###,###,##0.00")
        End If
        


    .Recordset("comentario").Value = vcoment2
        
        
        
     Me.bfactura.Recordset.Update
        
    End With
    
    Call GuardarIva
    
    'Call GuardarOtro(Me.txtTipoMovimiento(0).Text)
    
 '  Me.bfactura.Recordset.Update
    
    
    If Err Then
        MsgBox "La factura no fue guardada correctamente.", vbCritical, "Error..."
        GrabarLog "GuardarFactura", Err.Number & " " & Err.Description, Me.Name
        checksum(1) = False
    Else
        checksum(1) = True
        GuardarFactura = True
    End If
   
End Function

Function Validar() As Boolean
Validar = True

' Validar datos de Notas de Credito
If Me.vNroFacturaNotaC = 0 And Me.opTipoDoc(2) Then

   If MsgBox("No ha seleccionado una factura para generar la Nota de Credito. Quiere seguir ?", vbYesNo) = vbYes Then
        Validar = True
    Else
        Validar = False
    End If

End If

If Abs(vtotal_global - CDbl(Me.txtSubtotal)) > 0.1 Then
        MsgBox ("No coincide el total de la factura con los items ingresados")
    Validar = False
End If

' --------------------------------------


End Function

Private Sub GeneraNuevoDocumento()
If Me.vGrabaModo = 1 Then
    
        If vduplicando Then
             Me.vGrabaModo = 0
             txtNroRemito.Text = NroRemitoNuevo
             txtNroInterno = UltimoNroInterno2
             vduplicando = False
        End If
End If
       ' If MsgBox("Confirma .", vbYesNo) = vbYes Then
       '     Me.vGrabaModo = 0
       '      txtNroRemito.Text = NroRemitoNuevo
       '      txtNroInterno = UltimoNroInterno2 + 1
       ' Else
       '     Me.vGrabaModo = 1
       ' End If
    
       ' txtNroInterno = UltimoNroInterno2 + 1
    'End If
End Sub

Private Function Guardar() As Boolean
    On Error Resume Next
    
    If Not Validar() Then
        Guardar = False
        Exit Function
    End If
    Call GeneraNuevoDocumento ' pregunto si quiere grabar un nuevo documento o quiere reemplazar el anterior
    
        
    If ((KlexDetalle.Rows = 2 And chkTotalManual.Value = 0 And KlexDetalle.TextMatrix(1, 5) = "") And (opOtrosDocumentos.Value = False)) And (Not TipoDocumento = "Documento") Then
        MsgBox "Debe cargar detalles para poder GRABAR el Documento", vbExclamation, "Mensaje ..."
        Guardar = False
        Exit Function
    End If
    
    If Trim(txtCliente(0).Tag) = "" Then
        MsgBox "Debe Ingresar un cliente", vbExclamation, "Mensaje ..."
        Guardar = False
        Exit Function
    End If
    
    If Not opOtrosDocumentos.Value = True Then
        If chkTotalManual.Value = 0 Then
            CalcularTotales ' calcular los totales nuevamente
        End If
    End If
    
    vremito = Val(txtNroRemito.Text)  ' paso el nro de remito a la variable vremito
        
                
    GuardarFactura ' paso 2
    'Guardo la operacion en la cta cte
    
    
    If Not (Me.opTipoDoc(1).Value Or Me.opTipoDoc(4).Value) And Not Me.chkReingresarFact = 1 Then   ' no tienen que ser presupuesto ni remito
        WCtaCte (vremito)
    End If
    
    If vNoSaveDoc = True Then
        NuevoCliente
        vNoSaveDoc = False
        Exit Function
    End If
    
    UltimaVenta
        
    If chkTotalManual.Value = 0 Then
        If Not opOtrosDocumentos.Value = True Then
            ConfirmarDetalle
        End If
    End If
    
    vGrabaModo = 0
    cargando = 0
    
    If Err < 0 Then  ' paso xx
        MsgBox "La factura no fue cargada correctamente. Revisar las operaciones!", vbCritical, "Error..."
        GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name
        vGrabaModo = 0
        checksum(3) = False
    Else
        Guardar = True
        checksum(3) = True
    End If

End Function
Private Function NroRemito() As Long
    On Error Resume Next

    Dim rsNroRemito As New ADODB.Recordset, sqlNroRemito As String
    
    sqlNroRemito = "SELECT * FROM Factura ORDER BY Remito DESC"

    With rsNroRemito
        .CursorLocation = adUseClient
        
        Call .Open(sqlNroRemito, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            NroRemito = Val(.Fields("Remito").Value) + 1
        Else
            NroRemito = 1
        End If
    
    End With
    
    sqlNroRemito = ""
    
    If rsNroRemito.State = 1 Then
        rsNroRemito.Close
        Set rsNroRemito = Nothing
    End If
    
    If Err Then GrabarLog "NroRemito", Err.Number & " " & Err.Description, Me.Name
End Function

Private Function NroRemitoNuevo2() As Long
    On Error Resume Next

    Dim vnro As Integer
    Dim vsql As String
    
    
    vnro = TraerDato2("select * from t_nroremito order by numero desc", "numero", pathDBMySQL)
    
    vsql = "insert into t_nroremito (numero) values (" + Str(vnro + 1) + ")"
    
    Call EjecutarScript(vsql, pathDBMySQL)
    
    NroRemitoNuevo2 = vnro + 1
    
    If Err < 0 Then
            MsgBox "Cuidado!", vbCritical
            NroRemitoNuevo2 = 0
            Exit Function
    End If
End Function

Private Sub GuardarCliente()
    On Error Resume Next

    Dim rsNuevoCliente As New ADODB.Recordset, sqlNuevoCliente As String, vlocalidad As String
    
    sqlNuevoCliente = "SELECT * FROM Clientes WHERE 1=2"
    
    With rsNuevoCliente
        .CursorLocation = adUseClient
        Call .Open(sqlNuevoCliente, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        .AddNew
        
        If Not Trim(txtCliente(2).Text) = "" Then
        
            If Not Trim(TraerDato("Localidades", "Localidad LIKE '%" & Trim(txtCliente(2).Text) & "%'", "CodigoPostal")) = "" Then
                .Fields("CodigoPostal").Value = Trim(TraerDato("Localidades", "Localidad LIKE '%" & Trim(txtCliente(2).Text) & "%'", "CodigoPostal"))
                .Fields("Localidad").Value = Trim(TraerDato("Localidades", "Localidad LIKE '%" & Trim(txtCliente(2).Text) & "%'", "Localidad"))
                .Fields("Provincia").Value = Trim(TraerDato("Localidades", "Localidad LIKE '%" & Trim(txtCliente(2).Text) & "%'", "Provincia"))
            End If
        
        Else
            
        End If
        .Fields("codigo") = txtCliente(0).Tag
        .Fields("CodigoNum").Value = Val(txtCliente(0).Tag)
        
        .Fields("Nombre") = Mid(txtCliente(0).Text, 1, 255)
        .Fields("RazonSocial") = Mid(txtCliente(0).Text, 1, 255)
        
        .Fields("Direccion") = Trim(txtCliente(1).Text)
        .Fields("Telefono") = ""
        
        .Fields("idTipoIva") = Trim(cboTipoIva.Tag)
        .Fields("Cuit") = EsNulo(txtCliente(3).Text)

        
        .Fields("CreditoMax").Value = 0
        
        .Fields("Fecha_Alta").Value = Date
        .Fields("Observaciones").Value = "CargadoPorRemito"
        
        .Update
    
    End With
    

    If Err Then
        GrabarLog "GuardarCliente", Err.Number & " " & Err.Description, Me.Name
    Else
        MsgBox "Los datos del cliente fueron guardados", vbInformation, "Mensaje ..."
    End If
End Sub
Public Sub GuardarDoc()
    On Error Resume Next
    
    ReDim checksum(4)
    
    txtEmpleados(0).BackColor = vbWhite
    
    MousePointer = vbHourglass
    
    'CalcularTotales

    checksum(0) = True
    
       If Guardar = True Then
    
        If vNoSaveDoc = True Then
            MousePointer = vbDefault
            Exit Sub
        Else
            
        End If

        If vConfigGral.vIncluyeContabilidad = True Then
            With frmAsientosAlta
                .Show
                .chkControlar.Value = xtpChecked
                .txtCuentaVieneDe.Text = Me.Caption
                .txtCuentaVieneDe.Tag = txtCliente(0).Tag
                .txtLeyenda.Text = vLeyendaAsiento
                .dtpFecha.Value = dtpFecha.Value
                .txtImporteVieneDe.Text = vTotalAsiento
                .lblNroInterno.Caption = Val(txtNroInterno.Text)
                .cboTipoMovimiento.Tag = UCase(Trim(txtTipoMovimiento(0).Text))
                .cboTipoMovimiento.Text = Trim(txtTipoMovimiento(1).Text)
                
                vTotalAsiento = 0
        
                .vVieneTabla = "Factura"
                .vVieneIdNombre = "idFactura"
                .vVieneIdValor = vIdFactura
                
                ' ---------------- mas datos del asiento -----------
                .vtipomovimiento = Me.txtTipoMovimiento(0).Text
                .vCodigoCliente = Me.txtCliente(0).Tag
                .vCodigoProveedor = ""
                .vnrointerno = Val(Me.txtNroInterno)
                '----------------------------------------------------
                
            End With
        End If
    End If
    
    'RecargarForm
    
    ZOrder (1)

    If Err < 0 Then
        MsgBox "Error! Revisar las operaciones... : " & Trim(Err.Description), vbCritical
        'GrabarLog "GuardarDoc", Err.Number & " " & Err.Description, Me.Name
    End If

    MousePointer = vbDefault

End Sub
Private Sub Imprimir()
    On Error Resume Next

    If Not UCase(LeerXml("Puesto")) = "DIEGOsacado" Then
        Call guardarFdetalleTemp
    End If

    If Trim(txtCliente(0).Tag) = "" Then
        MsgBox "Debe ingresar un cliente", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    'DatosEmpresa ' actualiza los datos de la empresa
    
    ReDim checksum(4)
    
    If Guardar = True Then  ' paso 1  ' acá donde se guarda el doc
    
        MousePointer = vbHourglass

        vGrabaModo = 1
        
        
            If UCase(LeerXml("Puesto")) = "DIEGOsaccado" Then
                    setFacturaDatosRemito (Val(txtNroRemito.Text))
                    Call llenarDetallesRemito(Val(txtNroRemito.Text), Me.bdetalle)
                    setMarcarImpresaRemito (Val(txtNroRemito.Text))
                    
                    Call mostrar_Doc2Remito(vTipoDocumento, vncomprobante)
                    MousePointer = vbDefault
                    NuevoCliente
                    Exit Sub
            End If

        'If Not MsgBox("¿ Esta seguro que desea imprimir este Documento ? ", vbYesNo + vbInformation, "Mensaje ...") = vbYes Then Exit Sub

        'bdetalle.Refresh

        Dim i, t As Integer
            
        't = Val(EsNulo(bdetalle.Recordset.RecordCount))

        t = getRecordCount()
        
        Dim vvsql As String
        
        
        vvsql = "delete from Relleno where remito=" + txtNroRemito.Text '+ " where IdRelleno=1"
       Call EjecutarScript(vvsql, pathDBMySQL)
        
        
        vvsql = "insert into Relleno (remito) values (" + txtNroRemito.Text + ")" ' " where IdRelleno=1"
       
       ' vvsql = "UPDATE Relleno SET remito=" + txtNroRemito.Text + " where IdRelleno=1"

        'Call EjecutarScript("UPDATE Relleno SET remito = " & Val(txtNroRemito.Text) & "")
        Call EjecutarScript(vvsql, pathDBMySQL)

        'margenfactura = (18 - t) * 230
        
        If Not LeerXml("MostrarSaldoEnDoc") = "SI" Then margenfactura = (30 - t) * 208

        With Mantenimiento.rscfact
            If Not .State = 0 Then .Close
        
          .Source = "SHAPE {SELECT * FROM Factura WHERE remito = " & Val(txtNroRemito.Text) & "}  AS cfact APPEND ((SHAPE {SELECT * FROM relleno}  AS crelleno APPEND ({SELECT FDetalle.*,relleno.* FROM relleno,FDetalle WHERE (fdetalle.remito = relleno.remito) AND (fdetalle.remito =" & Val(txtNroRemito.Text) & ") ORDER BY idFDetalle ASC} AS cdetalle RELATE 'remito' TO  PARAMETER 0) AS cdetalle) AS crelleno RELATE 'Remito' TO 'remito') AS crelleno"

         
         '.Source = "SHAPE {SELECT * FROM Factura}  AS cfact APPEND (( SHAPE {SELECT * FROM `relleno`}  AS crelleno APPEND ({SELECT fdetalle.*,relleno.* FROM relleno,fdetalle WHERE fdetalle.remito = relleno.remito}  AS cdetalle RELATE 'remito' TO 'Remito') AS cdetalle) AS crelleno RELATE 'Remito' TO 'remito') AS crelleno"
            If Not .State = 1 Then .Open
            .Close
            .Open
        End With
        
        NuevoCliente
        
        
        Unload Mantenimiento
        Load Mantenimiento
    
        'Me.WindowState = vbMinimized
        
        Select Case vTipoDocumento
    
            Case "Fact A"
                'Call ifacta.Hide
                'Call ifacta.PrintReport(False, rptRangeAllPages)
                
                mostrar_ifacta
                
               ' ifacta.Show
                
                'Call ImprimirTicket(Mantenimiento.rscfact.Fields("Remito").Value)
                

            Case "Fact B"
               ' imonotributo.Show
                 mostrar_ifacta
                 'mostrar_ifactb
                 'ifacta.Show

            Case "Presupuesto"
                'ipresupuesto.Show
                
                mostrar_documentos
                'idocumento.Show

            Case "Remito"
                 mostrar_remito
            
            Case "Nota C"
                'mostrar_documentos (iNotaCredito)
                
                'mostrar_iNotaCredito
                'inotac.Show
                 mostrar_ifacta
    
            Case "Documento"
                If f5 = 1 Then
                    f5 = 0
                    
                    MsgBox "Prepare la impresora !", vbInformation, "Mensaje ..."
        
                    With idocumento
                        .Hide
                        '.PrintReport False, rptRangeAllPages
                        .PrintReport False

                    End With
                    Me.WindowState = vbNormal
                Else
                    'Call idocumento.PrintReport(False, rptRangeAllPages)
                    mostrar_documentos
                    'idocumento.Refresh
                    
                End If
        
            Case Else
                mostrar_ifacta
                
                If f5 = 1 Then
                    f5 = 0
                    ifacta.PrintReport False
                    Unload ifacta
                Else
                    mostrar_ifacta
                End If

        End Select
    
        MousePointer = vbDefault
    
        'Me.WindowState = 1

    End If
 
   ' NuevoCliente

If Err Then GrabarLog "Imprimir", Err.Number & " " & Err.Description, Me.Name
End Sub


Function getRecordCount() As Long
On Error Resume Next

If UCase(LeerXml("Puesto")) = "EMPRESAS" Then
    getRecordCount = 0
    Exit Function
End If

bdetalle.RecordSource = "select * from fdetalle where remito=" + Str(Val(txtNroRemito.Text)) + " and codigo is not  null "
bdetalle.Refresh


If bdetalle.Recordset.EOF Then Exit Function
getRecordCount = bdetalle.Recordset.RecordCount

If Err Then
    getRecordCount = 0
    Exit Function
End If

End Function
Private Sub mostrar_documentos()
Dim i As Integer

With idocumento
'----------- titulos -------
.Sections("titulos").Controls("enroremito").Caption = Str(vnrocomprobante)
'.Sections("titulos").Controls("ecventa").Caption = Me.vcventa
.Sections("titulos").Controls("t1").Caption = UCase(vtipoDocDescript)

.Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
.Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
.Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
.Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
.Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)

'---------------------------

.Sections("totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
.Sections("totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
.Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")


If Me.cboTipoIva = "Iva Responsable Inscripto" And Not vtipoDocDescript = "Documento" Then

'If UCase(LeerXml("cliente")) = UCase("diego") Then
                    .Sections("Totales").Controls("eiva105").Visible = True
                    .Sections("Totales").Controls("eiva21").Visible = True
                    .Sections("Totales").Controls("eiva27").Visible = True
                    
                    .Sections("Totales").Controls("etiva10").Visible = True
                    .Sections("Totales").Controls("etiva21").Visible = True
                    .Sections("Totales").Controls("etiva27").Visible = True
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
End If
                    
                    
If UCase(LeerXml("Cliente")) = "PONS" Then
    .Sections("Totales").Controls("esaldodeudor").Caption = "S. deudor: " + Format(vsaldodeudor, "#,###,##0.00")
End If


End With



End Sub
Private Sub mostrar_iNotaCredito()

With iNotaCredito

' mustro el nro de factura correspondiente a la NC
.Sections("cabecera").Controls("eNrofactura").Caption = Format(vLetraNotaC + " - " + vPuntoDeVentaNotaC + " - " + Str(vNroFacturaNotaC), "#####################")

.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
.Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")

.Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
.Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
.Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")

.Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")

'.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")

.Show
End With
End Sub


Private Sub instancia_iafacta(ByRef vdr As ifacta)

If vestructura = 0 Then Set vdr = New ifacta

If vestructura = 1 Then Set vdr = ifacta1

If vestructura = 2 Then Set vdr = New ifacta2

If vestructura = 3 Then Set vdr = New ifacta3

If vestructura = 4 Then Set vdr = New ifacta4

If vestructura = 5 Then Set vdr = New ifacta5

If vestructura = 6 Then Set vdr = New ifacta6

If vestructura = 7 Then Set vdr = New ifacta7

If vestructura = 8 Then Set vdr = New ifacta8

If vestructura = 9 Then Set vdr = New ifacta9

End Sub

Private Sub llearDocumentos()

'If Not UCase(vConfigGral.vEmpresa) = UCase("WgestionPoli") Then

If Not UCase(LeerXml("ObtieneCAE")) = UCase("SI") Or UCase(LeerXml("ObtieneCAE")) = UCase("Prueba") Then



    If vestructura = 0 Then

            With ifacta
                    '----------- titulos -------
                    .Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
                    .Sections("titulos").Controls("ecventa").Caption = Me.vcventa
                    
                    .Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
                    .Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
                    .Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
                    .Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
                    .Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
                    .Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
                    
                    '-------------cae-----------
                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae
                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
                    '--------------------------
                    
                    .Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                    .Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
                    '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                    
                    '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
            End With

        End If



    If Left(vestructura, 2) = "01" Then '  Or vestructura = 2 Or vestructura = 9 Or vestructura = 10 Then

            With ifacta1
                    '----------- titulos -------
                    .Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
                    .Sections("titulos").Controls("ecventa").Caption = Me.vcventa
                    
                    .Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
                    .Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
                    .Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
                    .Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
                    .Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
                    .Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
                    
                    '-------------cae-----------
                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae
                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
                    '--------------------------
                    
                    .Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                    .Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
                    '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                    
                    '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
            End With

        End If


    If Left(vestructura, 2) = "05" Then  ' Or vestructura = 6 Then

            With ifacta5
                    '----------- titulos -------
                    .Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
                    .Sections("titulos").Controls("ecventa").Caption = Me.vcventa
                    
                    .Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
                    .Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
                    .Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
                    .Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
                    .Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
                    .Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
                    
                    '-------------cae-----------
                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae
                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
                    '--------------------------
                    
                    .Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                    .Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
                    '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                    
                    '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
            End With

        End If

    If Left(vestructura, 2) = "11" Then '  Or vestructura = 4

            With ifacta4
                    '----------- titulos -------
                    .Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
                    .Sections("titulos").Controls("ecventa").Caption = Me.vcventa
                    
                    .Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
                    .Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
                    .Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
                    .Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
                    .Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
                    .Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
                    
                    '-------------cae-----------
                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae
                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
                    '--------------------------
                    
                    .Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                    .Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
                    '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                    
                    '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
            End With

        End If

    If Left(vestructura, 2) = "07" Then '  Or vestructura = 2 Or vestructura = 9 Or vestructura = 10 Then

            With ifacta7
                    '----------- titulos -------
                    .Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
                    .Sections("titulos").Controls("ecventa").Caption = Me.vcventa
                    
                    .Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
                    .Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
                    .Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
                    .Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
                    .Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
                    .Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
                    
                    '-------------cae-----------
                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae
                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
                    '--------------------------
                    
                    .Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                    .Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
                    '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                    
                    '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
            End With

        End If



If Left(vestructura, 2) = "08" Then

            With ifacta8
                    '----------- titulos -------
                    .Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
                    .Sections("titulos").Controls("ecventa").Caption = Me.vcventa
                    
                    .Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
                    .Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
                    .Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
                    .Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
                    .Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
                    .Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
                    
                    '-------------cae-----------
                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae
                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
                    '--------------------------
                    
                    .Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                    .Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
                    '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                    
                    '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
            End With

        End If



    If Left(vestructura, 2) = "12" Then

            With ifacta12
                    '----------- titulos -------
                    .Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
                    .Sections("titulos").Controls("ecventa").Caption = Me.vcventa
                    
                    .Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
                    .Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
                    .Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
                    .Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
                    .Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
                    .Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
                    
                    '-------------cae-----------
                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae
                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
                    '--------------------------
                    
                    .Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                    .Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
                    '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                    
                    '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
            End With

        End If




    If Left(vestructura, 2) = "03" Then

            With ifacta3
                    '----------- titulos -------
                    .Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
                    .Sections("titulos").Controls("ecventa").Caption = Me.vcventa
                    
                    .Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
                    .Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
                    .Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
                    .Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
                    .Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
                    .Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
                    
                    '-------------cae-----------
                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae
                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                    .Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
                    '--------------------------
                    
                    .Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
                    .Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
                    .Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")
                    
                    .Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
                    '.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")
                    
                    '.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
            End With

        End If


End If


'If UCase(vConfigGral.vEmpresa) = UCase("WgestionPoli") Then


If UCase(LeerXml("ObtieneCAE")) = UCase("SI") Or UCase(LeerXml("ObtieneCAE")) = UCase("Prueba") Then

With ifactaPoli



' --- control datos afip ---
' 2021-06-13, lo saco de acá porque si sale el cartel me arruna la impresión. No carga los datos
'Call control_afip_factura
'----------- titulos -------

.Sections("titulos").Controls("eDocumento").Caption = vetipoFactura
.Sections("titulos").Controls("eCodigo").Caption = vtipoFactura
.Sections("titulos").Controls("eLetra").Caption = UCase(vLetra)
.Sections("titulos").Controls("ePtoVta").Caption = Format(vPtoVta, "0000")
.Sections("titulos").Controls("eNro").Caption = Format(vnrocomprobante2, "00000000")
.Sections("titulos").Controls("eiva").Caption = cboTipoIva.Text
Debug.Print "nro comprobante : " + .Sections("titulos").Controls("eNro").Caption

.Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
.Sections("titulos").Controls("ecventa").Caption = Me.vcventa

.Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
.Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
.Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
.Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
.Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
'---------------------------

'-------------cae-----------
.Sections("Totales").Controls("lblCAE").Caption = Me.vcae
.Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

.Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
.Sections("Totales").Controls("lblVtoCAE").Caption = vcaeFecha
'--------------------------

If UCase(LeerXml("Cliente")) = "PONS" Then
    .Sections("Totales").Controls("esaldodeudor").Caption = "S. Deudor: " + Format(vsaldodeudor, "#,###,##0.00")
End If

.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")
.Sections("Totales").Controls("esubtotal").Caption = Format(vgTsubtotal, "#,###,##0.00")

If vtipoFactura > 5 Then ' saco los subtotales deiva

.Sections("Totales").Controls("eiva105").Visible = False
.Sections("Totales").Controls("eiva21").Visible = False
.Sections("Totales").Controls("eiva27").Visible = False

.Sections("Totales").Controls("etiva10").Visible = False
.Sections("Totales").Controls("etiva21").Visible = False
.Sections("Totales").Controls("etiva27").Visible = False
.Sections("Totales").Controls("ttiva").Visible = False

End If

.Sections("Totales").Controls("eiva105").Caption = Format(vgTiva105, "#,###,##0.00")
.Sections("Totales").Controls("eiva21").Caption = Format(vgTiva21, "#,###,##0.00")
.Sections("Totales").Controls("eiva27").Caption = Format(vgTiva27, "#,###,##0.00")

.Sections("Totales").Controls("edescuento").Caption = Format(vgTPdescuento, "#,###,##0.00")
.Sections("Totales").Controls("eleyenda").Caption = Me.vleyenda.Text



If UCase(LeerXml("Puesto")) = UCase("Poliwheel") Then  ' cambiar or true

    .Sections("titulos").Controls("logo2").Visible = True
    .Sections("titulos").Controls("eddesc").Visible = False
    .Sections("detalle").Controls("tddesc").Visible = False
    .Sections("Totales").Controls("etdescuent").Visible = False
    .Sections("Totales").Controls("edescuento").Visible = False
End If


'.Sections("Totales").Controls("eimpuesto").Caption = Format(vgTimpuesto, "#,###,##0.00")

'.Sections("Totales").Controls("etotal").Caption = Format(vgTtotal, "#,###,##0.00")

End With
End If
End Sub


Private Sub control_afip_factura()
    
    
    ' si estoy en el modo reingreso de factura no hago el control
    ' si el nro de factura en afip es el mismo
    
    If Me.chkReingresarFact Then
        Exit Sub
    End If
    

    Dim vtf As Boolean
    Dim mensaje As String
    
    vtf = Trim(vnrocomprobante2) = Trim(vnrocomprobante_control1)
    
    If Not vtf Then
        
        mensaje = "Erro: Atención !!!!" + _
        Chr(13) + "No coincide el nro de comprobante de AFIP" + _
        Chr(13) + "Nro AFIP " + vnrocomprobante_control1 + _
        Chr(13) + "Nro Factura " + vnrocomprobante2 + _
        Chr(13) + "El sistema asumirá que el nro de comprobante correct es:" + vnrocomprobante_control1
        
        
        ' guardo el mensaje de error en el log
        Log22 (mensaje)
        
        MsgBox mensaje, vbCritical
        
        vnrocomprobante2 = vnrocomprobante_control1
        
        
        
    End If
    
End Sub


Private Sub mostrar_ifacta()
 
Call control_afip_factura

llearDocumentos

 
'If UCase(vConfigGral.vEmpresa) = UCase("WgestionPoli") Then

If UCase(LeerXml("ObtieneCAE")) = UCase("SI") Or UCase(LeerXml("ObtieneCAE")) = UCase("Prueba") Then

                        With ifactaPoli
           
                                    ' ---- poner los datos del cae ----
                                    .Sections("Totales").Controls("lblCAE").Caption = Me.vcae.Text
                                    .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                                    .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                                    .Sections("Totales").Controls("lblVtoCAE").Caption = Trim(Right(Trim(vcaeFecha), 2) + "/" + Mid(Trim(vcaeFecha), 5, 2) + "/" + Left(Trim(vcaeFecha), 4))
                                    '-----------------------------------
                                    
                                    
                                    If vgTiva105 = 0 Then
                                        .Sections("Totales").Controls("eiva105").Visible = False
                                        .Sections("Totales").Controls("etiva10").Visible = False
                                    End If
                                
                                
                                If vgTiva21 = 0 Then
                                .Sections("Totales").Controls("eiva21").Visible = False
                                .Sections("Totales").Controls("etiva21").Visible = False
                                End If
                                
                                
                                If vgTiva27 = 0 Then
                                .Sections("Totales").Controls("eiva27").Visible = False
                                .Sections("Totales").Controls("etiva27").Visible = False
                                End If
                                
                                .Show
                        End With

End If

'If Not UCase(vConfigGral.vEmpresa) = UCase("WgestionPoli") Then

If Not UCase(LeerXml("ObtieneCAE")) = UCase("SI") And Not UCase(LeerXml("Puesto")) = UCase("Empresas") Then

                With ifacta
                            ' ---- poner los datos del cae ----
                            .Sections("Totales").Controls("lblCAE").Caption = Me.vcae.Text
                            .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

                            .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
                            .Sections("Totales").Controls("lblVtoCAE").Caption = Trim(Right(Trim(vcaeFecha), 2) + "/" + Mid(Trim(vcaeFecha), 5, 2) + "/" + Left(Trim(vcaeFecha), 4))
                            '-----------------------------------
                                
                            If vgTiva105 = 0 Then
                                .Sections("Totales").Controls("eiva105").Visible = False
                                .Sections("Totales").Controls("etiva10").Visible = False
                            End If
                        
                            If vgTiva21 = 0 Then
                                .Sections("Totales").Controls("eiva21").Visible = False
                                .Sections("Totales").Controls("etiva21").Visible = False
                            End If
                            
                            
                            If vgTiva27 = 0 Then
                                .Sections("Totales").Controls("eiva27").Visible = False
                                .Sections("Totales").Controls("etiva27").Visible = False
                            End If
                        
                          .Show
                End With
End If

End Sub


Private Sub mostrar_ifactb()

llearDocumentos

With ifacta

            ' ---- poner los datos del cae ----
            .Sections("Totales").Controls("lblCAE").Caption = Me.vcae.Text
            .Sections("Totales").Controls("enroCodigoBarra").Caption = vCodigoBarra

            .Sections("Totales").Controls("lblCAE2").Caption = vCodigoBarra
            .Sections("Totales").Controls("lblVtoCAE").Caption = Trim(Right(Trim(vcaeFecha), 2) + "/" + Mid(Trim(vcaeFecha), 5, 2) + "/" + Left(Trim(vcaeFecha), 4))
            '-----------------------------------
                
            
            '---------- acomoda datos
            .Sections("Totales").Controls("eiva105").Caption = ""
            .Sections("Totales").Controls("eiva21").Caption = ""
            .Sections("Totales").Controls("eiva27").Caption = ""
            .Sections("Totales").Controls("ttiva").Caption = ""
            
            .Sections("Totales").Controls("etiva10").Caption = ""
            .Sections("Totales").Controls("etiva21").Caption = ""
            .Sections("Totales").Controls("etiva27").Caption = ""
            
            .Sections("titulos").Controls("ettiva").Caption = "IVA excento"
            '--------------------------
            
            .Show
End With

End Sub

Private Sub mostrar_remito()

'llearDocumentos

With drRemito

'----------- titulos -------
.Sections("titulos").Controls("enroremito").Caption = Me.vnroremito2
.Sections("titulos").Controls("ecventa").Caption = Me.vcventa

.Sections("titulos").Controls("enombre").Caption = txtCliente(0).Text
.Sections("titulos").Controls("edomicilio").Caption = txtCliente(1).Text
.Sections("titulos").Controls("elocalidad").Caption = txtCliente(2).Text
.Sections("titulos").Controls("ecuit").Caption = txtCliente(3).Text
.Sections("titulos").Controls("efecha").Caption = Str(dtpFecha)
'---------------------------


'---------- acomoda datos
.Sections("Totales").Controls("erecibio").Caption = Me.vRemitoRecibio
.Sections("Totales").Controls("eTransportistaNombre").Caption = Me.vTransportistaNombre
.Sections("Totales").Controls("eTransportistaCuit").Caption = Me.vTransportistaCuit
.Sections("Totales").Controls("eTransportistaDomicilio").Caption = Me.vTransportistaDomicilio
.Sections("Totales").Controls("elentrega").Caption = Me.vlentrega
.Sections("Totales").Controls("eobservaciones").Caption = vobservacion

.Show
'--------------------------
End With


Me.vRemitoRecibio = ""
Me.vTransportistaNombre = ""
Me.vTransportistaCuit = ""
Me.vTransportistaDomicilio = ""
Me.vlentrega = ""
Me.txtObservaciones = ""


End Sub


Private Sub txtIva_Change(Index As Integer)
    On Error Resume Next
    
    If Not CBool(chkTotalManual.Value) = True Then
        txtIva(Index).Text = Format(txtIva(Index).Text, "#######0.000")
    Else
        
    
    End If
    
    If Err Then GrabarLog "txtIva_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtIva_KeyPress(Index As Integer, _
                         KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        If Index = 0 Then txtIva(1).SetFocus
        If Index = 1 Then txtIva(2).SetFocus
        If Index = 2 Then txtPDescuento.SetFocus
    End If

    If Err Then GrabarLog "txtIva_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtIva_LostFocus(Index As Integer)
    On Error Resume Next
    
    txtIva(Index).Text = Format(txtIva(Index).Text, "#####0.000")

    If Err Then GrabarLog "iva_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub LimpiarBase()
    On Error Resume Next
    
    KlexDetalle.Enabled = False

    LimpiarFDetalle
    
    KlexDetalle.Enabled = True

    If Err Then GrabarLog "limpiabase", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()

    Dim i As Integer

    For i = 0 To txtDetalle.Count - 1
        txtDetalle(i).Text = ""
        txtDetalle(i).Tag = ""
    Next
    
    If fraCargaDetalle.Enabled = True Then
        focoEnLinea
        'txtDetalle(0).SetFocus
    Else
        focoEnLinea
        'txtCliente(0).SetFocus
    End If
    
    vpcosto = 0
    
    
End Sub
Public Sub limpiarCliente()
Dim i As Integer
    
        For i = 0 To txtCliente.Count - 1
            txtCliente(i).Text = ""
            txtCliente(i).Tag = ""
        

        cboTipoIva.Text = ""
        cboTipoIva.Tag = ""
        Next
       dtpFecha.Value = Date
       
       Me.vnroremito2 = ""
       
       If UCase$(LeerXml("Puesto")) = "KIOSCO" Then
                Me.txtCliente(0).Text = "GENERICO"
                Me.txtCliente(0).Tag = 1
                Me.txtCliente(3) = "20249182940"
                Me.cboTipoIva = "Consumidor Final"

                dgClientes.Visible = False
                Call Habilitar(True)

                Me.opTipoDoc(0).Value = True
                'Me.txtDetalle(0).SetFocus
                    Me.txtCliente(3) = "20249182940"
                Me.cboTipoIva = "Consumidor Final"
                Me.txtDetalle(0).SetFocus
                
                Call Me.opTipoDoc_Click(0)
                
       End If
     
       
End Sub

Public Sub LimpiarCampos()
    On Error Resume Next
    Dim i As Integer



vcancelartrans = False

    'dtpFecha.Value = Date

'    If Not CBool(LeerConfig(25)) = True Then
'        For i = 0 To txtCliente.Count - 1
'            txtCliente(i).Text = ""
'            txtCliente(i).Tag = ""
'        Next
'
'        cboTipoIva.Text = ""
'        cboTipoIva.Tag = ""
'
'    End If

    Me.vNroFacturaNotaC = 0
    Me.vLetraNotaC = ""
    Me.vPuntoDeVentaNotaC = ""

    vOpenGrilla(0) = False
    vOpenGrilla(1) = False
    txtSubtotal.Text = ""
    txtIva(0).Text = ""
    txtIva(1).Text = ""
    txtIva(2).Text = ""
    txtTotal.Text = ""
    txtDescuento.Text = ""
    txtImpuesto.Text = ""
    txtObservaciones.Text = ""
    LimpiarFDetalle
    vCantidadControl = 0
    vRemitoControl = 0
    
    txtNroInterno.Text = ""
    
    For i = 0 To 10
        If Not i = 8 And Not i = 9 Then
            txtIB(i).Text = ""
        End If
    Next
    
    cboBienesServicios.Text = ""
    txtTipoMovimiento(0).Text = ""
    txtTipoMovimiento(1).Text = ""
    
    'KlexDetalle.Clear
    KlexDetalle.Rows = 2
    KlexDetalle.Tag = 0
    
    cboPuntoDeVenta.Text = ""
    cboLetra.Text = ""
    txtNroComprobante.Text = ""
    
    lblTipoDocumento.Caption = ""
    
    HabilitarDocAFactura (False)
    

     vcodEmpresa.Tag = ""
   '  vcodEmpresa.Text = ""
   '  vdescEmpresa.Text = ""
     
     vcodRepartidor.Tag = ""
     vcodRepartidor.Text = ""
     vDesRepartidor.Text = ""
    


   ' fraTipoDocumento.Enabled = True
    
    FormatoGrillaDetalle (1)
    
    With bdetalle
        If Not .ConnectionString = "" Then
            If Not .Recordset.EOF Then
                .Refresh
                .Recordset.MoveLast
            End If
        End If
    End With
    
    If Err Then
        GrabarLog "LimpiarCampos", Err.Number & " " & Err.Description, Me.Name
    Else
        checksum(4) = True
    End If
End Sub
Private Sub LimpiarFDetalle()
On Error Resume Next
    
    Call BorrarBase("fdetalle WHERE confirmado = 'N'", pathDBMySQL)
    
If Err Then GrabarLog "LimpiarFDetalle", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub MostrarDetalle()
    On Error Resume Next
    
    'Cargo Codigo, Detalle, Precio, PCosto, TipoIVA
    With barticulo
        txtDetalle(1).Text = EsNulo(.Recordset("descrip").Value)
        txtDetalle(1).Tag = EsNulo(.Recordset("codigo").Value)
        
        txtDetalle(4).Text = TraerDato("PorcentajeIva", "idPorcentajeIva =  '" & EsNuloEntero(.Recordset("idPorcentajeIva").Value) & "'", "Porcentaje")
        
            txtDetalle(2).Text = Val(.Recordset("Pventa" & Val(cbolista.Text)).Value)
            
            vpcosto = Val(.Recordset("Pcosto").Value)
 
        Select Case Trim(cboTipoIva.Tag)
            
            Case "001"
                If opTipoDoc(0).Value = True Or opTipoDoc(2).Value = True Or LeerXml("PrecioConIva") Then
                    'Me meto si es Responsable Inscrito con Fact (A) o Nota C (A)
                    txtDetalle(2).Text = Val(.Recordset("Pventa" & Val(cbolista.Text)).Value)
                Else
                    'Es Responsable pero va otro documento
                    
                    txtDetalle(2).Text = .Recordset("Pventa" & Val(cbolista.Text)).Value + .Recordset("Pventa" & Val(cbolista.Text)).Value * txtDetalle(4).Text / 100
                    'Panic: Preguntar cuando tiene iva o no!!!
                    '* Val("1." & Replace(EsNuloEntero(.Recordset("idporcentajeiva").Value), ".", ""))
                End If
                        
            Case "0002", "0003", "0004", "0005"
                txtDetalle(2).Text = .Recordset("Pventa" & Val(cbolista.Text)).Value + .Recordset("Pventa" & Val(cbolista.Text)).Value * txtDetalle(4).Text / 100
            
            Case Else
                'MsgBox "ESTO ESTA MALLLLLLLLLLL.........." 'txtDetalle(2).Text = Val(.Recordset("Pventa" & Val(cbolista.Text)).Value) * Val("1." & Replace(EsNuloEntero(.Recordset("idporcentajeiva").Value), ".", ""))
        
        End Select
        
        
        If Val(.Recordset("Stock").Value) - Val(txtDetalle(0).Text) < 0 Then
            txtDetalle(0).BackColor = vbRed
            txtDetalle(2).BackColor = vbRed
            txtDetalle(3).BackColor = vbRed
            txtDetalle(3).BackColor = vbRed
            If UCase(LeerXml("stockcero")) = "SI" Then
                MsgBox "No hay STOCK", vbCritical
            End If
        End If
        
        
        If Right(Me.txtDetalle(1).Text, 1) = "*" Then
        
            Me.txtDetalle(2).Text = Val(txtDetalle(2).Text) * 1.21
        
        End If
        
        
        
        'If Not IsNull(.Recordset("TipoIva").Value) = True Then
        '    f(2).Text = .Recordset("Pventa" & Val(cboLista.Text)) * Val("1." & Replace(.Recordset("TipoIva").Value, ".", ""))
        '    vpespecial = False
        'Else
        '    f(2).Text = .Recordset("Pventa" & Val(cboLista.Text)).Value
        '    MsgBox "Cuidado .... el Articulo seleccionado no tiene un valor asignado de IVA", vbExclamation, "Mensaje ..."
        'End If
        
    End With
    
    ElegirTipoPrecio

    If Err Then GrabarLog "MostrarDetalle", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Function EsNuloEntero(n As Variant)
If IsNull(n) Then
    EsNuloEntero = 0
Else
    EsNuloEntero = n
End If
End Function
Private Sub NuevoCliente()
    On Error Resume Next
    
    RecargarForm
    vGrabaModo = 0

    If Err Then GrabarLog "NuevoCliente", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub NuevoDoc()
    On Error Resume Next
    
    'MousePointer = vbHourglass
    
    ban = "1"
    'Form_Load
    
    Call init
    
    'LimpiarCampos
    Limpiar
    txtCliente(0).SetFocus
       
    If Err Then
        MousePointer = vbDefault
        'MsgBox "Error! Revisar las operaciones", vbCritical, "Mensaje ..."
        If Err Then GrabarLog "NuevoDoc", Err.Number & " " & Err.Description, Me.Name
    
        Exit Sub
    End If

    'MousePointer = vbDefault
End Sub
Private Sub PagarArticulo(ByRef rsFDetalle As Recordset, i As Integer)
On Error Resume Next
    
    Dim vrubro As String
    
    With barticulo
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Articulos WHERE (codigo = '" & Trim(rsFDetalle.Fields("Codigo").Value) & "')"
        .Refresh
    
        If Not .Recordset.EOF = True Then
            If IsNull(.Recordset("idRubros").Value) = True Then
                vrubro = ""
            Else
                vrubro = .Recordset("idRubros").Value
            End If
        Else
            vrubro = ""
        End If
    End With
    
    Dim rsArticulosGanancia As New ADODB.Recordset, sqlArticulosGanancia As String
    
    sqlArticulosGanancia = "SELECT * FROM Articulos_Ganancia WHERE (CodEmp = '" & Trim(txtEmpleados(0).Text) & "') AND (CodCli = '" & Trim(txtCliente(0).Tag) & "') AND (CodRub = '" & Trim(vrubro) & "')"
    
    ' ------- busco la ganancia que tiene el artículo ----------------------
    With rsArticulosGanancia
        Call .Open(sqlArticulosGanancia, ConnDDBB, adOpenStatic, adLockPessimistic)
    
        If .EOF = True Then ' en el caso q no tengo asignado rubro, cliente, empleado porcentaje
            If Not barticulo.Recordset.EOF = True Then
                vganancia = Val(Format(barticulo.Recordset("ganancia_vendedor").Value, "#######0.000"))
            Else
                vganancia = 0
            End If
        Else ' en el caso que tenga asignado porcentaje
            vganancia = Val(Format(.Fields("porcentaje").Value, "#######0.000"))
        End If
    End With
    
    sqlArticulosGanancia = ""
    
    If rsArticulosGanancia.State = 1 Then
        rsArticulosGanancia.Close
        Set rsArticulosGanancia = Nothing
    End If
    
    '-----------------------------------------------------------------------
    
    With rsFDetalle
        .Fields("Ganancia").Value = vganancia
    
        'esta linea va dentro del if
        .Fields("Sueldo") = (vganancia * Val(KlexDetalle.TextMatrix(i, 5)) * Val(KlexDetalle.TextMatrix(i, 7))) / 100

        
        'Juan : 2010-07-19
        
        'If (Me.chkTotalContado.Value) Then
            
            .Fields("total_cdo").Value = Val(KlexDetalle.TextMatrix(i, 11))
            .Fields("Pagado").Value = "SI"
            .Fields("Pago").Value = Format(Val(KlexDetalle.TextMatrix(i, 5)) * Val(KlexDetalle.TextMatrix(i, 7)), "######0.000")
            .Fields("resta").Value = "0.000"
        'Else
         '   If Not vGrabaModo = 1 Then
                '.Fields("total_ctacte").Value = Val(KlexDetalle.TextMatrix(i, 11))
                '.Fields("Pagado").Value = "NO"
                '.Fields("resta").Value = Val(KlexDetalle.TextMatrix(i, 11))
         '   Else
                'SI MODIFICA
         '   End If
    
        'End If
    
    End With
    
If Err Then GrabarLog "PagarArticulo", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Pasar(Index As Integer)
    On Error Resume Next

    If Index >= 5 Then
        'If Val(f(6).Text) <= 0 Then
        '    MsgBox "La cantidad y el precio deben ser valores positivos !", vbCritical, "Error..."
        '   Exit Sub
        'End If
        
        
                GrabarRenglon
                vArticuloNuevo = False
                
                Exit Sub

        
        With barticulo
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM Articulos WHERE (codigo = '" & Trim(.Recordset.Fields("Codigo").Value) & "')"
            .Refresh
    
            If Not .Recordset.EOF = True Then
                If opTipoDoc(0).Value = True Or opTipoDoc(2).Value = True Then
                    
                    'NO USAR ESTOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOOO
                    If Trim(cboTipoIva.Tag) = "008888" Then
                        Select Case .Recordset.Fields("idPorcentajeIVA").Value
                               
                         Case 1
                            txtIva(0).Text = Val(txtIva(0).Text) + Val(txtDetalle(6).Text) * 0.105
                            
                         Case 2
                            txtIva(1).Text = Val(txtIva(1).Text) + Val(txtDetalle(6).Text) * 0.21
                    
                         Case 3
                            txtIva(2).Text = Val(txtIva(2).Text) + Val(txtDetalle(6).Text) * 0.27
                    
                        End Select
                
                        End If
                    End If
                End If
        
               ' fraTipoDocumento.Enabled = Not True
                                
                'If vArticuloNuevo = True Then
                '    If Trim(cboLetra.Text) = "A" Then
                '        Call GuardarArticuloNuevo(Trim(txtDetalle(1).Text), txtDetalle(2).Text, Trim(txtDetalle(4).Text))
                '    Else
                '        Call GuardarArticuloNuevo(Trim(txtDetalle(1).Text), Val(txtDetalle(2).Text) / Val("1." & Replace(Val(txtDetalle(4).Text), ".", "")), Trim(txtDetalle(4).Text))
                '    End If
                'Else
                '
                'End If
                GrabarRenglon
                vArticuloNuevo = False
           
         End With
    Else
        Pasar2 (Index)
    End If

    If Err Then
        Exit Sub
    'GrabarLog "Pasar (" & Index & ")", Err.Number & " " & Err.Description, Me.Name
    End If
End Sub
Public Sub Pasar2(Index As Integer)
    'If True Then
    
    If UCase(LeerXml("Puesto")) = "PONS" Then
                'If Trim(vConfigGral.vEmpresa) = Trim("wgestionpons") Then
               Select Case Index
                    Case 0
                        Index = 1
                    Case 1
                        Index = -1
                    Case Is > 1
                    Index = 5
                
                End Select

                    txtDetalle(Index + 1).SetFocus
    End If
    
        txtDetalle(Index + 1).SetFocus
    
End Sub


Public Sub RecargarForm()
    On Error Resume Next
    
    Me.Visible = Not True
    ban = "1" ' panic!! que quiso poner Juan con esto
    LimpiarCampos
    'Form_Load
    
    'limpiarCliente
    
    vGrabaModo = 0
    Me.Visible = True
    txtCliente(0).SetFocus

    If Err Then GrabarLog "RecargarForm", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GrabarRenglonViejo()
    On Error Resume Next
    Dim i As Integer

    With bdetalle
        
        .Recordset.AddNew
        .Recordset("fecha").Value = dtpFecha.Value
        .Recordset("confirmado").Value = "N"
        .Recordset("cantidad").Value = Val(txtDetalle(0).Text)
        .Recordset("detalle").Value = txtDetalle(1).Text
        .Recordset("codigo").Value = txtDetalle(1).Tag
        .Recordset("repartidor").Value = Trim(txtEmpleados(0).Text)
        
        If opTipoDoc(1).Value = True Or Trim(cboTipoIva.Text) = "Consumidor Final" Or Trim(cboTipoIva.Text) = "Responsable Monotributo" Then
            .Recordset("precio").Value = Format(Val(txtDetalle(2).Text), "######0.000")
        Else

            If bienes.Value = 1 Then
                .Recordset("precio").Value = Format(Val(txtDetalle(2).Text) - (Val(txtDetalle(2).Text) * 9.5 / 100), "######0.000")
                .Recordset("tiva").Value = "10.5"
            Else
                .Recordset("precio").Value = Format(Val(txtDetalle(2).Text), "######0.000")
                .Recordset("tiva").Value = TraerDato("Articulos", "Codigo = '" & txtDetalle(1).Tag & "'", "TipoIva")
            End If
        End If
        
        'If venta.Text = "Contado" Then PANIC VOLVER ATRAS CUANDO SE HAGA GENERAL
        .Recordset("Ganancia").Value = vganancia
        .Recordset("sueldo").Value = (vganancia * .Recordset("precio").Value * .Recordset("cantidad").Value) / 100
        'End If
        
        .Recordset("descuento").Value = Val(txtDetalle(3).Text)
        .Recordset("impuesto").Value = Val(txtDetalle(5).Text)
        .Recordset("total").Value = Format(.Recordset("precio").Value * .Recordset("cantidad").Value, "######0.000")

        'Juan: 2010-07-19
        'If cboVenta = "Contado" Then
            .Recordset("total_cdo").Value = Format(.Recordset("precio").Value * .Recordset("cantidad").Value, "######0.000")
        'Else
        '    .Recordset("total_ctacte").Value = Format(.Recordset("precio").Value * .Recordset("cantidad").Value, "######0.000")
        'End If
        
        .Recordset("totaliva").Value = Format(.Recordset("precio").Value * .Recordset("cantidad").Value, "######0.000") * 1.21
        .Recordset("remito").Value = Val(txtNroRemito.Text)
        .Recordset("envase").Value = venvase
    
        'Juan: 2010-07-19
        'If cboVenta = "Contado" Then
            .Recordset("Pagado").Value = "SI"
            .Recordset("Pago").Value = Format(.Recordset("precio").Value * .Recordset("cantidad").Value, "######0.000")
            .Recordset("resta").Value = "0.000"
        'Else
        '    .Recordset("Pagado").Value = 0
        '    .Recordset("Pagado").Value = "NO"
        '    .Recordset("resta").Value = Format(.Recordset("precio").Value * .Recordset("cantidad").Value, "######0.000") * 1.21
        'End If
        
        'MARCAR ACA SI EL ART. PERTENECE A UNA PERSONA CON PRECIO ESPECIAL - LE PASO UNA VARIABLE BOOLEAN QUE SE SETEA EN F(index)
        .Recordset("pespecial").Value = vpespecial
        .Recordset.Update
        
        .Refresh
        
        Limpiar
        
        bdetalle.Refresh
        
        vRemitoControl = Val(txtNroRemito.Text)
        vCantidadControl = vCantidadControl + 1

        CalcularTotales
    End With

    If Err Then GrabarLog "GrabarRenglonViejo", Err.Number & " " & Err.Description, Me.Name
End Sub
Function validarFilasFDtalles() As Boolean
On Error Resume Next

vlineasDetalles.Text = KlexDetalle.Rows

validarFilasFDtalles = True

If KlexDetalle.Rows > 28 And UCase(LeerXml("cliente")) = "PONS" Then
    MsgBox "No se pueden ingresar mas lineas de detalles para este documento", vbInformation
    validarFilasFDtalles = False
End If


If Not Me.txtDetalle(6) > 0 And UCase(LeerXml("puesto")) = "ASOCIAL" Then
    validarFilasFDtalles = False
    MsgBox "No se pueden ingresar mas lineas de detalles para este documento", vbInformation
End If

If Me.txtDetalle(0) = 0 Then validarFilasFDtalles = False

If Err Then Exit Function
End Function

Public Sub GrabarRenglon()
    On Error Resume Next
    Dim i, j As Integer
    
    If Not validarFilasFDtalles Then Exit Sub
    
    Dim vIdTipoIva As String

    vIdTipoIva = TraerDato("Clientes", "Codigo = '" & Me.txtCliente(0).Tag & "'", "idTipoIva")
    
    
    With KlexDetalle
    
        If Val(.Tag) = 0 And .Rows = 2 And .TextMatrix(.Rows - 1, 4) = "" Then
            FormatoGrillaDetalle (1)
            '.Rows = .Rows + 1
            .Tag = 1
            
            
        Else
            .Rows = .Rows + 1
            .Tag = 1 ' borrar esta
        End If
        j = .Rows - 1
        
        If TabTipoDetalle.tab = 0 Then
        
            .TextMatrix(j, 1) = ""
            '.TextMatrix(j, 2) = dtpFecha.Value
            '.TextMatrix(j, 3) = Val(v(6).Text)              '(remito)
            
            If (Me.chkIncCod.Value) Then .TextMatrix(j, 4) = "[" & Trim(txtDetalle(1).Tag) & "]"   '(codigo)
            
            .TextMatrix(j, 5) = Val(txtDetalle(0).Text)     '(cantidad)
            If (Me.chkIncCod.Value) Then
                .TextMatrix(j, 6) = "[" + EsNulo(txtDetalle(1).Tag) + "] " + EsNulo(txtDetalle(1).Text)  '(detalle)
            
            Else
                .TextMatrix(j, 6) = EsNulo(txtDetalle(1).Text)   '(detalle)
            End If
            
            
           .TextMatrix(j, 4) = "[" & Trim(txtDetalle(1).Tag) & "]"   '(codigo)
            
            
            
            .TextMatrix(j, 7) = EsNulo(txtDetalle(2).Text)  '+ EsNulo(txtDetalle(2).Text) * Val(txtDetalle(4).Text) / 100 'P. Venta
            .TextMatrix(j, 8) = EsNulo(txtDetalle(3).Text)  'Descuento
            .TextMatrix(j, 9) = EsNulo(txtDetalle(4).Text)  'Tipo Iva
            .TextMatrix(j, 10) = EsNulo(txtDetalle(5).Text)  'Impuesto
            .TextMatrix(j, 12) = vpcosto  'pcosto
            .TextMatrix(j, 11) = EsNulo(txtDetalle(6).Text)
                
              
        
            
            If opTipoDoc(1).Value = True Or Not (vIdTipoIva = "001") Then
            
            'Trim(cboTipoIva.Text) = "Consumidor Final" Or Trim(cboTipoIva.Text) = "Responsable Monotributo" Or Trim(cboTipoIva.Text) = "Exento" Then
               ' .TextMatrix(j, 11) = Format(Val(txtDetalle(0).Text), "######0.000") * EsNulo(.TextMatrix(j, 7))
                
                  If (UCase(LeerXml("Puesto")) = "PONS" And UCase(LeerXml("Cliente")) = "PONS") Then
                        .TextMatrix(j, 7) = EsNulo(Val(txtDetalle(2).Text) * (1 + Val(txtDetalle(4).Text) / 100))
                        .TextMatrix(j, 11) = Format(Val(txtDetalle(6).Text * (1 + Val(txtDetalle(4).Text) / 100)), "######0.000")
                    End If
                
                
                
            Else
                If bienes.Value = 1 Then
                    .TextMatrix(j, 7) = EsNulo(txtDetalle(2).Text)  'P. Venta
                    '.TextMatrix(j, 5) = Format(Val(txtDetalle(2).Text) - (Val(txtDetalle(2).Text) * 9.5 / 100), "######0.000") ' (precio)
                    .TextMatrix(j, 11) = "10.5" '(tiva)
                Else
                    .TextMatrix(j, 7) = EsNulo(txtDetalle(2).Text)  'P. Venta
                    .TextMatrix(j, 11) = Val(txtDetalle(6).Text) 'Val(txtDetalle(0).Text) * Val(txtDetalle(2).Text)
                    '.TextMatrix(j, 11) = TraerDato("Articulos", "Codigo = '" & TxtDetalle(1).Tag & "'", "idPorcentajeIva")
                    '.TextMatrix(j, 5) = Format(Val(f(2).Text), "######0.000") 'Format(Val(f(2).Text) - (Val(f(2).Text) * 17.3553 / 100), "######0.000") (precio)
                    '.TextMatrix(j, 11) = "21" '(tiva)
                End If
            End If


            If UCase(LeerXml("PrecioConIva")) = UCase("True") Then

                                    If Me.opTipoDoc(0).Value And Me.cboTipoIva = "Iva Responsable Inscripto" Then
                                    
                                    .TextMatrix(j, 7) = EsNulo(Val(txtDetalle(2).Text) / (1 + Val(txtDetalle(4).Text) / 100))
                                    .TextMatrix(j, 11) = Format(Val(txtDetalle(6).Text / (1 + Val(txtDetalle(4).Text) / 100)), "######0.000")
                        
                                     End If
            End If

            
            
            If Not (vIdTipoIva = "001") And LeerXml("PrecioSinIva") = "True" Then
                        .TextMatrix(j, 7) = EsNulo(Val(txtDetalle(2).Text) * (1 + Val(txtDetalle(4).Text) / 100))
                        .TextMatrix(j, 11) = Format(Val(txtDetalle(6).Text * (1 + Val(txtDetalle(4).Text) / 100)), "######0.000")
            End If
            
            
            
            If (vIdTipoIva = "001") And LeerXml("PrecioSinIva") = "todoch" And txtDetalle(1).Tag = "" Then
                        .TextMatrix(j, 7) = EsNulo(Val(txtDetalle(2).Text) / (1 + Val(txtDetalle(4).Text) / 100))
                        .TextMatrix(j, 11) = Format(Val(txtDetalle(6).Text / (1 + Val(txtDetalle(4).Text) / 100)), "######0.000")
            End If
            
            
            ' cuando txtDetalle(1).Tag = "" es porque agrega un articulo a mano y lo pone con iva. Entonces se lo saco
            If (vIdTipoIva = "001") And LeerXml("PrecioSinIva") = "tito_iva" And txtDetalle(1).Tag = "" Then
                        .TextMatrix(j, 7) = EsNulo(Val(txtDetalle(2).Text) / (1 + Val(txtDetalle(4).Text) / 100))
                        .TextMatrix(j, 11) = Format(Val(txtDetalle(6).Text / (1 + Val(txtDetalle(4).Text) / 100)), "######0.000")
            End If
            
            
            If (vIdTipoIva = "005") And LeerXml("PrecioSinIva") = "tito_iva" And Not txtDetalle(1).Tag = "" Then
                        .TextMatrix(j, 7) = EsNulo(Val(txtDetalle(2).Text) * (1 + Val(txtDetalle(4).Text) / 100))
                        .TextMatrix(j, 11) = Format(Val(txtDetalle(6).Text * (1 + Val(txtDetalle(4).Text) / 100)), "######0.000")
            End If
            
            If (vIdTipoIva = "005") And LeerXml("PrecioSinIva") = "tito_iva" And txtDetalle(1).Tag = "" Then
                        .TextMatrix(j, 7) = EsNulo(Val(txtDetalle(2).Text))
                        .TextMatrix(j, 11) = Format(Val(txtDetalle(6).Text), "######0.000")
            End If
            
            
            If (vIdTipoIva = "004") And LeerXml("PrecioSinIva") = "tito_iva" And Not txtDetalle(1).Tag = "" Then
                        .TextMatrix(j, 7) = EsNulo(Val(txtDetalle(2).Text) * (1 + Val(txtDetalle(4).Text) / 100))
                        .TextMatrix(j, 11) = Format(Val(txtDetalle(6).Text * (1 + Val(txtDetalle(4).Text) / 100)), "######0.000")
            End If
            
            
            
            
            
            
            txtDetalle(1).Tag = ""
            
            '.TextMatrix(j, 8) = Format(.TextMatrix(j, 5) * Val(.TextMatrix(j, 2))) ' (total)

            'If cboVenta = "Contado" Then
                '.TextMatrix(j, 9) = Format(.TextMatrix(j, 5) * Val(.TextMatrix(j, 2)), "######0.000") ' (total_cdo)
            'Else
                '.TextMatrix(j, 10) = Val(.TextMatrix(j, 5)) * Val(.TextMatrix(j, 2)) '(total_ctacte)
            'End If
        

            '.TextMatrix(j, 11) = f(1).Text ' (tiva)
            '.TextMatrix(j, 12) = f(1).Text ' (id)

            '.TextMatrix(j, 13) = venvase ' (envase)
            '.TextMatrix(j, 15) = .TextMatrix(j, 5) * Val(.TextMatrix(j, 2)) ' pago
            '.TextMatrix(j, 16) = TxtDetalle(1).Text ' resta

            'If cboVenta.Text = "Contado" Then
                '.TextMatrix(j, 14) = "SI" ' pagado
                '.TextMatrix(j, 15) = Val(.TextMatrix(j, 5)) * Val(.TextMatrix(j, 2)) 'pago
                '.TextMatrix(j, 16) = 0 '(resta)
            'Else
                '.TextMatrix(j, 14) = "NO"
                '.TextMatrix(j, 15) = 0
                '.TextMatrix(j, 16) = Val(.TextMatrix(j, 5)) * Val(.TextMatrix(j, 2)) * (1 + ((Val(.TextMatrix(j, 11)) / 100))) '(resta)
            'End If

            '.TextMatrix(j, 17) = .TextMatrix(j, 16) ' (totaliva)
            '.TextMatrix(j, 18) = vganancia ' ganancia
            '.TextMatrix(j, 19) = (vganancia * Val(.TextMatrix(j, 5)) * Val(.TextMatrix(j, 3))) / 100 ' (sueldo)
            '.TextMatrix(j, 20) = Trim(txtEmpleados(0).Text) ' (repartidor)
            '.TextMatrix(j, 21) = "N"
            '.TextMatrix(j, 21) = f(1).Text ' (id_ctacte)
            '.TextMatrix(j, 22) = vpespecial ' (pespecial)
            '.TextMatrix(j, 23) = f(1).Text ' (remito_ant)
            '.TextMatrix(j, 24) = f(1).Text ' (id_original)
            '.TextMatrix(j, 25) = f(1).Text ' (modifica)
        Else
            .TextMatrix(j, 1) = txtDetalle(1).Text       '(detalle)
        End If

        Limpiar
        
        vRemitoControl = Val(txtNroRemito.Text)
        vCantidadControl = vCantidadControl + 1
        KlexDetalle.RowSel = KlexDetalle.Rows
       
       Call LastKlexRow(KlexDetalle)
    
        chkfijo.Value = Val(traerDatos2("select * from configuracion", "FijoDocVenta", PathDBConfig))
        
        CalcularTotales
    
    End With
    
    If Err Then GrabarLog "GrabarRenglon", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Private Sub txtPDescuento_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        txtSubtotal.SetFocus
        txtIva(0).SetFocus
        txtIva(1).SetFocus
        txtIva(2).SetFocus
        txtDescuento.SetFocus
        txtImpuesto.SetFocus
    End If

If Err Then GrabarLog "txtPDescuento_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtSubtotal_change()
    On Error Resume Next

    If opTipoDoc(0).Value = True Then
        txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtImpuesto.Text)
    Else
        txtTotal.Text = Val(txtSubtotal.Text)
    End If

    If Err Then GrabarLog "txtSubtotal_change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtSubtotal_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        txtIva(0).SetFocus
        'chkTotalManual.Value = 1
    End If

    If Err Then GrabarLog "subtotal_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtDescuento_Change()
    On Error Resume Next
    
    txtDescuento.Text = Format(txtDescuento.Text, "#######0.000")

    If Err Then GrabarLog "tdescuento_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub opTipoDoc_Click(Index As Integer)
    On Error Resume Next
    
    If Index = 0 Then KlexDetalle.BackColorFixed = &H8080FF
    If Index = 1 Then KlexDetalle.BackColorFixed = &HC0C000
    If Index = 2 Then KlexDetalle.BackColorFixed = &HC0C000
    If Index = 3 Then KlexDetalle.BackColorFixed = &HFF00&
    If Index = 4 Then KlexDetalle.BackColorFixed = &HFFFF&
    
    
    
    CalcularTotales

    GBOtrosDocumentos.Visible = Not True
    fraDetalle.Visible = True
    fraTotales.Visible = True
    PbAcciones(0).Visible = True
    PbAcciones(1).Visible = True
    PbAcciones(2).Visible = True
    
    'lblComentario.Visible = True
    txtObservaciones.Visible = True
    
    'FraAccionesDoc.Top = 7440
    'FraAccionesDoc.Width = 6315
   
    Me.Top = 0
    Me.Height = 9465
    
    'Habilito el boton IMPRIMIR si es una Nota de Credito
   ' BarraDocumento.Buttons(3).Enabled = Not opTipoDoc(2).Value
   ' BarraDocumento.Buttons(4).Enabled = opTipoDoc(2).Value
    
    If cboTipoIva.Tag = "001" Then
        cboLetra.Text = "A"
       ' cboPuntoDeVenta.Text = "0002"
    '    cboLetra_Click
       ' cboPuntoDeVenta.Text = "0002"
        'cboPuntoDeVenta_Click
    Else
        cboLetra.Text = "B"
       ' cboPuntoDeVenta.Text = "0002"
   '     cboLetra_Click
      '  cboPuntoDeVenta.Text = "0002"
        'cboPuntoDeVenta_Click
    End If
    
    
    If LeerXml("Puesto") = "CajaComuna" Then
       cboLetra.Text = "X"
    End If
    
    cboLetra_Click
    
    
   If Trim(Me.cboPuntoDeVenta) = "" Then Me.cboPuntoDeVenta.Text = traerDatos2("select * from configuracion", "SucursalDocVenta", PathDBConfig)
    
    If Err Then GrabarLog "opTipoDoc_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub LimpiarTotales()
    On Error Resume Next

    If chkTotalManual.Value = 0 Then
        txtSubtotal.Text = ""
        
       ' txtPDescuento.Text = ""
        txtImpuesto.Text = ""
        'txtDescuento.Text = ""
        txtTotal.Text = ""
    End If
    '  Me.vcae.Text = ""
    'Me.vcaeFecha.Text = ""
   ' txtNroInterno = UltimoNroInterno2 + 1
    
    If Err Then GrabarLog "LimpiarTotales", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub BarraCliente_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next

    Select Case Button.Index

        Case 1
            If ValidarClienteNuevo() = True Then
                GuardarCliente
                Call txtCliente_KeyPress(0, 13)
            End If

        Case 2
            NuevoCliente
            
        Case 3
            BuscaDoc
    
    End Select

    If Err Then GrabarLog "Toolbar1_ButtonClick (" & Button & ")", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function ValidarClienteNuevo() As Boolean
On Error Resume Next

    ValidarClienteNuevo = False
    
    If Trim(cboTipoIva.Tag) = "" Then
        MsgBox "Debe Ingresar un Tipo de Iva para el Cliente", vbExclamation, "Mensaje ..."
        Exit Function
    End If
    
    If (Trim(cboTipoIva.Tag) = "001" Or Trim(cboTipoIva.Tag) = "003") And (Trim(txtCliente(3).Text) = "") Then
        MsgBox "Debe ingresar el CUIT del Cliente", vbExclamation, "Mensaje ..."
        Exit Function
    End If
        
    If (Trim(cboTipoIva.Tag) = "001" Or Trim(cboTipoIva.Tag) = "003") And (ValidarCuit(txtCliente(3).Text) = False) Then
        MsgBox "Debe ingresar el CUIT valido del Cliente", vbExclamation, "Mensaje ..."
        Exit Function
    End If
    
    ValidarClienteNuevo = True

If Err Then GrabarLog "ValidarClienteNuevo", Err.Number & " " & Err.Description, Me.Caption
End Function

Private Sub txtTipoMovimiento_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        Select Case Index
        
            Case 0
                txtTipoMovimiento(0).Text = UCase(txtTipoMovimiento(0).Text)
                txtTipoMovimiento(1).Text = TraerDato("TipoMovimientos", "Codigo = '" & Trim(txtTipoMovimiento(0).Text) & "'", "TipoMovimiento")
                
            Case 1
                txtTipoMovimiento(0).Text = UCase(TraerDato("TipoMovimientos", "TipoMovimiento = '" & Trim(txtTipoMovimiento(1).Text) & "'", "Codigo"))
        
        End Select
    
    
        txtIB(0).SetFocus
    

    End If
    
    

If Err Then GrabarLog "txtTipoMovimiento_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtTipoMovimiento_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        pbCarga_Click (1)
    End If

If Err Then GrabarLog "txtTipoMovimiento_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtTipoMovimiento_LostFocus(Index As Integer)
Me.txtIB(7) = "[Doc: " + Me.txtTipoMovimiento(0) + " " + cboLetra + " " + txtNroComprobante + "]"
End Sub

Private Sub txtTotal_Change()
    On Error Resume Next
    
    txtTotal.Text = Format(txtTotal.Text, "#######0.000")

    If Err Then GrabarLog "txtTotal_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtTotal_GotFocus()
    On Error Resume Next
    
    txtTotal.Text = Val(txtSubtotal.Text) + Val(txtIva(0).Text) + Val(txtIva(1).Text) + Val(txtIva(2).Text) - Val(txtDescuento.Text) + Val(txtImpuesto.Text)

    If Err Then GrabarLog "txtTotal_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub tprecio_Click()
    On Error Resume Next
    
    focoEnLinea
    
    'txtDetalle(0).SetFocus

    If Err Then GrabarLog "tprecio_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub focoEnLinea()
If UCase(LeerXml("Puesto")) = "PONS" Then
'If Trim(vConfigGral.vEmpresa) = Trim("wgestionpons") Then
 txtDetalle(1).SetFocus
Else
 txtDetalle(0).SetFocus
End If

End Sub
Private Sub UltimaVenta()
    On Error Resume Next

    If Not Val(txtTotal.Text) = 0 Or Not Val(txtIB(10).Text) = 0 Then
        EjecutarScript ("UPDATE Clientes SET U_Venta = '" & strfechaMySQL(dtpFecha.Value) & "' WHERE Codigo = '" & txtCliente(0).Tag & "'")
    Else
        Exit Sub
    End If
    
    If Err Then GrabarLog "UltimaVenta", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtCliente_Change(Index As Integer)
On Error Resume Next

    'If Not vActualizaNombre = True Then
    'If Index = 0 And Not vGrabaModo = 1 Then
    If Index = 0 Then
        If Not vOpenGrilla(0) = True Then
            Call MostrarCoincidencias("Clientes", Trim(txtCliente(0).Text))
        End If
    End If
    'End If
        
If Err Then GrabarLog "txtCliente_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtCliente_KeyPress(Index As Integer, _
                      KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
    
        If Index = 3 Then
            focoEnLinea
            'txtDetalle(0).SetFocus
            'MsgBox "" 'NroComprobante
            Exit Sub
        End If
       
        If Index = 0 Then
        
            txtCliente(0).Text = EsNulo(rsClientes.Fields("Codigo").Value)
            txtcodigoCliente.Text = txtCliente(0).Text
         
            If BuscarCliente = True Then
                dtpFecha.SetFocus
                
                Call fijarTipoDocumento
            Else
                    frmClientesAlta.viente = Me.Name
                    frmClientesAlta.txtAlta(1) = Me.txtCliente(0).Text
                    frmClientesAlta.Show
            End If
            
        Else
            If Index >= 3 Then
                focoEnLinea
                'txtDetalle(0).SetFocus
            Else
                txtCliente(Index + 1).SetFocus
            End If
        End If
    
    End If

    If Err Then GrabarLog "txtCliente_KeyPress (" & Index & "-" & KeyAscii & ")", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub fijarTipoDocumento()
On Error Resume Next
' fija el tipo de documento
Dim vsql As String
Dim vIdTipoIva, vTipoIva As String

    vIdTipoIva = TraerDato("Clientes", "Codigo = '" & Me.txtCliente(0).Tag & "'", "idTipoIva")
    vTipoIva = TraerDato("TipoIva", "idTipoIva = '" & vIdTipoIva & "'", "TipoIva")
    
   Call Me.opTipoDoc_Click(3)
    
    Select Case UCase(LeerXml("Remito"))
    
        Case "FACTURA"
           Call Me.opTipoDoc_Click(0)
           Me.opTipoDoc(0).Value = True
           
        Case "DOCUMENTO"
           Call Me.opTipoDoc_Click(3)
           Me.opTipoDoc(3).Value = True
          
    End Select
    
If Err Then GrabarLog "VerDocumento", Err.Number & " " & Err.Description, Me.Caption




End Sub



Private Sub txtEmpleados_KeyPress(Index As Integer, KeyAscii As Integer)
    On Error Resume Next

    Exit Sub
    If KeyAscii = 13 And vOpenGrilla(2) = False Then
        Dim rsempleados As New ADODB.Recordset, sqlEmpleados As String
        
        sqlEmpleados = "SELECT * FROM empleados WHERE (Codigo LIKE '%" + Trim(txtEmpleados(0).Text) + "%') OR (Nombre like '%" + Trim(txtEmpleados(0).Text) + "%')"
        
        With rsempleados
            .CursorLocation = adUseClient
            
            Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockReadOnly)
    
            If Not .EOF = True Then
                txtEmpleados(0).Text = .Fields("Codigo").Value
                txtEmpleados(1).Text = .Fields("Nombre").Value
        
                txtEmpleados(0).BackColor = vbWhite
                
                focoEnLinea
                'txtDetalle(0).SetFocus 'Se va a cantidad

            Else
                MsgBox "El Empleado no fue encontrado.", vbInformation, "Mensaje ..."
            End If
        
        End With

        sqlEmpleados = ""
        
        If rsempleados.State = 1 Then
            rsempleados.Close
            Set rsempleados = Nothing
        End If
    End If

    If Err Then GrabarLog "txtEmpleado_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Function TipoDocumento() As String
    On Error Resume Next
    
    Dim i As Integer

    If Not opOtrosDocumentos.Value = True Then
        For i = 0 To 5

            If opTipoDoc(i).Value = True Then

                Select Case i
    
                    Case 0
                        If Trim(cboTipoIva.Text) = "Iva Responsable Inscripto" Then
                            TipoDocumento = "Fact A"
                            Exit For
                        Else
                            'Este es el original - LO ATO CON ALAMBRE!!!!
                            TipoDocumento = "Fact B"
                            Exit For
                        End If

                    Case 1
                        TipoDocumento = "Presupuesto"
                        Exit For

                    Case 2
                        TipoDocumento = "Nota C"
                        Exit For

                    Case 3
                        TipoDocumento = "Documento"
                        Exit For

                    Case 4
                        TipoDocumento = "Remito"
                        Exit For
                
                    Case 5
                        TipoDocumento = "Nota D"
                        Exit For
            
                End Select

            End If

        Next
    Else
        Select Case UCase(txtTipoMovimiento(0).Text)
            
            Case "RI", "RV", "CD", "CC", "FC", "RG", "SU", "SI", "AD", "AC", "IB"
                TipoDocumento = "Fact " & Trim(cboLetra.Text)
                
            Case "NC"
                TipoDocumento = "Nota C"
                
            Case "ND"
                TipoDocumento = "Nota D"
                
            Case "RC"
                TipoDocumento = "Recibo"
                
        End Select
    
    End If
    
    If Err Then GrabarLog "TipoDocumento", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub ControlRemito()
    On Error Resume Next
    
    Dim rsControl As New ADODB.Recordset, sqlControl As String
    
    sqlControl = "SELECT * FROM Factura INNER JOIN FDetalle ON Factura.remito = FDetalle.remito WHERE (Factura.remito = " & vRemitoControl & ")"
    
    With rsControl
        Call .Open(sqlControl, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .RecordCount <> vCantidadControl Then
    
            'MsgBox "Error cuando se graba el documento de " & v(0).Text & ""
            
        Else
        
            'Todo bien
        
        End If
    
    End With
    
    sqlControl = ""
    
    rsControl.Close
    Set rsControl = Nothing

    If Err Then GrabarLog "ControlRemito", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarIva()
    On Error Resume Next
    
    Dim vTipoDocumentoIva As String
    
    
    If Not opOtrosDocumentos.Value = True Then
    
        vTipoDocumentoIva = TipoDocumento
    
        If vTipoDocumentoIva = "Fact A" Or vTipoDocumento = "Fact B" Or vTipoDocumentoIva = "Nota C" Or vTipoDocumentoIva = "Nota D" Then
        
        
            If vGrabaModo = 0 Then
                'If vTipoDocumentoIva = "Nota C" Then
                  Call EjecutarScript("INSERT INTO IvaFacturaVenta (nrointerno, remito, Iva105, Iva210, Iva270) VALUES (" & Str(Val(txtNroInterno)) & "," & vremito & ", " & Val(txtIva(0).Text) & ", " & Val(txtIva(1).Text) & ", " & Val(txtIva(2).Text) & ");")
                'Else
                '  Call EjecutarScript("INSERT INTO IvaFacturaVenta (nrointerno, remito, Iva105, Iva210, Iva270) VALUES (" & Str(Val(txtNroInterno)) & "," & vremito & ", " & Val(txtIva(0).Text) & ", " & Val(txtIva(1).Text) & ", " & Val(txtIva(2).Text) & ");")
                'End If
            Else
                'Panic 'Controlar si es Doc A Fact
                'Call EjecutarScript("INSERT INTO IvaFacturaVenta (remito, Iva105, Iva210, Iva270) VALUES (" & vRemito & ", " & Val(txtIva(0).Text) & ", " & Val(txtIva(1).Text) & ", " & Val(txtIva(2).Text) & ");")
                          
                'If vTipoDocumentoIva = "Nota C" Then
                    Call EjecutarScript("UPDATE IvaFacturaVenta SET nrointerno=" + Str(Val(txtNroInterno)) + ",Iva105 = " & Val(txtIva(0).Text) & ", Iva210 = " & Val(txtIva(1).Text) & ", Iva270 = " & Val(txtIva(2).Text) & " WHERE (remito = " & vremito & ")")
                'Else
                '    Call EjecutarScript("UPDATE IvaFacturaVenta SET nrointerno=" + Str(Val(txtNroInterno)) + ",Iva105 = " & Val(txtIva(0).Text) & ", Iva210 = " & Val(txtIva(1).Text) & ", Iva270 = " & Val(txtIva(2).Text) & " WHERE (remito = " & vremito & ")")
                'End If

            End If
    
        End If
    
    Else
    
        If vGrabaModo = 0 Then
            Select Case Val(txtIB(1).Text)
            
             Case 10.5
                 Call EjecutarScript("INSERT INTO IvaFacturaVenta (nrointerno, remito, Iva105,Retenciones, Percepciones, NoGravado, ITC, ImpExento) VALUES (" & txtNroInterno & "," & vremito & ", " & Val(txtIB(2).Text) & "," & Str(Val(txtIB(4).Text)) & "," & Str(Val(txtIB(3).Text)) & "," & Str(Val(txtIB(5).Text)) & ",0," & Str(Val(txtIB(6).Text)) & ");")
             Case 21
                 Call EjecutarScript("INSERT INTO IvaFacturaVenta (nrointerno, remito, Iva210,Retenciones, Percepciones, NoGravado, ITC, ImpExento) VALUES (" & txtNroInterno & "," & vremito & ", " & Val(txtIB(2).Text) & "," & Val(txtIB(4).Text) & "," & Val(txtIB(3).Text) & "," & Val(txtIB(5).Text) & ",0," & Val(txtIB(6).Text) & ")")
             Case 27
                 Call EjecutarScript("INSERT INTO IvaFacturaVenta (nrointerno, remito, Iva270,Retenciones, Percepciones, NoGravado, ITC, ImpExento) VALUES (" & txtNroInterno & "," & vremito & ", " & Val(txtIB(2).Text) & "," & Val(txtIB(4).Text) & "," & Val(txtIB(3).Text) & "," & Val(txtIB(5).Text) & ",0," & Val(txtIB(6).Text) & ");")
            
             Case Else
                 Call EjecutarScript("INSERT INTO IvaFacturaVenta (nrointerno,remito, Retenciones, Percepciones, NoGravado, ITC, ImpExento) VALUES (" & txtNroInterno & "," & vremito & "," & Val(txtIB(4).Text) & "," & Val(txtIB(3).Text) & "," & Val(txtIB(5).Text) & ",0," & Val(txtIB(6).Text) & ");")
            
            End Select
        Else
            
        End If
    End If
    
    If Err < 0 Then
        'MsgBox "Cuidado!. Verificar si los datos fueron actualizados en IVA VENTA"
        'GrabarLog "GuardarIva", Err.Number & " " & Err.Description, Me.Name
    End If
End Sub
Public Sub DecorarTalles()
Dim i, j, k As Integer

On Error Resume Next

    With KlexDetalle
    
        k = 0
        For i = 0 To .Rows - 1

            .TextMatrix(.Row, 8) = Format(.TextMatrix(.Row, 8), "#####0.000")
            .TextMatrix(.Row, 5) = Format(.TextMatrix(.Row, 5), "#####0.000")

            .Row = i
    
            For j = 1 To 9
                .Col = j
                If k = 1 Then
                    .CellBackColor = &HFFC0C0
                End If
            Next
            If k = 0 Then
                k = 1
            Else
                k = 0
            End If
        Next
    
    End With

If Err Then GrabarLog "DecorarTalles", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub WCaja(vConceptoCaja, vFechaCaja As Date, vImporteCaja As Double, vCodCliente As String)
    On Error Resume Next
    
    Dim rsCaja As New ADODB.Recordset
    Dim sqlCaja As String
    
    With rsCaja
        'If Not grabamodo = 1 Then
        
            sqlCaja = "SELECT * FROM caja"
            .Open sqlCaja, ConnDDBB, adOpenDynamic, adLockPessimistic
        
            .AddNew
            .Fields("remito").Value = vremito
        
        'Else
            
        '    sqlCaja = "SELECT * FROM caja WHERE (remito = " & Trim(vremito) & ")"
        '    .Open sqlCaja, ConnDDBB, adOpenDynamic, adLockPessimistic
        '
        '    If .EOF = True Then
        '        .AddNew
        '        .Fields("Remito").Value = vremito
        '    End If
            
        'End If
        
        .Fields("fecha").Value = strfechaMySQL(vFechaCaja)
        .Fields("Importe").Value = Val(vImporteCaja)
        
        .Fields("CodigoCliente").Value = vCodCliente
        
        .Fields("Usuario").Value = vConfigGral.vUser
        .Fields("CodigoConcepto").Value = vConceptoCaja
        '.Fields("comentario").Value = ""
            
        .Fields("NroCheque") = Null
        .Fields("FechaDeposito") = Null
        .Fields("FechaConfeccion") = Null
        .Fields("idCajas") = Null
        
        .Update
    
    End With
    
    sqlCaja = ""
    
    rsCaja.Close
    Set rsCaja = Nothing
   
If Err Then GrabarLog "WCAJA", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Public Sub WCtaCte(vnroremito As Long)
On Error Resume Next


    If Me.vcventa.Text = "Contado" Then
        frmCobros.cpInstancia = "cobro"
        frmCobros.HabilitarControles (True)
        frmCobros.txtCliente(0).Text = Me.txtCliente(0).Tag
        frmCobros.txtCliente(0).Tag = Me.txtCliente(0).Tag
        frmCobros.txtCliente(1).Text = Me.txtCliente(0).Text
        
        frmCobros.vfechaCredito = strfechaMySQL(dtpFecha.Value)
        
        frmCobros.txtImporteEfectivoPesos = Val(Format(Me.txtTotal.Text, "#####0.000"))
      '  Exit Sub
        
    End If


' -------------- vuelvo a gral el nro de asiento para grabar en el movi de wctacte --------
    vnroasiento = Val(GenerarDato("SELECT MAX(Numero) AS UAsiento FROM Asientos WHERE NroBalance = " + Str(vnrobalance), "UAsiento")) + 1
' ------------------------------------------------------
 

    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC, vsql2 As String
    Dim vnrointerno2 As Long
    
    vsql2 = "select nrointerno as c from factura where remito = " + Str(vnroremito)
    vnrointerno2 = traerDatos2(vsql2, "c", pathDBMySQL)
    
    If Me.vGrabaModo = 1 Then
        sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (nrointerno = " & Str(vnrointerno2) & ")"
    Else
        sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE 1=2"
    End If
    With rsCtaCteC
        .CursorLocation = adUseClient
        
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
       
        If Not .EOF = True Then
            
        
        Else
            
            If Me.vGrabaModo = 1 Then
                MsgBox "Revise el estado de la CtaCte "
            End If
         
            .AddNew
            .Fields("remito").Value = Trim(vnroremito)
        
        
        End If
        
        
        
        .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Fechainput").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Codigo").Value = txtCliente(0).Tag
        .Fields("Nombre").Value = txtCliente(0).Text
       
        .Fields("anomes").Value = Right(.Fields("Fecha").Value, 4) & Mid(.Fields("Fecha").Value, 4, 2)
    
        If (TipoDocumento = "Nota C") Then
            
            .Fields("credito").Value = Val(Format(Me.txtTotal.Text, "#####0.000"))
            .Fields("debito").Value = 0
            .Fields("saldo") = 0 'bcliente.Recordset("saldo") - bfactura_temp.Recordset("Total")
            .Fields("idMedioPago").Value = 8
            .Fields("comentario").Value = Left("Nro. " & TipoDocumento & " " & Trim(cboPuntoDeVenta.Text) & "-" & FormatoNroFactura(Trim(txtNroComprobante.Text)), 100)
        
            .Fields("TipoMovimiento").Value = UCase(Trim(txtTipoMovimiento(0).Text))
        
        Else
        
            If (TipoDocumento = "Documento") Or (TipoDocumento = "Fact A") Or (TipoDocumento = "Fact B") Or (TipoDocumento = "Nota D") Then
            
                If (opOtrosDocumentos.Value = True) Then
                    
                    If Val(txtIB(10).Text) < 0 Then
                        
                        .Fields("Credito").Value = Val(Format(txtIB(10).Text, "#####0.000")) * (-1)
                        .Fields("Debito").Value = 0
                        .Fields("TipoMovimiento").Value = UCase(Trim(txtTipoMovimiento(0).Text))
                    
                    Else
                        
                        If txtTipoMovimiento(0).Text = "AC" Then ' veo si es un ajuste de credito, en los otros casos toma debito
                            .Fields("Credito").Value = Val(Format(txtIB(10).Text, "#####0.000"))
                            .Fields("Debito").Value = 0
                        Else
                            .Fields("Debito").Value = Val(Format(txtIB(10).Text, "#####0.000"))
                            .Fields("Credito").Value = 0
                        End If
                        
                        
                            .Fields("TipoMovimiento").Value = UCase(Trim(txtTipoMovimiento(0).Text))
                    
                    End If
                    
                    .Fields("comentario").Value = Left(TraerDato("Factura", "Remito = " & Val(vnroremito) & "", "Comentario"), 100)
                    .Fields("saldo") = 0
                Else
                    .Fields("debito") = Val(Format(txtTotal.Text, "#####0.000"))
                    .Fields("credito") = 0
                    .Fields("saldo") = Val(Format(txtTotal.Text, "#####0.000"))
                    .Fields("comentario").Value = Left("Nro. " & TipoDocumento & " " & Trim(cboPuntoDeVenta.Text) & "-" & FormatoNroFactura(Trim(txtNroComprobante.Text)), 100)
                End If
            End If
        
        End If
        
        .Fields("NroInterno").Value = Val(txtNroInterno.Text)
        .Fields("NroAsiento").Value = Val(vnroasiento)
        .Update

        '--------------------------------------
        'vgidCtaCte = .Fields("id") ' guardo el id de la ctacte para poner el nro de asiento
        'vgTablaCtacCte = "cuentascorrientes"
        '-----------------------------------------

        Select Case txtTipoMovimiento(0).Text
        
            Case "CD"
                Call EjecutarScript("INSERT INTO CuentasCorrientes (Fecha,Codigo,Nombre,Credito,Comentario,Remito,NroInterno,TipoMovimiento,idMedioPago, NroAsiento) VALUES ('" & strfechaMySQL(Me.dtpFecha.Value) & "','" & txtCliente(0).Tag & "','" & txtCliente(0).Text & "'," & Val(txtIB(10).Text) & ",'" & Trim(.Fields("Comentario").Value) & "'," & vnroremito & "," & Val(txtNroInterno.Text) & ",'CC', 11," + Str(vnroasiento) + ")")
                Call EjecutarScript("INSERT INTO BancosMovimientos (idBancos,idBancosCuentas,Fecha,Credito,Comentario,NroCheque,TipoMovimiento,NroInterno,idTipoValor, NroAsiento) VALUES ('" & Trim(txtCaja(0).Text) & "'," & Val(txtCaja(2).Text) & ",'" & strfechaMySQL(dtpFecha.Value) & "'," & Val(txtCaja(7).Text) & ",'" & Trim(txtCaja(8).Text) & "'," & Val(txtCaja(6).Text) & ",'CC'," & Val(txtNroInterno.Text) & ",'" & Trim(txtCaja(4).Text) & "'," + Str(vnroasiento) + ")")
        
            Case "AD"
                ' si es un ajuste débito quiere decir que solo se ingresa un movi en el debito y no en la credito
                'Call EjecutarScript("INSERT INTO CuentasCorrientes (Fecha,Codigo,Nombre,Credito,Comentario,Remito,NroInterno,TipoMovimiento, NroAsiento) VALUES ('" & strfechaMySQL(dtpFecha.Value) & "','" & txtCliente(0).Tag & "','" & txtCliente(0).Text & "'," & Val(txtIB(10).Text) & ",'" & Trim(.Fields("Comentario").Value) & "'," & vNroRemito & "," & Val(txtNroInterno.Text) & ",'AC'," + Str(vnroasiento) + ")")
            
            Case "SU", "SI", "RG", "IB", "RV"
           '    Call EjecutarScript("INSERT INTO CuentasCorrientes (Fecha,Codigo,Nombre,Credito,Comentario,Remito,NroInterno,TipoMovimiento) VALUES ('" & strfechaMySQL(dtpFecha.value) & "','" & txtCliente(0).Tag & "','" & txtCliente(0).Text & "'," & Val(.Fields("Debito").value) & ",'" & Trim(.Fields("Comentario").value) & "'," & vNroRemito & "," & Val(txtNroInterno.Text) & ",'CC')")
            
            Case "FC"
                'No pasa NADA
            Case Else
                   'MsgBox "OJO"
        
        End Select
    
            
    End With
                    
If Err < 0 Then
    'GrabarLog "WCtaCte", Err.Number & " " & Err.Description, Me.Name
    If Err < 0 Then
        'MsgBox "Cuidado!. La información pudo no haberse guardado en la Cuenta Corriente"
    End If
End If
End Sub
'Metodo obsoleto
Private Function TipoVenta(vTipoVenta As String, vnroremito As Long) As Boolean
On Error Resume Next

    Select Case Trim(vTipoVenta)
        
        Case "Cuenta Corriente"
            WCtaCte (vnroremito)
        
        Case "Cheques"
            WCtaCte (vnroremito)
            'wcheques
        
        Case "Contado"
            WCtaCte (vnroremito)
        
        Case "Credito"
            'wcredito
    
    End Select
    
If Err Then GrabarLog "TipoVenta", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub txtCliente_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If Index = 0 Then
        If KeyCode = 38 Then
            With rsClientes
                If Not .EOF = True And Not .BOF = True Then
                    .MovePrevious
                    'Me.dgClientes.Bookmark = Me.dgClientes.Row

                Else
                    .MoveLast
                End If
            End With
        End If

        If KeyCode = 40 Then
            With rsClientes
                If Not .EOF = True And Not .BOF = True Then
                    .MoveNext
                    
                   'Me.dgClientes.Bookmark = Me.dgClientes.Row
                   
                    
                Else
                    .MoveFirst
                End If
            End With
        End If
    End If
    
    If KeyCode = 13 Then
        'If Not rsClientes.EOF = True Then
            dgClientes_DblClick
            
            ' init con clinentes
            
        'End If
    End If
    
If Err Then GrabarLog "txtCliente_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtEmpleados_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

    Exit Sub
    If KeyCode = 38 Then
        With rsEmpleadosGrilla
            If Not .EOF = True And Not .BOF = True Then
                .MovePrevious
            Else
                .MoveLast
            End If
        End With
    End If

    If KeyCode = 40 Then
        With rsEmpleadosGrilla
            If Not .EOF = True And Not .BOF = True Then
                .MoveNext
            Else
                .MoveFirst
            End If
        End With
    End If

    
    If KeyCode = 13 Then
        dgEmpleados_DblClick
    End If
    
If Err Then GrabarLog "txtEmpleado_KeyUp", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrillaDetalle(vCantidadRenglones As Integer)
On Error Resume Next

    Err.Clear

    Dim i As Integer

    With KlexDetalle
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 26
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 125
        
        'Aca Pego el IdFDetalle-Entonces se si modifico o NO
        .TextMatrix(0, 1) = "idFDetalle"
        .ColWidth(1) = 0
        
        .TextMatrix(0, 2) = "Fecha"
        .ColWidth(3) = 0
        
        .TextMatrix(0, 3) = "Remito"
        .ColWidth(2) = 0
        
        .TextMatrix(0, 4) = "Codigo"
        .ColWidth(4) = 0
        
        .TextMatrix(0, 5) = "Cant."
        .ColWidth(5) = 750
        .ColDisplayFormat(5) = "#0.000"
        
        .TextMatrix(0, 6) = "Detalle"
        .ColWidth(6) = 9500
        .ColAlignment(6) = 0
        
        .TextMatrix(0, 7) = "P. Venta"
        .ColWidth(7) = 850
        .ColDisplayFormat(7) = "#0.000"
        
        .TextMatrix(0, 8) = "% Desc."
        .ColWidth(8) = 850
        .ColDisplayFormat(8) = "#0.000"
                
        .TextMatrix(0, 9) = "% Iva"
        .ColWidth(9) = 850
        .ColDisplayFormat(9) = "#0.000"
        
        .TextMatrix(0, 10) = "% Imp."
        .ColWidth(10) = 850
        .ColDisplayFormat(10) = "#0.000"

        .TextMatrix(0, 11) = "$ Total"
        .ColWidth(11) = 900
        .ColDisplayFormat(11) = "#0.000"
        
        .ColWidth(25) = 200
        .TextMatrix(0, 25) = ""
        
        .Col = 25

        If .Rows = 2 Then
            .Row = 1
        Else
            .Row = .Rows - 1
        End If
        
        .CellBackColor = &HFFFCCC
        '.CellBackColor = vbRed
        
        .Editable = True

        '.EnterKeyBehaviour = klexEKMoveDown
        .EnterKeyBehaviour = klexEKNone
        .BackColorAlternate = &HE0E0E0

    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarGrilla()
On Error Resume Next
    
    With KlexDetalle

        .TextMatrix(.Rows - 1, 1) = ""
        .TextMatrix(.Rows - 1, 2) = EsNulo(txtNroRemito.Text)
        .TextMatrix(.Rows - 1, 3) = dtpFecha.Value
        .TextMatrix(.Rows - 1, 4) = vvcodigo
        .TextMatrix(.Rows - 1, 5) = EsNulo(txtDetalle(0).Text)
        .TextMatrix(.Rows - 1, 6) = EsNulo(txtDetalle(1).Text)
        .TextMatrix(.Rows - 1, 7) = EsNulo(txtDetalle(2).Text)
        .TextMatrix(.Rows - 1, 8) = EsNulo(txtDetalle(3).Text)
        .TextMatrix(.Rows - 1, 9) = EsNulo(txtDetalle(4).Text)
        .TextMatrix(.Rows - 1, 10) = EsNulo(txtDetalle(5).Text)
        .TextMatrix(.Rows - 1, 11) = EsNulo(txtDetalle(6).Text)
    
        If Trim(vLeyendaAsiento) = "" Then
            vLeyendaAsiento = txtDetalle(1).Text
        Else
            vLeyendaAsiento = vLeyendaAsiento & " - " & EsNulo(txtDetalle(1).Text)
        End If
    
        CalcularTotales
    End With
    
    Limpiar

If Err Then GrabarLog "CargarGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function VerIvaArticulo(vCodigoArticulo As String) As String
On Error Resume Next

    Dim rsPorcentajeIva As New ADODB.Recordset, sqlPorcentajeIva As String
    
    sqlPorcentajeIva = "SELECT * FROM Articulos A INNER JOIN PorcentajeIva P ON A.IdPorcentajeIva=P.idPorcentajeIva WHERE (A.Codigo = '" & vCodigoArticulo & "')"
    
    With rsPorcentajeIva
        Call .Open(sqlPorcentajeIva, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If (.State = 1) And Not (.EOF = True) Then
            
            VerIvaArticulo = Replace(.Fields("Porcentaje").Value, ".", "")
        
            VerIvaArticulo = "0." & VerIvaArticulo
        Else
            VerIvaArticulo = ""
        End If
    
    End With

    sqlPorcentajeIva = ""
    
    If rsPorcentajeIva.State = 1 Then
        rsPorcentajeIva.Close
        Set rsPorcentajeIva = Nothing
    End If
    
If Err Then GrabarLog "VerIvaArticulo", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Function VerDocumento(vTipoIva As String, vTipoDocumento As String, vCodCliente As String) As String
On Error Resume Next

    Dim vIdTipoIva As String

    vIdTipoIva = TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "idTipoIva")
    vTipoIva = TraerDato("TipoIva", "idTipoIva = '" & vIdTipoIva & "'", "TipoIva")
    
    Select Case vTipoIva
    
        Case "Cons. Final", "Consumidor Final"
            Select Case vTipoDocumento
            
                Case "T"
                    VerDocumento = TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "TipoDocumento")
                Case "N"
                    VerDocumento = TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "NroDocumento")
                
            End Select
            If vTipoDocumento = "T" Then
                
            Else
                
            End If
            
        Case "Exento"
            
            
            If vTipoDocumento = "T" Then
                VerDocumento = TraerDato("TipoIva", "idTipoIva = '" & vIdTipoIva & "'", "AliasAfip")
            Else
                VerDocumento = TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "Cuit")
            End If
            
            
            
        Case "Iva Responsable Inscripto"
            If vTipoDocumento = "T" Then
                VerDocumento = TraerDato("TipoIva", "idTipoIva = '" & vIdTipoIva & "'", "AliasAfip")
            Else
                VerDocumento = Replace(TraerDato("Clientes", "Codigo = '" & vCodCliente & "'", "Cuit"), "-", "")
            End If
    
    End Select
    
If Err Then GrabarLog "VerDocumento", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Function VerTipoIva(vcliente As String) As String
On Error Resume Next

    Dim vIdTipoIva As String

    vIdTipoIva = TraerDato("Clientes", "Codigo = '" & vcliente & "'", "idTipoIva")
    
    VerTipoIva = TraerDato("TipoIva", "idTipoIva = '" & vIdTipoIva & "'", "TipoIva")
    
If Err Then GrabarLog "VerTipoIva", Err.Number & " " & Err.Description, Me.Caption
End Function
Public Function SaldoCliente()
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT Codigo, SUM(Debito) , SUM(Credito), SUM(Debito) - SUM(Credito) AS SaldoCliente FROM CuentasCorrientes WHERE (codigo = '" & Trim(txtCliente(0).Tag) & "') GROUP BY Codigo"
  
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        Dim SaldoAnterior, totaldebito, totalCredito, credito, debito As Double
        
        SaldoAnterior = 0
        
        If Not .EOF = True Then
        
            SaldoAnterior = Format(.Fields("SaldoCliente").Value, "#######0.000")
        Else
            SaldoAnterior = Format(0, "#######0.000")
        End If
        
        'Do While Not .EOF
        '    If IsNull(.Fields("debito")) Or .Fields("debito") = "" Then
        '        debito = 0
        '    Else
        '        debito = .Fields("debito")
        '    End If
                        
        '    If IsNull(.Fields("credito")) Or .Fields("credito") = "" Then
        '        credito = 0
        '    Else
        '        credito = .Fields("credito")
        '    End If
                        
       '     SaldoAnterior = SaldoAnterior + debito - credito
       '     totalCredito = totalCredito + credito
       '     totaldebito = totaldebito + debito
            
       '     .MoveNext
       ' Loop
        
    End With
    
    'If totaldebito = "" Then
    '    totaldebito = 0
    'End If
    
    'If totalCredito = "" Then
    '    totalCredito = 0
    'End If
    
    'If SaldoAnterior = "" Then
    '    SaldoAnterior = 0
    'End If
    
    If rsCtaCteC.State = 1 Then
        rsCtaCteC.Close
        Set rsCtaCteC = Nothing
    End If
    
    sqlCtaCteC = ""
    
    SaldoCliente = SaldoAnterior
    
If Err Then GrabarLog "SaldoCliente", Err.Number & " " & Err.Description, Me.Name
End Function

Function fcomoponerlistachoferes() As String
Dim vstring As String
vstring = ""
Dim i As Integer
For i = 0 To Me.grd_Choferes.Rows - 1
    vstring = vstring + (Me.grd_Choferes.TextMatrix(i, 1)) + ","
Next

fcomoponerlistachoferes = vstring
End Function


Public Sub CargarChoferAGrilla(vlista As String)
On Error Resume Next

Dim valista() As String
Dim i As Integer

valista = Split(vlista, ",")

grd_Choferes.Rows = 0

For i = 0 To UBound(valista)

    With grd_Choferes
        .Cols = 3
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 1) = valista(i)
        .TextMatrix(.Rows - 1, 2) = traerDatos2("select * from empleados where codigo=" + valista(i), "Nombre", pathDBMySQL)
    End With

Next

grd_Choferes.Rows = grd_Choferes.Rows - 1

If Err Then Exit Sub
End Sub

Private Sub vcodEmpresa_Change()
    Call cambio_empresa
    
    If UCase(LeerXml("Puesto")) = "EMPRESAS" Then
        Me.txtNroComprobante.Text = getNroCompAfip
    End If
    
End Sub

Private Sub cambio_empresa()
    Dim vsql As String
    
    vnroempresa = Val(Me.vcodEmpresa.Text)
    
    vsql = "select * from DatosEmpresas where idEmpresas = " + Str(vnroempresa + 1)
    
    
    eeCuit = traerDatos2(vsql, "cuit", PathDBConfig)
    eeDireccion = traerDatos2(vsql, "Direccion", PathDBConfig)
    eeLocalidad = traerDatos2(vsql, "Localidad", PathDBConfig)
    eeIngresosBrutos = traerDatos2(vsql, "EMail", PathDBConfig)
    
    
'-------------
    
    vsql = "select * from Empresas where idEmpresas= " + Str(vnroempresa + 1)
    eeEmpresa = traerDatos2(vsql, "Empresa", PathDBConfig)
   

    
    If Val(Me.vcodEmpresa.Text) > 0 Then
        cboPuntoDeVenta.Text = "1001"
        
        
        Me.rdFE.Value = True
        Exit Sub
    End If
    
    
    
    If UCase(LeerXml("Login")) = "MANUAL" Then
    MsgBox "No tiene permiso"
    Me.vcodRepartidor.Text = ""
    Me.vcodRepartidor.Text = vCodigoRepartidor2
   ' Exit Sub
End If
   ' vsql = "select (select * from configuracion where id = " + Str(Val(Me.vcodEmpresa.Text) + 1) + ") a "
    
   ' Me.cboPuntoDeVenta.Text = traerDatos2("select * from configuracion", "SucursalDocVenta", PathDBConfig)
End Sub

Private Sub vcodRepartidor_KeyPress(KeyAscii As Integer)
If UCase(LeerXml("Login")) = "MANUAL" Then
   ' MsgBox "No tiene permiso"
    Me.vcodRepartidor.Text = ""
    Me.vcodRepartidor.Text = vCodigoRepartidor2
   ' Exit Sub
End If
End Sub

Private Sub vcodRepartidor_LostFocus()
If UCase(LeerXml("Login")) = "MANUAL" Then
    Me.vcodRepartidor.Text = vCodigoRepartidor2
    Exit Sub
End If

End Sub

Private Sub vdescEmpresa_Change()
Me.vcodEmpresa.Text = Me.vdescEmpresa.Tag
End Sub

Private Sub vDesRepartidor_Change()
If UCase(LeerXml("Login")) = "MANUAL" Then
    Exit Sub
End If

Me.vcodRepartidor.Text = Me.vDesRepartidor.Tag
End Sub

Private Sub vFechaIva_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    focoEnLinea
    'Me.txtDetalle(0).SetFocus
End If

End Sub

Private Sub vnroremito_LostFocus()

End Sub


Private Sub pruebaT2()

Exit Sub

End Sub


Private Sub frellenar_renglones(vremito As Long, vCantidad)

Dim vsql As String
Dim vh, vi As Integer

If Not LeerXml("MostrarSaldoEnDoc") = "SI" Then Exit Sub

vsql = "insert into fdetalle (remito) value (" + Str(vremito) + ")"

If vCantidad < 18 Then
    vh = 17 - vCantidad
Else
    vh = 60 - vCantidad
End If

If Not vCantidad = 18 Then  ' si es 18 no tiene que poner nada
    For vi = 1 To vh
        Call EjecutarScript(vsql, pathDBMySQL)
    Next
End If


End Sub

