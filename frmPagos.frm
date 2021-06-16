VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2008.ocx"
Begin VB.Form frmPagos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gestion de Pagos"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   18150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   18150
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMontoTotalPendienteSeleccionado 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6960
      TabIndex        =   55
      Top             =   3720
      Width           =   1455
   End
   Begin Grid.KlexGrid KlexDetalle 
      Height          =   1455
      Left            =   240
      TabIndex        =   54
      ToolTipText     =   "Documentos a cobrar"
      Top             =   4080
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   2566
      EnterKeyBehaviour=   0
      BackColorAlternate=   16761024
      GridLinesFixed  =   2
      AllowUserResizing=   1
      BackColor       =   12640511
      BackColorFixed  =   -2147483626
      Cols            =   8
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
      MouseIcon       =   "frmPagos.frx":0000
      SelectionMode   =   1
   End
   Begin XtremeSuiteControls.Resizer Resizer 
      Height          =   3495
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   8655
      _Version        =   851968
      _ExtentX        =   15266
      _ExtentY        =   6165
      _StockProps     =   1
      VScrollLargeChange=   1500
      VScrollSmallChange=   140
      VScrollMaximum  =   8000
      HScrollSmallChange=   140
      BorderStyle     =   4
      Begin VB.Frame fraDepositos 
         Caption         =   "Deposito"
         Height          =   1845
         Left            =   120
         TabIndex        =   77
         Top             =   6005
         Width           =   8175
         Begin XtremeSuiteControls.FlatEdit txtDepositoBanco 
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   78
            Top             =   240
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   3
            Left            =   1750
            TabIndex        =   79
            Tag             =   "BancoDeposito"
            Top             =   240
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   4
            Left            =   1750
            TabIndex        =   80
            Tag             =   "BancoCuentaDeposito"
            Top             =   600
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoBanco 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   81
            Top             =   240
            Width           =   5775
            _Version        =   851968
            _ExtentX        =   10186
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoBanco 
            Height          =   315
            Index           =   2
            Left            =   1200
            TabIndex        =   82
            Top             =   600
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoBanco 
            Height          =   315
            Index           =   3
            Left            =   2160
            TabIndex        =   83
            Top             =   600
            Width           =   5775
            _Version        =   851968
            _ExtentX        =   10186
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoImporte 
            Height          =   315
            Left            =   1200
            TabIndex        =   84
            Top             =   960
            Width           =   2655
            _Version        =   851968
            _ExtentX        =   4683
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtDepositoComentario 
            Height          =   315
            Left            =   1200
            TabIndex        =   85
            Top             =   1320
            Width           =   6735
            _Version        =   851968
            _ExtentX        =   11880
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label lblDeposito 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Banco/Caja:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   89
            Top             =   285
            Width           =   1100
         End
         Begin VB.Label lblDeposito 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cuenta: "
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   88
            Top             =   645
            Width           =   1100
         End
         Begin VB.Label lblDeposito 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Importe: "
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   87
            Top             =   1000
            Width           =   1100
         End
         Begin VB.Label lblDeposito 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Comentario: "
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   86
            Top             =   1360
            Width           =   1100
         End
      End
      Begin VB.Frame fraTarjeta 
         Caption         =   "Tarjeta"
         Height          =   1815
         Left            =   120
         TabIndex        =   43
         Top             =   4168
         Width           =   8175
         Begin VB.TextBox txtCantCuotas 
            Height          =   285
            Left            =   3840
            TabIndex        =   53
            Top             =   1440
            Width           =   975
         End
         Begin VB.ComboBox cboBancoTarjeta 
            Height          =   315
            ItemData        =   "frmPagos.frx":001C
            Left            =   1080
            List            =   "frmPagos.frx":001E
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   51
            Top             =   720
            Width           =   6855
         End
         Begin VB.TextBox txtNroCuponTarjeta 
            Height          =   285
            Left            =   1080
            TabIndex        =   46
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox cboTarjeta 
            Height          =   315
            Left            =   1080
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   45
            Top             =   1080
            Width           =   3735
         End
         Begin VB.TextBox txtImporteCuponTarjeta 
            Height          =   285
            Left            =   1080
            TabIndex        =   44
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblCantCuotas 
            Caption         =   "Cant. cuotas:"
            Height          =   255
            Left            =   2760
            TabIndex        =   52
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label lblNroCupón 
            Caption         =   "Nro. cupón:"
            Height          =   255
            Left            =   120
            TabIndex        =   50
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblTarjeta 
            Caption         =   "Tarjeta:"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   1080
            Width           =   615
         End
         Begin VB.Label lblBanco 
            Caption         =   "Banco:"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   720
            Width           =   615
         End
         Begin VB.Label lblImporte 
            Caption         =   "Importe:"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   1440
            Width           =   735
         End
      End
      Begin VB.Frame fraCheques 
         Caption         =   "Cheques"
         Height          =   3255
         Left            =   120
         TabIndex        =   28
         Top             =   840
         Width           =   8175
         Begin XtremeSuiteControls.PushButton cmdEliminarCheque 
            Height          =   315
            Left            =   6840
            TabIndex        =   29
            Top             =   1320
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Eliminar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton cmdAgregarCheque 
            Height          =   315
            Left            =   6840
            TabIndex        =   32
            Top             =   1020
            Width           =   1095
            _Version        =   851968
            _ExtentX        =   1931
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "Agregar"
            UseVisualStyle  =   -1  'True
         End
         Begin Aplisoft_CajasDeTexto.TxF dtpDepositoCheque 
            Height          =   315
            Left            =   1200
            TabIndex        =   69
            Top             =   960
            Width           =   1575
            _ExtentX        =   2778
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
            Appearance      =   0
            BackStyle       =   0
         End
         Begin XtremeSuiteControls.FlatEdit txtFirmanteCheque 
            Height          =   315
            Left            =   5160
            TabIndex        =   68
            Top             =   615
            Width           =   2895
            _Version        =   851968
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.ComboBox cboEstadoCheque 
            Height          =   315
            Left            =   5160
            TabIndex        =   67
            Top             =   960
            Width           =   1575
            _Version        =   851968
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
         End
         Begin XtremeSuiteControls.FlatEdit txtNroCheque 
            Height          =   315
            Left            =   1200
            TabIndex        =   66
            Top             =   615
            Width           =   2895
            _Version        =   851968
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtImporteCheque 
            Height          =   315
            Left            =   1200
            TabIndex        =   65
            Top             =   1320
            Width           =   2895
            _Version        =   851968
            _ExtentX        =   5106
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.TextBox txtImporteTotalCheque 
            Height          =   285
            Left            =   6600
            TabIndex        =   30
            Top             =   2880
            Width           =   1455
         End
         Begin MSDataGridLib.DataGrid dgCheques 
            Height          =   975
            Left            =   120
            TabIndex        =   31
            Top             =   1800
            Width           =   7935
            _ExtentX        =   13996
            _ExtentY        =   1720
            _Version        =   393216
            HeadLines       =   1
            RowHeight       =   15
            AllowDelete     =   -1  'True
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
                  LCID            =   1034
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
                  LCID            =   1034
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
         Begin XtremeSuiteControls.FlatEdit txtNroInternoCheque 
            Height          =   315
            Left            =   5160
            TabIndex        =   63
            Top             =   1320
            Width           =   1575
            _Version        =   851968
            _ExtentX        =   2778
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Alignment       =   2
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   0
            Left            =   1200
            TabIndex        =   71
            Top             =   240
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   0
            Left            =   1750
            TabIndex        =   72
            Tag             =   "Banco"
            Top             =   240
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   1
            Left            =   2160
            TabIndex        =   73
            Top             =   240
            Width           =   1935
            _Version        =   851968
            _ExtentX        =   3413
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   2
            Left            =   5160
            TabIndex        =   74
            Top             =   240
            Width           =   495
            _Version        =   851968
            _ExtentX        =   873
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   1
            Left            =   5700
            TabIndex        =   75
            Tag             =   "BancoCuenta"
            Top             =   240
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.FlatEdit txtBancoCheque 
            Height          =   315
            Index           =   3
            Left            =   6105
            TabIndex        =   76
            Top             =   240
            Width           =   1935
            _Version        =   851968
            _ExtentX        =   3413
            _ExtentY        =   556
            _StockProps     =   77
            BackColor       =   -2147483643
            Appearance      =   6
            UseVisualStyle  =   0   'False
         End
         Begin VB.Label lblCheques 
            Caption         =   "Cuenta:"
            Height          =   195
            Index           =   1
            Left            =   4150
            TabIndex        =   70
            Top             =   285
            Width           =   1000
         End
         Begin VB.Label lblCheques 
            Caption         =   "Nro Interno:"
            Height          =   195
            Index           =   7
            Left            =   4150
            TabIndex        =   64
            Top             =   1365
            Width           =   1000
         End
         Begin VB.Label lblCheques 
            Caption         =   "Estado:"
            Height          =   195
            Index           =   5
            Left            =   4150
            TabIndex        =   39
            Top             =   1005
            Width           =   1000
         End
         Begin VB.Label lblCheques 
            Caption         =   "Banco:"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   38
            Top             =   280
            Width           =   1000
         End
         Begin VB.Label lblCheques 
            Caption         =   "Nro. cheque:"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   37
            Top             =   655
            Width           =   1000
         End
         Begin VB.Label lblCheques 
            Caption         =   "Fec. depósito:"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   36
            Top             =   1000
            Width           =   1000
         End
         Begin VB.Label lblCheques 
            Caption         =   "Firmante:"
            Height          =   195
            Index           =   3
            Left            =   4150
            TabIndex        =   35
            Top             =   675
            Width           =   1000
         End
         Begin VB.Label lblCheques 
            Caption         =   "Importe:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   240
            Index           =   6
            Left            =   90
            TabIndex        =   34
            Top             =   1360
            Width           =   1000
         End
         Begin VB.Label lblTotalCheque 
            Caption         =   "Total cheque:"
            Height          =   255
            Left            =   5520
            TabIndex        =   33
            Top             =   2880
            Width           =   1095
         End
      End
      Begin XtremeSuiteControls.FlatEdit txtCotizacionDolar 
         Height          =   315
         Left            =   4320
         TabIndex        =   58
         Top             =   480
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtImporteEfectivoDolar 
         Height          =   315
         Left            =   2040
         TabIndex        =   60
         Top             =   480
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtImporteEfectivoPesos 
         Height          =   315
         Left            =   2040
         TabIndex        =   59
         Top             =   120
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtNroInterno 
         Height          =   315
         Left            =   6840
         TabIndex        =   62
         Top             =   480
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   300
         Left            =   6840
         TabIndex        =   97
         Top             =   120
         Width           =   1335
         _ExtentX        =   2355
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
         Appearance      =   0
         BackStyle       =   0
      End
      Begin VB.Label lblCobros 
         Caption         =   "Nro Interno:"
         Height          =   255
         Index           =   4
         Left            =   5760
         TabIndex        =   61
         Top             =   525
         Width           =   975
      End
      Begin VB.Label lblCobros 
         Caption         =   "Fecha :"
         Height          =   255
         Index           =   3
         Left            =   5760
         TabIndex        =   57
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblCobros 
         Caption         =   "Cotizacion:"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   42
         Top             =   525
         Width           =   855
      End
      Begin VB.Label lblCobros 
         Caption         =   "Importe efectivo dólar:"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   41
         Top             =   525
         Width           =   1695
      End
      Begin VB.Label lblCobros 
         Caption         =   "Importe efectivo pesos:"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   40
         Top             =   160
         Width           =   1815
      End
   End
   Begin VB.TextBox TxtTotalAPagar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   26
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   9000
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   8415
      Begin VB.TextBox txtTotal 
         Height          =   285
         Left            =   840
         TabIndex        =   24
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPendiente 
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   3960
         TabIndex        =   23
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtPagado 
         ForeColor       =   &H0000C000&
         Height          =   285
         Left            =   6600
         TabIndex        =   22
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtTipoComp 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6600
         TabIndex        =   19
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtNroComprobante 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         TabIndex        =   16
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   5880
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblNroComprobante 
         Caption         =   "Nro. comprobante:"
         Height          =   255
         Left            =   2520
         TabIndex        =   17
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label lblParcialA 
         Caption         =   "Importe:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblPendiente 
         Caption         =   "Pendiente:"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblPagado 
         Caption         =   "Pagado:"
         Height          =   255
         Left            =   5880
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmPagos.frx":0020
      ScaleHeight     =   555
      ScaleWidth      =   8895
      TabIndex        =   0
      Top             =   6760
      Width           =   8900
      Begin XtremeSuiteControls.PushButton cmdPagos 
         Height          =   375
         Left            =   6360
         TabIndex        =   98
         Top             =   90
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ejecutar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusBuscarCliente 
         Height          =   375
         Left            =   2160
         TabIndex        =   21
         Top             =   90
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar Proveedor"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusBuscarDocumento 
         Height          =   375
         Left            =   3600
         TabIndex        =   20
         Top             =   90
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2566
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Buscar documento"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PusGrabar 
         Height          =   375
         Index           =   0
         Left            =   5040
         TabIndex        =   1
         Top             =   90
         Visible         =   0   'False
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Ejecutar"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPagos.frx":50D3
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PusCerrar 
         Height          =   375
         Index           =   1
         Left            =   7560
         TabIndex        =   2
         Top             =   90
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPagos.frx":54DA
      End
      Begin VB.Label lblWGESTION2010 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
         TabIndex        =   3
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGESTION2010 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
         TabIndex        =   4
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.TabControl TabCobros 
      Height          =   4215
      Left            =   -8670
      TabIndex        =   5
      Top             =   -270
      Width           =   8415
      _Version        =   851968
      _ExtentX        =   14843
      _ExtentY        =   7435
      _StockProps     =   68
      Color           =   4
      ItemCount       =   3
      SelectedItem    =   1
      Item(0).Caption =   "Efectivo"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "Label1(0)"
      Item(0).Control(1)=   "txtImporteEfectivo"
      Item(1).Caption =   "Cheques"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "txtImporteTarjeta(1)"
      Item(2).Caption =   "Tarjeta"
      Item(2).ControlCount=   3
      Item(2).Control(0)=   "Pus(1)"
      Item(2).Control(1)=   "Picture1"
      Item(2).Control(2)=   "lblImporteCobrado(2)"
      Begin VB.TextBox txtImporteTarjeta 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         ForeColor       =   &H000000FF&
         Height          =   2.45745e5
         HelpContextID   =   1
         HideSelection   =   0   'False
         Index           =   1
         Left            =   100
         TabIndex        =   15
         Top             =   100
         Width           =   2.45745e5
      End
      Begin VB.TextBox txtImporteEfectivo 
         Height          =   285
         Left            =   -68560
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin XtremeSuiteControls.PushButton Pus 
         Height          =   315
         Index           =   1
         Left            =   -1.35920e5
         TabIndex        =   6
         Tag             =   "CodigoPostal"
         Top             =   1080
         Visible         =   0   'False
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin VB.PictureBox Picture1 
         Height          =   15
         Left            =   -70000
         ScaleHeight     =   15
         ScaleWidth      =   15
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label lblImporteCobrado 
         Caption         =   "Importe cobrado:"
         Height          =   255
         Index           =   2
         Left            =   -69760
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Importe cobrado:"
         Height          =   255
         Index           =   0
         Left            =   -69880
         TabIndex        =   8
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSAdodcLib.Adodc bcheques 
      Height          =   360
      Left            =   9000
      Top             =   2880
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   635
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
      Caption         =   "bcheques"
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
   Begin MSAdodcLib.Adodc bpccliente 
      Height          =   330
      Left            =   9000
      Top             =   2520
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "bpccliente"
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
   Begin MSAdodcLib.Adodc bbanco 
      Height          =   330
      Left            =   9000
      Top             =   1800
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "bbanco"
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
   Begin MSAdodcLib.Adodc bsucursal 
      Height          =   330
      Left            =   9000
      Top             =   1440
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "bsucursal"
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
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   9000
      Top             =   2160
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
      Caption         =   "bccliente"
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
   Begin XtremeSuiteControls.FlatEdit txtObservaciones 
      Height          =   315
      Left            =   1560
      TabIndex        =   90
      Top             =   5640
      Width           =   7095
      _Version        =   851968
      _ExtentX        =   12515
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin XtremeSuiteControls.GroupBox GroDatosCliente 
      Height          =   705
      Left            =   0
      TabIndex        =   91
      Top             =   6000
      Width           =   8895
      _Version        =   851968
      _ExtentX        =   15690
      _ExtentY        =   1235
      _StockProps     =   79
      Caption         =   "Datos cliente"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit txtProveedor 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   92
         Top             =   270
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.FlatEdit txtProveedor 
         Height          =   315
         Index           =   1
         Left            =   3120
         TabIndex        =   93
         Top             =   270
         Width           =   5535
         _Version        =   851968
         _ExtentX        =   9763
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   2
         Left            =   2730
         TabIndex        =   94
         Tag             =   "Proveedores"
         Top             =   270
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre:"
         Height          =   195
         Left            =   120
         TabIndex        =   95
         Top             =   315
         Width           =   1440
      End
   End
   Begin XtremeSuiteControls.Label lblObservaciones 
      Height          =   195
      Left            =   180
      TabIndex        =   96
      Top             =   5685
      Width           =   1365
      _Version        =   851968
      _ExtentX        =   2408
      _ExtentY        =   344
      _StockProps     =   79
      Caption         =   "Observaciones :"
   End
   Begin VB.Label lblMontoTotalPendente 
      Caption         =   "Monto total pendiente de documentos seleccionados:"
      Height          =   255
      Left            =   3120
      TabIndex        =   56
      Top             =   3720
      Width           =   3855
   End
   Begin VB.Label lblTotalA 
      Caption         =   "Total a pagar:"
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "frmPagos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vImportePagoPesos As Double
Public vIdCtaCteP As Long
Public codProveedor As String
Public pendiente, pagado, total As Double
Public remito As Long
Public esComprobanteAutomatico As Boolean
Dim rsCheque As New ADODB.Recordset
Dim rsRecibo As New ADODB.Recordset
Dim concepto As String
Public NroComprobante As Long
Public tipoComprobante As String
Public fechaDocumento As Date
Dim SaldoAnterior, totaldebito, totalCredito, credito, debito As Double
Dim vImporteTotalAPagar As Double
Dim vnrorecibo As Long

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
Private Sub cmdPagos_Click()
On Error Resume Next

    Dim vTotalCredito As Double
    
    vTotalCredito = 0
    
    VaciarTablaRecibo

    If Not Val(TxtTotalAPagar.Text) = Val(vImporteTotalAPagar) Then
        MsgBox "EEEEERRRRRRRRRRROOOOOOOOORRRRRR", vbCritical, "Mensaje ..."
        Exit Sub
    End If
    
    If Val(TxtTotalAPagar.Text) = 0 Then
        MsgBox "El monto a pagar debe ser mayor a 0", vbInformation, "WGestion"
        Exit Sub
    Else
        If Val(txtImporteEfectivoPesos) > 0 Then
    
            Dim importeAPagar, importePagadoPesos As Double
            importePagadoPesos = 0
            vImportePagoPesos = 0
            
            importePagadoPesos = Val(txtImporteEfectivoPesos.Text)
            
            
            'Agrego una linea en el recibo por el monto pagado en efectivo
            Call AgregarPagoRecibo(1, "Efectivo en Pesos:", importePagadoPesos)
            
            vImportePagoPesos = importePagadoPesos
            
        End If
        

        If Val(txtImporteTotalCheque.Text) > 0 Then
            Dim importePagadoCheque As Double
            importePagadoCheque = 0
            'Agrego una linea en el recibo por el monto pagado con cada cheque
            If Not rsCheque.RecordCount = 0 Then
                rsCheque.MoveFirst
                Do While Not rsCheque.EOF = True
                    AgregarPagoRecibo 4, "Cheque Nro.: " & Str(rsCheque.Fields("NCheque").Value) & " de Banco " & rsCheque.Fields("Banco") & "-" & rsCheque.Fields("Sucursal"), Val(rsCheque.Fields("Monto"))
                    importePagadoCheque = Val(importePagadoCheque) + Val(Format(rsCheque.Fields("Monto").Value, "#######0.00"))
                    rsCheque.MoveNext
                Loop
            End If
    
            'Guardo los cheques

    
        End If


        If Val(txtImporteCuponTarjeta.Text) > 0 Then
    
            Dim totalpagadoTarjeta As Double
            totalpagadoTarjeta = 0
            
            If esComprobanteAutomatico = True Then
                totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta.Text)
            Else
                Dim i As Integer
                'Guardo en cta cte
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(Me.txtImporteCuponTarjeta.Text) <= 0 Then
                        totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta.Text)
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (Me.KlexDetalle.TextMatrix(i, 5) > Val(Me.txtImporteCuponTarjeta.Text)) Then
                            Me.txtImporteCuponTarjeta.Text = 0
                            totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta)
                        Else
                            txtImporteCuponTarjeta.Text = Val(Me.txtImporteCuponTarjeta.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                            totalpagadoTarjeta = totalpagadoTarjeta + KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                
            End If
            
            'Agrego una linea en el recibo por el monto pagado con tarjeta
            AgregarPagoRecibo 3, "Tarjeta " & Me.cboTarjeta.Text & " de Banco " & Me.cboBancoTarjeta.Text, totalpagadoTarjeta
    
            Dim idCuponTarjetaNuevo As Integer
    
            'Guardo los datos de la operacion en cupon tarjeta
            idCuponTarjetaNuevo = GuardarCuponTarjeta
    
            'Guardo un movimiento en la cuenta de bancos por el total de lo pagado con tarjeta
            Call GuardarBancosMovimientos(vnrorecibo, cboBancoTarjeta.Tag, 1, 0, totalpagadoTarjeta, "", idCuponTarjetaNuevo, Val(txtNroInterno.Text))
    
    
        End If

        If Val(txtDepositoImporte.Text) > 0 Then
    
            Dim totalPagadoDeposito As Double
            totalPagadoDeposito = 0
            
            If esComprobanteAutomatico Then
                totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
            Else
                'Guardo en cta cte
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(Me.txtDepositoImporte.Text) <= 0 Then
                        totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (Me.KlexDetalle.TextMatrix(i, 5) > Val(txtImporteCuponTarjeta.Text)) Then
                            txtDepositoImporte.Text = 0
                            totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
                        Else
                            txtDepositoImporte.Text = Val(Me.txtDepositoImporte.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                            totalPagadoDeposito = totalPagadoDeposito + Me.KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                        
            End If
            
            'Agrego una linea en el recibo por el monto pagado con Deposito Bancario
            Call AgregarPagoRecibo(5, "Extraccion en la Cuenta " & Trim(txtDepositoBanco(3).Text) & " de Banco " & txtDepositoBanco(1).Text, totalPagadoDeposito)

            'Guardo un movimiento en la cuenta de bancos por el total de lo Depositado
            Call GuardarBancosMovimientos(vnrorecibo, Trim(txtDepositoBanco(0).Text), Val(txtDepositoBanco(2).Text), 0, totalPagadoDeposito, txtDepositoComentario.Text, 0, Val(txtNroInterno.Text))
    
        End If
        
        GuardoCreditoEnCtaCte

        If Err.Number = 0 Then
            MsgBox "El pago se realizo exitosamente", vbInformation, "WGestion"
        Else
            MsgBox Err.Description
        End If
        
        If vConfigGral.vIncluyeContabilidad = True Then CargarContabilidad
        
        If vConfigGral.vImprimirReciboProveedor = True Then ImprimirRecibo
        
        VaciarTablaRecibo
        
        
        HabilitarControles (False)

    End If

If Err Then
    MsgBox "Hubo errores al guardar, consulte el log del sistema", vbInformation, "WGestion"
Else
    Unload Me
End If
End Sub
Private Function GuardoCreditoEnCtaCte() As Double
On Error Resume Next

    Dim vTotalCredito As Double
    
    vTotalCredito = 0
    
    vTotalCredito = GenerarDato("SELECT SUM(Monto) as TotalRecibo FROM Recibo_Temp", "TotalRecibo")
    
    Dim rsCredito As New ADODB.Recordset, sqlCredito As String
    
    sqlCredito = "SELECT * FROM PCuentasCorrientes WHERE 1=2"
    
    With rsCredito
        Call .Open(sqlCredito, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        .AddNew
        
        .Fields("Fecha").Value = strfechaMySQL(Me.dtpFecha.Value)
        .Fields("Codigo").Value = Trim(txtProveedor(0).Text)
        .Fields("Nombre").Value = Trim(txtProveedor(1).Text)
        .Fields("Debito").Value = 0
        .Fields("Credito").Value = vTotalCredito
        .Fields("Comentario").Value = Trim(txtObservaciones.Text)
        .Fields("Remito").Value = Null
        .Fields("NroInterno").Value = Val(txtNroInterno.Text)
        .Fields("idMedioPago").Value = 99
        .Fields("NroAsiento").Value = Null
        .Fields("TipoMovimiento").Value = "RC"
        
        .Update
        
        vIdCtaCteP = .Fields(0).Value
    
    End With

    GuardoCreditoEnCtaCte = vTotalCredito

    sqlCredito = ""
    
    If rsCredito.State = 1 Then
        rsCredito.Close
        Set rsCredito = Nothing
    End If
    
If Err Then GrabarLog "GuardoCreditoEnCtaCte", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub dtpFecha_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        txtNroInterno.SetFocus
    End If


If Err Then GrabarLog "dtpFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag
    
    frmBusqueda.Show

If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtBancoCheque_GotFocus(Index As Integer)
On Error Resume Next

    Resizer.VScrollPosition = 840
    'Call CargarCombo("Bancos", "Descripcion", txtBancoCheque, False) ', Str(idBancos))

If Err Then GrabarLog "cboBanco_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtBancoCheque_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        Select Case Index
        
            Case 0
                txtBancoCheque(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtBancoCheque(Index).Text) & "'", "Descripcion")
                txtBancoCheque(Index + 2).SetFocus
            
            Case 2
                txtBancoCheque(Index + 1).Text = TraerDato("BancosCuentas", "idBancosCuentas = " & Trim(txtBancoCheque(Index).Text) & "", "Cuenta")
                txtNroCheque.SetFocus
        End Select
    
    
        If txtBancoCheque(Index).Text = "" Then txtNroCheque.SetFocus
    
    End If

If Err Then GrabarLog "cboBanco_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboBancoTarjeta_Click()
On Error Resume Next

    Call CargarComboTarjetaPorBanco("Tarjeta", "Nombre", cboTarjeta, False, "Nombre", cboBancoTarjeta.Text)
    cboTarjeta.SetFocus

If Err Then GrabarLog "cboBancoTarjeta_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboBancoTarjeta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        Me.cboTarjeta.SetFocus
    End If
        
End Sub

Private Sub cboEstadoCheque_GotFocus()
On Error Resume Next

    Call CargarComboNew("EstadoCheque", "Descripcion", cboEstadoCheque, False) ', Str(idEstadoCheque))

If Err Then GrabarLog "cboEstadoCheque_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cboEstadoCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtImporteCheque.SetFocus
    End If
End Sub
Private Sub cboTarjeta_GotFocus()
On Error Resume Next

    Call CargarComboTarjetaPorBanco("Tarjeta", "Nombre", cboTarjeta, False, "idBancos", Me.cboBancoTarjeta.Tag) ', Str(idTarjeta))

If Err Then GrabarLog "cboTarjeta_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cboTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtImporteCuponTarjeta.SetFocus
End If
End Sub
Private Sub dgCheques_AfterDelete()
    'rsCheque.Delete
End Sub
Private Sub dgCheques_BeforeDelete(Cancel As Integer)
    Me.txtImporteTotalCheque.Text = Val(Me.txtImporteTotalCheque.Text) - rsCheque.Fields("monto").Value
End Sub
Private Sub dgCheques_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        EliminarFilaCheque
    End If
End Sub
Private Sub dtpDepositoCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cboEstadoCheque.SetFocus
    End If
End Sub
Private Sub Form_Load()
On Error Resume Next

    txtTotal.Text = total
    
    With Me
        .Top = 0
        .Left = 0
        .Height = 7635
        .Width = 9060
        .KeyPreview = True
    End With
    
    Resizer.HScrollPosition = 0
    Resizer.VScrollPosition = 0
    
    Call CargarCombo("Bancos", "Descripcion", cboBancoTarjeta, False) ', Str(idBancos))
    
    txtCotizacionDolar.Text = ObtenerCotizacionMoneda("002", True)
    
    LimpiarCampos
    
    HabilitarControles (False)
    
    FormatoGrillaDetalle (1)

    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub VaciarChequesTemp()
On Error Resume Next

    Dim sqlCheque As String
    
    With rsCheque
         sqlCheque = "DELETE FROM cheques_temp"
         If .State = 0 Then
            .CursorLocation = adUseClient
            .Open sqlCheque, ConnDDBB, adOpenDynamic, adLockPessimistic
        Else
            Set rsCheque = ConnDDBB.Execute(sqlCheque)
        End If
        
        Set dgCheques.DataSource = rsCheque
    End With

If Err Then GrabarLog "VaciarChequesTemp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub GuardarCheques()
On Error Resume Next

    Dim rsChequeGuardar As New ADODB.Recordset, sqlCheque As String
    Dim i As Integer

    rsCheque.Requery
    
    If Not rsCheque.EOF Then
    
        sqlCheque = "SELECT * FROM cheques"
    
        With rsChequeGuardar
            .CursorLocation = adUseClient
            Call .Open(sqlCheque, ConnDDBB, adOpenStatic, adLockPessimistic)
    
            Do Until rsCheque.EOF = True
                .AddNew
                For i = 1 To rsCheque.Fields.Count - 3
                    .Fields(i).Value = EsNulo(rsCheque.Fields(i).Value)
                Next i
                .Update
                rsCheque.MoveNext
            Loop

        End With
        
        sqlCheque = ""

        If rsChequeGuardar.State = 1 Then
            rsChequeGuardar.Close
            Set rsChequeGuardar = Nothing
        End If
        
    End If

If Err Then GrabarLog "GuardarCheques", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    VaciarChequesTemp

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdAgregarCheque_Click()
    On Error Resume Next

    If (Val(remito) = 0) And (Trim(codProveedor) = "") Then
        MsgBox "Debe seleccionar un Proveedor o un remito antes de iniciar la operacion", vbOKOnly, "Mensaje ..."
    Else
        If ValidarIngresoCheque = True Then
            Dim sqlCheque As String
        
        
            With rsCheque
                sqlCheque = "SELECT * FROM cheques_temp"
                If .State = 0 Then
                    .CursorLocation = adUseClient
                    Call .Open(sqlCheque, ConnDDBB, adOpenStatic, adLockPessimistic)
                    If Not .State = 1 Then
                        MsgBox "No Pudo abrirse la DDBB", vbExclamation, "Mensaje ..."
                        Exit Sub
                    End If
                    
                End If
                
                .AddNew
                
                .Fields("idEstadoCheque").Value = Val(cboEstadoCheque.Tag)
                .Fields("Codigo").Value = EsNulo(txtProveedor(0).Text)
                .Fields("Nombre").Value = EsNulo(txtProveedor(1).Text)
                .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
                .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
                .Fields("FechaDeposito").Value = strfechaMySQL(dtpDepositoCheque.Value)
                .Fields("FechaAcreditacion").Value = strfechaMySQL(dtpDepositoCheque.Value)
                
                .Fields("Monto").Value = Val(txtImporteCheque.Text)
                
                '.Fields("Banco").Value = EsNulo(cboBancoCheque.Text)
                '.Fields("Codigo").Value = EsNulo(cboBancoCheque.Tag)
                  
                .Fields("NCheque").Value = Left(Trim(txtNroCheque.Text), 20)
            
                .Fields("Remito").Value = remito
                .Fields("Firmante").Value = Trim(txtFirmanteCheque.Text)
                .Fields("NroInterno").Value = Val(txtNroInternoCheque.Text)
                .Update
                
                txtImporteTotalCheque.Text = GenerarDato("SELECT SUM(Monto) as TotalCheques FROM Cheques_Temp", "TotalCheques")
            
                FormatoGrillaCheques
                
                LimpiarCheques
                
                txtBancoCheque(0).SetFocus
    
            End With
    
        End If
        
        'Set dgCheques.DataSource = rsCheque
    End If
    
If Err Then GrabarLog "cmdAgregarCheque_Click", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Private Sub FormatoGrillaCheques()
On Error Resume Next

    With dgCheques
        Set .DataSource = rsCheque
        
        .Columns(0).Width = 0
        .Columns(1).Width = 0
        .Columns(2).Width = 1000
        .Columns(3).Width = 0
        .Columns(4).Width = 0
        .Columns(5).Width = 1000
        .Columns(6).Width = 1000
        
        .Columns(7).Width = 750
        
        .Columns(8).Width = 750
        .Columns(9).Width = 750
        .Columns(10).Width = 750
        .Columns(11).Width = 750
        .Columns(12).Width = 0
        .Columns(13).Width = 0
        .Columns(14).Width = 0
    
        .Columns(15).Width = 0
        .Columns(16).Width = 0
        .Columns(17).Width = 0
        .Columns(18).Width = 0
    
    End With
    
If Err Then GrabarLog "FormatoGrillaCheques", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub LimpiarCheques()
On Error Resume Next

    txtBancoCheque(0).Text = ""
    txtBancoCheque(1).Text = ""
    txtBancoCheque(2).Text = ""
    txtBancoCheque(3).Text = ""
    txtNroCheque.Text = ""
    txtFirmanteCheque.Text = ""
    dtpDepositoCheque.Value = Date
    cboEstadoCheque.Text = ""
    cboEstadoCheque.Tag = ""
    txtImporteCheque.Text = ""
    txtNroCheque.Text = ""

If Err Then GrabarLog "LimpiarCheques", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdAgregarCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAgregarCheque_Click
    End If
End Sub

Private Sub PusBuscarCliente_Click()
On Error Resume Next

    With frmProveedores
        .Show
        .vienePago = True
        .txtBuscar.SetFocus
    End With
    
    HabilitarControles (True)

If Err Then GrabarLog "PusBuscarCliente_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PusBuscarDocumento_Click()
On Error Resume Next

    With frmBuscarCompra
        .Show
        
        If Trim(codProveedor) <> "" Then
            .txtProveedor(0).Text = Trim(txtProveedor(0).Text)
            .txtProveedor(1).Text = Trim(txtProveedor(1).Text)
        End If
        
        .cmdEjecutarPago.Enabled = True
        
        .vienePago = True
        
        .cmdBuscaryCalcular_Click
        
    End With
    
    HabilitarControles (True)
    

If Err Then GrabarLog "PusBuscarDocumento_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PusCerrar_Click(Index As Integer)
On Error Resume Next

    Unload Me
    
If Err Then GrabarLog "PusCerrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdEliminarCheque_Click()
On Error Resume Next

    EliminarFilaCheque

If Err Then GrabarLog "cmdEliminarCheque_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub EliminarFilaCheque()
On Error Resume Next

    With rsCheque
        If Not .EOF = True Then
            txtImporteTotalCheque.Text = Val(txtImporteTotalCheque.Text) - Val(rsCheque.Fields("monto").Value)
            rsCheque.Delete
        End If
    End With

If Err Then GrabarLog "EliminarFilaCheque", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PusGrabar_Click(Index As Integer)
On Error Resume Next

    'vImpresionCorrecta = False

    VaciarTablaRecibo

    If Val(TxtTotalAPagar.Text) = 0 Then
        MsgBox "El monto a pagar debe ser mayor a 0", vbInformation, "WGestion"
        Exit Sub
    Else
        If Val(txtImporteEfectivoPesos) > 0 Then
    
            Dim importeAPagar, importePagadoPesos As Double
            importePagadoPesos = 0
            vImportePagoPesos = 0
            
            If esComprobanteAutomatico = True Then
                Call PagarCtaCteAutomaticamente(Val(txtImporteEfectivoPesos.Text), 1)
                importePagadoPesos = Val(txtImporteEfectivoPesos.Text)
                
            Else
                'Guardo en cta cte
                Dim i As Integer
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(txtImporteEfectivoPesos.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (Val(KlexDetalle.TextMatrix(i, 5)) > Val(txtImporteEfectivoPesos.Text)) Then
                            Call PagarCtaCteProveedor2(KlexDetalle.TextMatrix(i, 7), Val(txtImporteEfectivoPesos.Text), dtpFecha.Value, 1)
                            importePagadoPesos = importePagadoPesos + Val(txtImporteEfectivoPesos.Text)
                            'vImportePagoPesos = Val(importePagadoPesos)
                            txtImporteEfectivoPesos.Text = 0
                        Else
                            Call PagarCtaCteProveedor2(KlexDetalle.TextMatrix(i, 7), Me.KlexDetalle.TextMatrix(i, 5), dtpFecha.Value, 1)
                            txtImporteEfectivoPesos.Text = Val(Me.txtImporteEfectivoPesos.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                            importePagadoPesos = importePagadoPesos + KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                        'Agrego una linea en el recibo por el monto pagado en efectivo
                        Call AgregarPagoRecibo(1, "Efectivo en Pesos:", importePagadoPesos)
                
                        'Guardo el importe en efectivo en caja
                        Call WCaja(importePagadoPesos)
                        
                        vImportePagoPesos = importePagadoPesos
            
                    End If
                Next i
                
            End If
            
        End If
        

        If Val(txtImporteTotalCheque.Text) > 0 Then
            Dim importePagadoCheque As Double
            importePagadoCheque = 0
            'Agrego una linea en el recibo por el monto pagado con cada cheque
            If Not rsCheque.RecordCount = 0 Then
                rsCheque.MoveFirst
                Do While Not rsCheque.EOF = True
                    AgregarPagoRecibo 4, "Cheque Nro.: " & Str(rsCheque.Fields("NCheque").Value) & " de Banco " & rsCheque.Fields("Banco") & "-" & rsCheque.Fields("Sucursal"), Val(rsCheque.Fields("Monto"))
                    importePagadoCheque = Val(importePagadoCheque) + Val(Format(rsCheque.Fields("Monto").Value, "#######0.00"))
                    rsCheque.MoveNext
                Loop
            End If
    
            'Guardo los cheques
            GuardarCheques
    
            If esComprobanteAutomatico Then
                PagarCtaCteAutomaticamente (Val(txtImporteTotalCheque.Text)), 4
            'Else
                'Guardo en cta cte
            '    Call PagarCtaCte2(remito, Val(txtImporteCheque), Me.dtpFecha.Value, 4)
            'End If
            Else
                'Guardo en cta cte
                
                For i = 1 To KlexDetalle.Rows - 1
                    If Val(txtImporteTotalCheque.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (KlexDetalle.TextMatrix(i, 5) > Val(txtImporteTotalCheque.Text)) Then
                            Call PagarCtaCteProveedor2(KlexDetalle.TextMatrix(i, 7), Val(Me.txtImporteTotalCheque), dtpFecha.Value, 4)
                            Me.txtImporteEfectivoPesos.Text = 0
                        Else
                            Call PagarCtaCteProveedor2(KlexDetalle.TextMatrix(i, 7), Me.KlexDetalle.TextMatrix(i, 5), dtpFecha.Value, 4)
                            txtImporteTotalCheque.Text = Val(Me.txtImporteTotalCheque.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                
            End If
    
        End If


        If Val(txtImporteCuponTarjeta.Text) > 0 Then
    
            Dim totalpagadoTarjeta As Double
            totalpagadoTarjeta = 0
            
            If Me.esComprobanteAutomatico Then
                Me.PagarCtaCteAutomaticamente (Val(Me.txtImporteCuponTarjeta.Text)), 3
                totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta.Text)
            Else
                'Guardo en cta cte
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(Me.txtImporteCuponTarjeta.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                        totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta.Text)
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (Me.KlexDetalle.TextMatrix(i, 5) > Val(Me.txtImporteCuponTarjeta.Text)) Then
                            Call PagarCtaCteProveedor2(KlexDetalle.TextMatrix(i, 7), Val(Me.txtImporteCuponTarjeta), dtpFecha.Value, 3)
                            Me.txtImporteCuponTarjeta.Text = 0
                            totalpagadoTarjeta = totalpagadoTarjeta + Val(Me.txtImporteCuponTarjeta)
                        Else
                            Call PagarCtaCteProveedor2(KlexDetalle.TextMatrix(i, 7), Me.KlexDetalle.TextMatrix(i, 5), dtpFecha.Value, 3)
                            txtImporteCuponTarjeta.Text = Val(Me.txtImporteCuponTarjeta.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                            totalpagadoTarjeta = totalpagadoTarjeta + KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                
            End If
            
            'Agrego una linea en el recibo por el monto pagado con tarjeta
            AgregarPagoRecibo 3, "Tarjeta " & Me.cboTarjeta.Text & " de Banco " & Me.cboBancoTarjeta.Text, totalpagadoTarjeta
    
            Dim idCuponTarjetaNuevo As Integer
    
            'Guardo los datos de la operacion en cupon tarjeta
            idCuponTarjetaNuevo = GuardarCuponTarjeta
    
            'Guardo un movimiento en la cuenta de bancos por el total de lo pagado con tarjeta
            Call GuardarBancosMovimientos(vnrorecibo, cboBancoTarjeta.Tag, 1, 0, totalpagadoTarjeta, "", idCuponTarjetaNuevo, Val(txtNroInterno.Text))
    
    
        End If

        If Val(txtDepositoImporte.Text) > 0 Then
    
            Dim totalPagadoDeposito As Double
            totalPagadoDeposito = 0
            
            If esComprobanteAutomatico Then
                Call PagarCtaCteAutomaticamente(Val(txtDepositoImporte.Text), 5)
                totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
            Else
                'Guardo en cta cte
                For i = 1 To Me.KlexDetalle.Rows - 1
                    If Val(Me.txtDepositoImporte.Text) <= 0 Then
                        MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " no pudo ser pagados", vbInformation, "WGestion"
                        totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
                    Else
                        'Si el total a pagar es menor que lo pendiente pago ese total, sino todo lo pendiente de ese documento
                        If (Me.KlexDetalle.TextMatrix(i, 5) > Val(txtImporteCuponTarjeta.Text)) Then
                            Call PagarCtaCteProveedor2(KlexDetalle.TextMatrix(i, 7), Val(Me.txtDepositoImporte.Text), dtpFecha.Value, 5)
                            txtDepositoImporte.Text = 0
                            totalPagadoDeposito = totalPagadoDeposito + Val(txtDepositoImporte.Text)
                        Else
                            Call PagarCtaCteProveedor2(KlexDetalle.TextMatrix(i, 7), Me.KlexDetalle.TextMatrix(i, 5), dtpFecha.Value, 5)
                            txtDepositoImporte.Text = Val(Me.txtDepositoImporte.Text) - Me.KlexDetalle.TextMatrix(i, 5)
                            totalPagadoDeposito = totalPagadoDeposito + Me.KlexDetalle.TextMatrix(i, 5)
                        End If
                        
                    End If
                Next i
                        
            End If
            
            'Agrego una linea en el recibo por el monto pagado con Deposito Bancario
            Call AgregarPagoRecibo(5, "Extraccion en la Cuenta " & Trim(txtDepositoBanco(3).Text) & " de Banco " & txtDepositoBanco(1).Text, totalPagadoDeposito)

            'Guardo un movimiento en la cuenta de bancos por el total de lo Depositado
            Call GuardarBancosMovimientos(vnrorecibo, Trim(txtDepositoBanco(0).Text), Val(txtDepositoBanco(2).Text), totalPagadoDeposito, 0, txtDepositoComentario.Text, 0, Val(txtNroInterno.Text))
    
        End If
        
        If Err.Number = 0 Then
            MsgBox "El pago se realizo exitosamente", vbInformation, "WGestion"
        Else
            MsgBox Err.Description
        End If
        
        If vConfigGral.vIncluyeContabilidad = True Then CargarContabilidad
        
        If vConfigGral.vImprimirReciboProveedor = True Then ImprimirRecibo
        
        VaciarTablaRecibo
        
        KlexDetalle.Rows = 1
        HabilitarControles (False)

    End If

If Err Then
    MsgBox "Hubo errores al guardar, consulte el log del sistema", vbInformation, "WGestion"
Else
    Unload Me
End If
End Sub
Private Sub CargarContabilidad()
On Error Resume Next

    With frmAsientosAlta
        .Show
        .chkControlar.Value = xtpChecked
        .txtCuentaVieneDe.Text = Me.Caption
        .txtCuentaVieneDe.Tag = Trim(txtProveedor(0).Text)
        If txtObservaciones.Text = "" Then
            .txtLeyenda.Text = "Pago: " & txtTipoComp.Text & " " & txtNroComprobante.Text
        Else
            .txtLeyenda.Text = Trim(txtObservaciones.Text)
        End If
        .dtpFecha.Value = dtpFecha.Value
        
        
        'Panic
        .txtImporteVieneDe.Text = Val(vImporteTotalAPagar)
        
        
        .lblNroInterno.Caption = Val(txtNroInterno.Text)
        
        .cboTipoMovimiento.Tag = "RC"
        .cboTipoMovimiento.Text = "Recibo de Cobro"
        
        .vVieneTabla = "PCuentasCorrientes"
        .vVieneIdNombre = "id"
        .vVieneIdValor = vIdCtaCteP

    End With

If Err Then GrabarLog "CargarContabilidad", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub WCaja(importePagado As Double)
    On Error Resume Next
    
    Dim rsCaja As New ADODB.Recordset
    Dim sqlCaja As String
    
    sqlCaja = "SELECT * FROM caja"
    
    With rsCaja
        Call .Open(sqlCaja, ConnDDBB, adOpenDynamic, adLockPessimistic)
    
        .AddNew
        .Fields("remito").Value = Val(remito)
        
        .Fields("fecha").Value = strfechaMySQL(dtpFecha.Value)
        .Fields("Importe").Value = importePagado
        
        .Fields("CodigoProveedor").Value = Trim(codProveedor)
        
        .Fields("Usuario").Value = vConfigGral.vUser
        .Fields("CodigoConcepto").Value = 221
        .Fields("comentario").Value = ""
            
        .Fields("NroCheque") = Null
        .Fields("FechaDeposito") = Null
        .Fields("FechaConfeccion") = Null
        .Fields("idCajas") = Null
        
        .Update
    
    End With
    
    sqlCaja = ""
    
    If rsCaja.State = 1 Then
        rsCaja.Close
        Set rsCaja = Nothing
    End If
    
If Err Then GrabarLog "WCaja", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub

Private Sub txtBancoCheque_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        If Index = 2 Then
            pbCarga_Click (1)
        End If
    End If
    
If Err Then GrabarLog "txtBancoCheque_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtCantCuotas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.PusGrabar(0).SetFocus
End If
End Sub

Private Sub txtCotizacionDolar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtBancoCheque(0).SetFocus
End If
End Sub

Private Sub txtDepositoBanco_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        Select Case Index
        
            Case 0
                txtDepositoBanco(Index + 1).Text = TraerDato("Bancos", "idBancos = '" & Trim(txtDepositoBanco(Index).Text) & "'", "Descripcion")
                txtDepositoBanco(Index + 2).SetFocus
            
            Case 2
                txtDepositoBanco(Index + 1).Text = TraerDato("BancosCuentas", "idBancosCuentas = " & Trim(txtDepositoBanco(Index).Text) & "", "Cuenta")
                txtDepositoImporte.SetFocus
        End Select
    
    
        If txtDepositoBanco(Index).Text = "" Then txtDepositoImporte.SetFocus
    
    End If

If Err Then GrabarLog "txtDepositoBanco_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtDepositoBanco_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    If KeyCode = vbKeyF3 Then
        If Index = 2 Then
            pbCarga_Click (4)
        End If
    End If
    
If Err Then GrabarLog "txtBancoCheque_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtFirmanteCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.dtpDepositoCheque.SetFocus
    End If
End Sub

Private Sub txtImporteCuponTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtCantCuotas.SetFocus
End If
End Sub

Private Sub txtImporteCheque_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAgregarCheque.SetFocus
    End If
End Sub

Private Sub txtImporteEfectivoPesos_Change()
    TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos.Text) + Val(txtImporteCuponTarjeta.Text) + Val(txtDepositoImporte.Text)
End Sub
Private Sub txtImporteCuponTarjeta_Change()
    TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos) + Val(txtImporteCuponTarjeta) + Val(txtDepositoImporte.Text)
End Sub
Private Sub txtdepositoImporte_Change()
    TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos) + Val(txtImporteCuponTarjeta) + Val(txtDepositoImporte.Text)
End Sub
Private Sub txtImporteEfectivoDolar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        If Me.txtImporteEfectivoDolar.Text = "" Then
            Me.txtBancoCheque(0).SetFocus
        Else
            Me.txtCotizacionDolar.SetFocus
        End If
    End If
End Sub

Private Sub txtImporteEfectivoPesos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtImporteEfectivoDolar.SetFocus
    End If
End Sub
Private Sub txtImporteTotalCheque_Change()
    TxtTotalAPagar.Text = Val(txtImporteTotalCheque.Text) + Val(txtImporteEfectivoPesos.Text) + Val(txtImporteCuponTarjeta.Text) + Val(txtDepositoImporte.Text)
End Sub

Private Sub txtNroCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.txtFirmanteCheque.SetFocus
End If
End Sub

Private Sub txtNroComprobante_Change()

    'Me.txtPendiente = CalcularSaldo(remito)
    'Me.txtTotal.Text = CalcularTotal(remito)
    'Me.txtPagado.Text = CalcularPagado(remito)

End Sub

Public Sub AgregarDocumentoAPagar(total As Double, pendiente As Double, pagado As Double, nroComp As Long, tipoComp As String, fechaComp As Date, remito As Integer)
    On Error Resume Next
    Dim i, j As Integer
    
    With KlexDetalle
        If .Rows <= 2 And .TextMatrix(.Rows - 1, 4) = "" Then
            FormatoGrillaDetalle (1)
        Else
            .Rows = .Rows + 1
        End If
        j = .Rows - 1
        
        .TextMatrix(j, 1) = fechaComp
        .TextMatrix(j, 2) = tipoComp
        .TextMatrix(j, 3) = nroComp
        .TextMatrix(j, 4) = total
        .TextMatrix(j, 5) = pendiente
        .TextMatrix(j, 6) = pagado
        .TextMatrix(j, 7) = remito

    End With
    
    If Err Then GrabarLog "AgregarDocumentoAPagar", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub
Public Sub BuscarDatosOperacionesProveedor(codProv As String, remito As Long)
    On Error Resume Next
    
    If LeerConfig(24) = True Then
        Dim rsCtaCteP As New ADODB.Recordset, sqlCtaCteP As String, i As Integer
    
    SaldoAnterior = 0
    totaldebito = 0
    totalCredito = 0
    credito = 0
    debito = 0
    If remito <> 0 Then
        sqlCtaCteP = "SELECT * FROM PCuentasCorrientes WHERE (codigo = '" & codProv & "') and (remito = " & remito & ")"
    Else
        sqlCtaCteP = "SELECT * FROM PCuentasCorrientes WHERE (codigo = '" & codProv & "')"
    End If
    
    With rsCtaCteP
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteP, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        Do While Not .EOF = True
            If IsNull(.Fields("debito").Value) Or .Fields("debito").Value = "" Then
                debito = 0
            Else
                debito = .Fields("debito").Value
            End If
                        
            If IsNull(.Fields("credito").Value) Or .Fields("credito").Value = "" Then
                credito = 0
            Else
                credito = .Fields("credito").Value
            End If
                        
            SaldoAnterior = SaldoAnterior + debito - credito
            totalCredito = totalCredito + credito
            totaldebito = totaldebito + debito
            
            .MoveNext
        Loop
        
    End With

    
    If totaldebito = "" Then
        totaldebito = 0
    End If
    
    If totalCredito = "" Then
        totalCredito = 0
    End If
    
    If SaldoAnterior = "" Then
        SaldoAnterior = 0
    End If
    
    total = totaldebito
    pendiente = SaldoAnterior
    pagado = totalCredito
    
    Dim yaCargado As Boolean
    yaCargado = False
    For i = 1 To Me.KlexDetalle.Rows - 1
        If Me.remito = Val(KlexDetalle.TextMatrix(i, 7)) And KlexDetalle.TextMatrix(i, 7) <> "" Then
            MsgBox KlexDetalle.TextMatrix(i, 2) & " " & KlexDetalle.TextMatrix(i, 3) & " ya ha sido cargado", vbInformation, "WGestion"
            yaCargado = True
        End If
    Next i
    
    If Not esComprobanteAutomatico Then
        If Not yaCargado Then
            Me.AgregarDocumentoAPagar Me.total, Me.pendiente, Me.pagado, Me.NroComprobante, Me.tipoComprobante, Me.fechaDocumento, Me.remito
            txtTotal.Text = totaldebito
            txtPendiente.Text = SaldoAnterior
            txtPagado.Text = totalCredito
            
            'Muestro el total pendiente
            Me.txtMontoTotalPendienteSeleccionado = Val(Me.txtMontoTotalPendienteSeleccionado) + Me.pendiente
        End If
    Else
        
        Dim rsPFac As New ADODB.Recordset, sqlFac As String
        txtMontoTotalPendienteSeleccionado = 0
        sqlFac = "SELECT * FROM PFactura WHERE (codigo = '" & codProv & "')"
     
        With rsPFac
            .CursorLocation = adUseClient
               
            Call .Open(sqlFac, ConnDDBB, adOpenStatic, adLockPessimistic)
            Do While Not rsPFac.EOF
                CalcularSaldosPorRemito rsPFac.Fields("remito").Value, codProv
                If Me.pendiente > 0 Then
                    Me.AgregarDocumentoAPagar Me.total, Me.pendiente, Me.pagado, rsPFac("NComprobante"), rsPFac.Fields("Tipo").Value, rsPFac.Fields("Fecha").Value, rsPFac.Fields("remito").Value
                
                    txtTotal.Text = Val(txtTotal.Text) + totaldebito
                    txtPendiente.Text = Val(txtPendiente.Text) + SaldoAnterior
                    txtPagado.Text = Val(txtPagado.Text) + totalCredito
                    
                    'Muestro el total pendiente
                    Me.txtMontoTotalPendienteSeleccionado = Val(txtMontoTotalPendienteSeleccionado.Text) + Val(pendiente)
                    
                End If
                rsPFac.MoveNext
            Loop
        End With
        
        End If
    
        If Val(Me.TxtTotalAPagar) < Val(txtMontoTotalPendienteSeleccionado) Then
            txtMontoTotalPendienteSeleccionado.ForeColor = &HFF&
        Else
            txtMontoTotalPendienteSeleccionado.ForeColor = &H80000008
        End If
        
    Else
        txtMontoTotalPendienteSeleccionado.Text = GenerarDato("SELECT Sum(Debito), Sum(Credito), Sum(Debito)-Sum(Credito) FROM PCuentasCorrientes WHERE Codigo = '" & codProv & "';", "Sum(Debito)-Sum(Credito)")
    End If
    
    SetearDatosProveedor (codProv)
    
If Err Then GrabarLog "BuscarDatosOperacionesProveedor", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CalcularSaldosPorRemito(remito As Long, codProv As String)
    
    SaldoAnterior = 0
    totaldebito = 0
    totalCredito = 0
    credito = 0
    debito = 0
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT * FROM PCuentasCorrientes WHERE (codigo = '" & codProv & "') and (remito = " & remito & ")"
    
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        Do While Not .EOF
            If IsNull(.Fields("debito").Value) Or .Fields("debito").Value = "" Then
                debito = 0
            Else
                debito = .Fields("debito").Value
            End If
                        
            If IsNull(.Fields("credito").Value) Or .Fields("credito").Value = "" Then
                credito = 0
            Else
                credito = .Fields("credito").Value
            End If
                        
            SaldoAnterior = SaldoAnterior + debito - credito
            totalCredito = totalCredito + credito
            totaldebito = totaldebito + debito
            
            .MoveNext
        Loop
        
    End With
    
    If totaldebito = "" Then
        totaldebito = 0
    End If
    
    If totalCredito = "" Then
        totalCredito = 0
    End If
    
    If SaldoAnterior = "" Then
        SaldoAnterior = 0
    End If
    
    total = totaldebito
    pendiente = SaldoAnterior
    pagado = totalCredito
End Sub
Private Sub FormatoGrillaDetalle(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With KlexDetalle
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 8
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 400
        
        .TextMatrix(0, 1) = "Fecha"
        .ColWidth(1) = 1100
               
        .TextMatrix(0, 2) = "Tipo comprobante"
        .ColWidth(2) = 2000
        
        .TextMatrix(0, 3) = "Nro. comprobante"
        .ColWidth(3) = 1500
        
        .TextMatrix(0, 4) = "Total"
        .ColWidth(4) = 1000
        .ColDisplayFormat(4) = "#0.000"
        
        .TextMatrix(0, 5) = "Pendiente"
        .ColWidth(5) = 1000
        .ColDisplayFormat(5) = "#0.000"
        
        .TextMatrix(0, 6) = "Pagado"
        .ColWidth(6) = 1000
        .ColDisplayFormat(6) = "#0.000"
        
        .TextMatrix(0, 7) = "Remito"
        .ColWidth(7) = 0

    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtNroCuponTarjeta_GotFocus()
    Me.Resizer.VScrollPosition = 4905
End Sub
Private Sub txtNroCuponTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.txtNroCuponTarjeta.Text = "" Then
        PusGrabar(0).SetFocus
    Else
        Me.cboBancoTarjeta.SetFocus
    End If
End If
End Sub

Private Sub txtPagado_Change()
    'If txtPagado <> "" Then
    '    Me.txtPagado.Text = pagado
    'End If
End Sub

Private Sub txtPendiente_Change()
    'If Me.txtPendiente <> "" Then
    '    Me.txtPendiente.Text = pendiente
    'End If
End Sub

Private Sub txtTotal_Change()
    'If txtTotal <> "" Then
    '    Me.txtTotal.Text = total
    'End If
    'me.txtPendiente.Text =
End Sub
Private Function SetearDatosProveedor(codProv As String) As Long
    On Error Resume Next
    
    Dim rsProv As New ADODB.Recordset, sqlCtaCteC As String, sqlProv As String
    
    sqlProv = "SELECT * FROM Proveedores WHERE (codigo = '" & codProv & "')"
     
    With rsProv
        .CursorLocation = adUseClient
               
        Call .Open(sqlProv, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .EOF Then
        
            .MoveFirst
            
            txtProveedor(1).Text = EsNulo(.Fields("Nombre").Value)
            txtProveedor(0).Text = EsNulo(.Fields("Codigo").Value)
            
        End If
    End With
    
    sqlProv = ""

    If rsProv.State = 1 Then
        rsProv.Close
        Set rsProv = Nothing
    End If
    
If Err Then GrabarLog "SetearDatosProveedor", Err.Number & " " & Err.Description, Me.Caption
End Function
Public Sub PagarCtaCteAutomaticamente(importeAPagar As Double, idMedioPago As Integer)
    On Error Resume Next
    
    Dim rsCtaCteP As New ADODB.Recordset, sqlCtaCteP As String
    Dim cmd As New ADODB.Command
    
    With cmd
        Set .ActiveConnection = ConnDDBB
        .CommandText = "traer_FacturaProveedorConImporteDeuda"
        .CommandType = adCmdStoredProc
        .Parameters.Append cmd.CreateParameter("codProv", adVarChar, adParamInput, 50, codProveedor)
        .Prepared = True
    End With
        
    Set rsCtaCteP = cmd.Execute
    
    With rsCtaCteP
        If Not .EOF = True Then .MoveFirst
        
        Dim importePagado As Double
        
        Do While Not (.EOF = True) And (importeAPagar >= importePagado)
            'Llamo a pagarCtaCte con ese nro de remito
            If Val(importeAPagar) - Val(importePagado) > Val(.Fields("ImporteDeudaDocumento").Value) Then
                Call PagarCtaCte(.Fields("Remito").Value, .Fields("ImporteDeudaDocumento").Value, idMedioPago)
            Else
                Call PagarCtaCte(Val(.Fields("Remito").Value), Val(importeAPagar) - Val(importePagado), idMedioPago)
            End If
            
            
            If Left(Trim(concepto), Len(.Fields("comentario").Value)) <> Left(Trim(.Fields("comentario").Value), Len(concepto)) Or Len(concepto) = 0 Then
                concepto = EsNulo(.Fields("Comentario").Value) & concepto & "; "
            End If
            
            importePagado = Val(Format(importePagado, "######0.00")) + Val(Format(.Fields("ImporteDeudaDocumento").Value, "######0.00"))
            .MoveNext
        Loop
                
    End With
    
If Err Then GrabarLog "PagarCtaCteAutomaticamente", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub LimpiarCampos()
On Error Resume Next

    
    cboEstadoCheque.Text = ""
    txtFirmanteCheque.Text = ""
    txtImporteCheque.Text = ""
    txtImporteTotalCheque.Text = ""
    txtImporteEfectivoPesos.Text = ""
    txtImporteCuponTarjeta.Text = ""
    txtProveedor(0).Text = ""
    txtProveedor(1).Text = ""
    txtNroComprobante.Text = ""
    txtNroCheque.Text = ""
    txtPagado.Text = ""
    txtPendiente.Text = ""
    txtTipoComp.Text = ""
    txtTotal.Text = ""
    txtNroCuponTarjeta.Text = ""
    txtMontoTotalPendienteSeleccionado.Text = ""
    
    
    concepto = ""
    
    
    codProveedor = ""
    remito = 0
    
    dtpDepositoCheque.Value = Date
    dtpFecha.Value = Date

    VaciarChequesTemp

If Err Then GrabarLog "LimpiarCampos", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function GuardarCuponTarjeta() As Integer
    On Error Resume Next
    
    Dim rs As New ADODB.Recordset
    
    If (remito = 0) And (codProveedor = "") Then
        MsgBox "Debe seleccionar un Proveedor o un remito antes de iniciar la operacion", vbOKOnly, "Mensaje ..."
    Else
        Dim sql As String
    
        With rs
            .CursorLocation = adUseClient
            sql = "SELECT * FROM CuponTarjeta"
            Call .Open(sql, ConnDDBB, adOpenDynamic, adLockPessimistic)
            
            .AddNew
            .Fields("idtarjeta").Value = Me.cboTarjeta.Tag
        
            .Fields("idBanco").Value = Me.cboBancoTarjeta.Tag
            .Fields("Importe").Value = Val(Me.txtImporteCuponTarjeta.Text)
        
            .Fields("CantCuotas").Value = Val(Me.txtCantCuotas.Text)
        
            .Fields("NroCupon").Value = Trim(Me.txtNroCuponTarjeta.Text)
            
            .Update
           
            GuardarCuponTarjeta = .Fields("idCuponTarjeta").Value
                
        End With
    
    End If
    
If Err Then GrabarLog "GuardarCuponTarjeta", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Function
Public Sub HabilitarControles(b As Boolean)
On Error Resume Next

    Dim i As Integer

    
    cboEstadoCheque.Enabled = b
    txtFirmanteCheque.Enabled = b
    txtImporteCheque.Enabled = b
    txtMontoTotalPendienteSeleccionado.Enabled = b
    txtImporteTotalCheque.Enabled = b
    txtImporteEfectivoPesos.Enabled = b
    txtImporteCuponTarjeta.Enabled = b
    txtProveedor(0).Enabled = b
    txtProveedor(1).Enabled = b
    txtNroComprobante.Enabled = b
    txtNroCheque.Enabled = b
    txtPagado.Enabled = b
    txtPendiente.Enabled = b
    txtTipoComp.Enabled = b
    txtTotal.Enabled = b
    txtNroCuponTarjeta.Enabled = b
    
    For i = 0 To Val(txtBancoCheque.Count - 1)
        txtBancoCheque(i).Enabled = b
    Next
    
    cboBancoTarjeta.Enabled = b
    For i = 0 To Val(pbCarga.Count - 1)
        pbCarga(i).Enabled = b
    Next
    
    txtImporteEfectivoDolar.Enabled = b
    dtpDepositoCheque.Enabled = b
    cmdAgregarCheque.Enabled = b
    cmdEliminarCheque.Enabled = b
    cboBancoTarjeta.Enabled = b
    cboTarjeta.Enabled = b
    txtCantCuotas.Enabled = b
    cboTarjeta.Enabled = b
    PusGrabar(0).Enabled = b
    dtpFecha.Enabled = b

If Err Then GrabarLog "HabilitarControles", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function ValidarIngresoCheque() As Boolean
    
    ValidarIngresoCheque = True
    
    If Val(txtImporteCheque.Text) = 0 Or Val(txtImporteCheque.Text) < 0 Then
        MsgBox "El importe del cheque debe ser un valor mayor o igual a cero.", vbInformation, "WGestion"
        ValidarIngresoCheque = False
        Exit Function
    End If
    
    If Val(txtNroCheque.Text) = 0 Then
        MsgBox "El campo Nro. cheque es de ingreso obligatorio", vbInformation, "WGestion"
        ValidarIngresoCheque = False
        Exit Function
    End If
        
End Function
Private Sub AgregarPagoRecibo(idMedioPago As Integer, desc As String, monto As Double)
On Error Resume Next

    Dim sqlRecibo As String

    sqlRecibo = "SELECT * FROM Recibo_Temp"
    
    With rsRecibo
        If .State = 1 Then .Close
        
        Call .Open(sqlRecibo, ConnDDBB, adOpenDynamic, adLockPessimistic)
        
        If .State = 1 Then
            .AddNew
            .Fields("idMedioPago").Value = idMedioPago
            .Fields("Descripcion").Value = Left(desc, 45)
            .Fields("Monto").Value = Val(monto)
            .Fields("Lugar").Value = vDatosEmpresa.Localidad
            .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
            .Fields("Concepto").Value = ""
            .Fields("Total").Value = Val(TxtTotalAPagar.Text)
            .Update
        End If
    
    End With
    
    sqlRecibo = ""

If Err Then GrabarLog "AgregarPagoRecibo", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub VaciarTablaRecibo()
On Error Resume Next

    Call BorrarBase("Recibo_Temp", pathDBMySQL)

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ImprimirRecibo()
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora", vbInformation, "Mensaje ..."
    
    With drRecibo
        .Sections(2).Controls("lbllugar").Caption = vDatosEmpresa.Localidad & ", "
        .Sections(2).Controls("lblfecha").Caption = Date
        .Sections(2).Controls("lblCliente").Caption = txtProveedor(0).Text & "-" & txtProveedor(1).Text
        If Me.esComprobanteAutomatico Then
            .Sections(5).Controls("lblconcepto").Caption = concepto
        Else
            .Sections(5).Controls("lblconcepto").Caption = txtTipoComp.Text & " " & txtNroComprobante.Text
        End If
        .Sections(5).Controls("lbltotal").Caption = Me.TxtTotalAPagar.Text
        .Hide
    End With
  
    Call drRecibo.PrintReport(False, rptRangeAllPages)
    
    LimpiarCampos

If Err Then GrabarLog "ImprimirRecibo", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PagarCtaCte(vnroremito As Long, importe As Double, idMedioPago As Integer) 'Este metodo estaba en Remito
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String, vTipoComprobante As String, vnrocomprobante As String
    Dim SaldoAnterior, debito, credito As Double
        
    sqlCtaCteC = "SELECT * FROM PCuentasCorrientes WHERE (remito = " & vnroremito & ")"
     
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenDynamic, adLockPessimistic)
        
        Do While Not .EOF
            If IsNull(.Fields("debito").Value) Then
                debito = 0
            Else
                debito = .Fields("debito").Value
            End If
            
            If IsNull(.Fields("credito").Value) Then
                credito = 0
            Else
                credito = .Fields("credito").Value
            End If
                        
            SaldoAnterior = Val(SaldoAnterior) + Val(debito) - Val(credito)
            
            .MoveNext
        Loop
        
        vTipoComprobante = TraerDato("PFactura", "Remito = " & vnroremito & "", "Tipo")
        vnrocomprobante = TraerDato("PFactura", "Remito = " & vnroremito & "", "NComprobante")
        
        If .EOF = False Then .MoveLast
        

        .AddNew
        .Fields("remito").Value = Trim(vnroremito)
        .Fields("comentario").Value = "Pago: Nro. " & vTipoComprobante & " " & Trim(vnrocomprobante)
        
        .Fields("Fecha").Value = strfechaMySQL(dtpFecha.Value)
        '.Fields("Fechainput").value = strfechaMySQL(dtpFecha.value)
        
        'Buscar el cliente segun el remito
        .Fields("Codigo").Value = TraerDato("PFactura", "Remito = " & vnroremito & "", "Codigo")
        .Fields("Nombre").Value = TraerDato("PFactura", "Remito = " & vnroremito & "", "Nombre")
       
        '.Fields("anomes").value = Right(.Fields("Fecha").value, 4) & Mid(.Fields("Fecha").value, 4, 2)
    
        .Fields("idMedioPago") = idMedioPago

        If (vTipoComprobante = "Fact A") Or (vTipoComprobante = "Fact B") Or (vTipoComprobante = "Fact C") Then
            
            .Fields("Debito").Value = 0
            .Fields("Credito").Value = importe
            .Fields("Saldo").Value = SaldoAnterior - Val(Format(.Fields("credito").Value, "#######0.00"))
                    
        End If
        
        .Fields("TipoMovimiento").Value = "RC"
        .Update
        
        vIdCtaCteP = Val(.Fields("idPCuentasCorrientes").Value)
        
    End With

If Err Then GrabarLog "PagarCtaCte", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub TxtTotalAPagar_Change()
On Error Resume Next

    If Val(TxtTotalAPagar.Text) > Val(txtMontoTotalPendienteSeleccionado.Text) Then
        TxtTotalAPagar.BackColor = vbRed
    Else
        TxtTotalAPagar.BackColor = vbWhite
    End If
    
    If Not Val(TxtTotalAPagar.Text) = 0 Then vImporteTotalAPagar = Val(TxtTotalAPagar.Text)
        
If Err Then GrabarLog "TxtTotalAPagar_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
