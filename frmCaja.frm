VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmCaja 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caja Diaria"
   ClientHeight    =   8535
   ClientLeft      =   3240
   ClientTop       =   -2595
   ClientWidth     =   13905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   13905
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc bsaldo_caja 
      Height          =   330
      Left            =   7200
      Top             =   6120
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
      Caption         =   "bsaldo_caja"
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
   Begin VB.PictureBox Picture1 
      Height          =   525
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   9945
      TabIndex        =   0
      Top             =   6480
      Width           =   10005
      Begin VB.Label vtotal_saldo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         DataField       =   "total"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   2
         EndProperty
         DataSource      =   "bsaldo_caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   7710
         TabIndex        =   6
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblLabel15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         DataField       =   "SumaDeRetiro"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   2
         EndProperty
         DataSource      =   "bsaldo_caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4680
         TabIndex        =   5
         Top             =   90
         Width           =   855
      End
      Begin VB.Label lblLabel14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label14"
         DataField       =   "SumaDeDeposito"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   2
         EndProperty
         DataSource      =   "bsaldo_caja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1920
         TabIndex        =   4
         Top             =   90
         Width           =   855
      End
      Begin VB.Label lblTotalDepositado 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "> Saldo Actual :"
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
         Height          =   285
         Index           =   1
         Left            =   6240
         TabIndex        =   3
         Top             =   120
         Width           =   1665
      End
      Begin VB.Label lblTotalDepositado 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "> Total Retiros:"
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
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   2
         Top             =   150
         Width           =   1425
      End
      Begin VB.Label lblSaldo 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   ">Total Depositado:"
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
         Height          =   255
         Left            =   210
         TabIndex        =   1
         Top             =   150
         Width           =   1665
      End
   End
   Begin MSAdodcLib.Adodc bcaja 
      Height          =   330
      Left            =   3720
      Top             =   6120
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
      Caption         =   "bcaja"
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
   Begin MSAdodcLib.Adodc bcajatotales 
      Height          =   330
      Left            =   0
      Top             =   6720
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
      Caption         =   "bcajatotales"
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
   Begin TabDlg.SSTab TabGeneral 
      Height          =   6465
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10065
      _ExtentX        =   17754
      _ExtentY        =   11404
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ing. Movimiento"
      TabPicture(0)   =   "frmCaja.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "saldo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "bProveedor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "bCliente"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCrearConcepto"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdVerResumen"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraFechaImporte"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdIngresarMovimiento(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdLimpiar"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fraMovi"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Ficha Cliente"
      TabPicture(1)   =   "frmCaja.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdBuscar"
      Tab(1).Control(1)=   "cmdBorrar"
      Tab(1).Control(2)=   "cmdImprimir(0)"
      Tab(1).Control(3)=   "cmdImprimir(1)"
      Tab(1).Control(4)=   "Picture2"
      Tab(1).Control(5)=   "DgCaja"
      Tab(1).Control(6)=   "saldo2"
      Tab(1).Control(7)=   "Label9"
      Tab(1).Control(8)=   "vsaldo1"
      Tab(1).ControlCount=   9
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   -74880
         Style           =   1  'Graphical
         TabIndex        =   66
         ToolTipText     =   "Ejecutar consulta"
         Top             =   5670
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.Frame fraMovi 
         Height          =   4125
         Left            =   150
         TabIndex        =   28
         Top             =   1410
         Width           =   9825
         Begin VB.ComboBox cboUsuario 
            Height          =   315
            Left            =   1200
            TabIndex        =   69
            Top             =   3210
            Width           =   7245
         End
         Begin VB.CheckBox chkPagaConCheque 
            Caption         =   "Se recibe cheque"
            Height          =   255
            Left            =   7680
            TabIndex        =   46
            Top             =   2760
            Width           =   1695
         End
         Begin VB.Frame FraCheque 
            Caption         =   "Cheque"
            Height          =   1815
            Left            =   6720
            TabIndex        =   39
            Top             =   840
            Width           =   3495
            Begin VB.TextBox txtNroCheque 
               Height          =   315
               Left            =   720
               TabIndex        =   40
               Top             =   360
               Width           =   2565
            End
            Begin MSComCtl2.DTPicker dtFechaDeposito 
               Height          =   315
               Left            =   1920
               TabIndex        =   41
               Top             =   840
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   57540609
               CurrentDate     =   40122
            End
            Begin MSComCtl2.DTPicker dtFechaConfeccion 
               Height          =   315
               Left            =   1920
               TabIndex        =   42
               Top             =   1320
               Width           =   1365
               _ExtentX        =   2408
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   57540609
               CurrentDate     =   40122
            End
            Begin VB.Label lblNroCheque 
               Caption         =   "> Nro. :"
               Height          =   255
               Left            =   120
               TabIndex        =   45
               Top             =   360
               Width           =   645
            End
            Begin VB.Label lblFechaDeposito 
               Caption         =   "> Fecha deposito:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   44
               Top             =   840
               Width           =   1635
            End
            Begin VB.Label lblFechaConfeccion 
               Caption         =   "> Fecha confeccion:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   43
               Top             =   1320
               Width           =   1995
            End
         End
         Begin VB.TextBox vcomentario 
            Height          =   345
            Left            =   1200
            TabIndex        =   31
            Top             =   3600
            Width           =   7245
         End
         Begin VB.CommandButton cmdBuscarConcepto 
            Caption         =   "Buscar"
            Height          =   495
            Left            =   7680
            Picture         =   "frmCaja.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   30
            ToolTipText     =   "Ejecutar búsqueda"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.TextBox txtConcepto 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   29
            Top             =   360
            Width           =   6375
         End
         Begin TabDlg.SSTab TabProveedor 
            Height          =   2265
            Left            =   120
            TabIndex        =   47
            Top             =   720
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   3995
            _Version        =   393216
            Tabs            =   1
            TabsPerRow      =   1
            TabHeight       =   520
            TabCaption(0)   =   "Tab 0"
            TabPicture(0)   =   "frmCaja.frx":013A
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblProveedor(4)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lblProveedor(1)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lblProveedor(3)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lblProveedor(2)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lblProveedor(5)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lblProveedor(0)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "cboTipoIva(0)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "txtProveedor(5)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "txtProveedor(4)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txtProveedor(3)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txtProveedor(2)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pbCarga(0)"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "txtProveedor(1)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "txtProveedor(0)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).ControlCount=   14
            Begin XtremeSuiteControls.FlatEdit txtProveedor 
               Height          =   315
               Index           =   0
               Left            =   1440
               TabIndex        =   70
               Top             =   240
               Width           =   855
               _Version        =   851968
               _ExtentX        =   1508
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtProveedor 
               Height          =   315
               Index           =   1
               Left            =   2835
               TabIndex        =   71
               Top             =   240
               Width           =   3405
               _Version        =   851968
               _ExtentX        =   6015
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton pbCarga 
               Height          =   315
               Index           =   0
               Left            =   2400
               TabIndex        =   72
               Tag             =   "Proveedor"
               Top             =   240
               Width           =   315
               _Version        =   851968
               _ExtentX        =   556
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit txtProveedor 
               Height          =   315
               Index           =   2
               Left            =   1440
               TabIndex        =   73
               Top             =   600
               Width           =   4815
               _Version        =   851968
               _ExtentX        =   8493
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtProveedor 
               Height          =   315
               Index           =   3
               Left            =   1440
               TabIndex        =   74
               Top             =   960
               Width           =   4815
               _Version        =   851968
               _ExtentX        =   8493
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtProveedor 
               Height          =   315
               Index           =   4
               Left            =   1440
               TabIndex        =   75
               Top             =   1320
               Width           =   4815
               _Version        =   851968
               _ExtentX        =   8493
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtProveedor 
               Height          =   315
               Index           =   5
               Left            =   4800
               TabIndex        =   76
               Top             =   1680
               Width           =   1455
               _Version        =   851968
               _ExtentX        =   2566
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.ComboBox cboTipoIva 
               Height          =   315
               Index           =   0
               Left            =   1440
               TabIndex        =   77
               Top             =   1680
               Width           =   2415
               _Version        =   851968
               _ExtentX        =   4260
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin VB.Label lblProveedor 
               Caption         =   "Proveedor :"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   53
               Top             =   280
               Width           =   1250
            End
            Begin VB.Label lblProveedor 
               Caption         =   "C.U.I.T :"
               Height          =   195
               Index           =   5
               Left            =   3960
               TabIndex        =   52
               Top             =   1720
               Width           =   600
            End
            Begin VB.Label lblProveedor 
               Caption         =   " Localidad :"
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   51
               Top             =   1000
               Width           =   1250
            End
            Begin VB.Label lblProveedor 
               Caption         =   "Teléfono :"
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   50
               Top             =   1360
               Width           =   1250
            End
            Begin VB.Label lblProveedor 
               Caption         =   "Dirección :"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   49
               Top             =   640
               Width           =   1250
            End
            Begin VB.Label lblProveedor 
               Caption         =   "Tipo de I.V.A. :"
               Height          =   195
               Index           =   4
               Left            =   90
               TabIndex        =   48
               Top             =   1720
               Width           =   1250
            End
         End
         Begin TabDlg.SSTab TabCliente 
            Height          =   2265
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   6495
            _ExtentX        =   11456
            _ExtentY        =   3995
            _Version        =   393216
            Tabs            =   1
            TabsPerRow      =   1
            TabHeight       =   520
            TabCaption(0)   =   "Tab 0"
            TabPicture(0)   =   "frmCaja.frx":0156
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "lblCU(1)"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "lblLocalidad(5)"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "lblTeléfono(4)"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "lblDirección(3)"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).Control(4)=   "lblTipoDe(2)"
            Tab(0).Control(4).Enabled=   0   'False
            Tab(0).Control(5)=   "lblCliente(1)"
            Tab(0).Control(5).Enabled=   0   'False
            Tab(0).Control(6)=   "txtCliente(5)"
            Tab(0).Control(6).Enabled=   0   'False
            Tab(0).Control(7)=   "cboTipoIva(1)"
            Tab(0).Control(7).Enabled=   0   'False
            Tab(0).Control(8)=   "txtCliente(4)"
            Tab(0).Control(8).Enabled=   0   'False
            Tab(0).Control(9)=   "txtCliente(3)"
            Tab(0).Control(9).Enabled=   0   'False
            Tab(0).Control(10)=   "txtCliente(2)"
            Tab(0).Control(10).Enabled=   0   'False
            Tab(0).Control(11)=   "pbCarga(1)"
            Tab(0).Control(11).Enabled=   0   'False
            Tab(0).Control(12)=   "txtCliente(1)"
            Tab(0).Control(12).Enabled=   0   'False
            Tab(0).Control(13)=   "txtCliente(0)"
            Tab(0).Control(13).Enabled=   0   'False
            Tab(0).ControlCount=   14
            Begin XtremeSuiteControls.FlatEdit txtCliente 
               Height          =   315
               Index           =   0
               Left            =   1440
               TabIndex        =   78
               Top             =   240
               Width           =   855
               _Version        =   851968
               _ExtentX        =   1508
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtCliente 
               Height          =   315
               Index           =   1
               Left            =   2835
               TabIndex        =   79
               Top             =   240
               Width           =   3405
               _Version        =   851968
               _ExtentX        =   6006
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.PushButton pbCarga 
               Height          =   315
               Index           =   1
               Left            =   2400
               TabIndex        =   80
               Tag             =   "CodigoCliente"
               Top             =   240
               Width           =   315
               _Version        =   851968
               _ExtentX        =   556
               _ExtentY        =   556
               _StockProps     =   79
               Caption         =   "..."
               UseVisualStyle  =   -1  'True
            End
            Begin XtremeSuiteControls.FlatEdit txtCliente 
               Height          =   315
               Index           =   2
               Left            =   1440
               TabIndex        =   81
               Top             =   600
               Width           =   4815
               _Version        =   851968
               _ExtentX        =   8493
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtCliente 
               Height          =   315
               Index           =   3
               Left            =   1440
               TabIndex        =   82
               Top             =   960
               Width           =   4815
               _Version        =   851968
               _ExtentX        =   8493
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtCliente 
               Height          =   315
               Index           =   4
               Left            =   1440
               TabIndex        =   83
               Top             =   1320
               Width           =   4815
               _Version        =   851968
               _ExtentX        =   8493
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.ComboBox cboTipoIva 
               Height          =   315
               Index           =   1
               Left            =   1440
               TabIndex        =   84
               Top             =   1680
               Width           =   2415
               _Version        =   851968
               _ExtentX        =   4260
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin XtremeSuiteControls.FlatEdit txtCliente 
               Height          =   315
               Index           =   5
               Left            =   4800
               TabIndex        =   85
               Top             =   1680
               Width           =   1455
               _Version        =   851968
               _ExtentX        =   2566
               _ExtentY        =   556
               _StockProps     =   77
               BackColor       =   -2147483643
            End
            Begin VB.Label lblCliente 
               Caption         =   "Cliente: "
               ForeColor       =   &H00000000&
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   38
               Top             =   280
               Width           =   1245
            End
            Begin VB.Label lblTipoDe 
               Caption         =   "Tipo de IVA :"
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   37
               Top             =   1720
               Width           =   1250
            End
            Begin VB.Label lblDirección 
               Caption         =   "Dirección :"
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   36
               Top             =   640
               Width           =   1250
            End
            Begin VB.Label lblTeléfono 
               Caption         =   "Teléfono :"
               Height          =   195
               Index           =   4
               Left            =   90
               TabIndex        =   35
               Top             =   1360
               Width           =   1250
            End
            Begin VB.Label lblLocalidad 
               Caption         =   "Localidad :"
               Height          =   195
               Index           =   5
               Left            =   90
               TabIndex        =   34
               Top             =   1000
               Width           =   1250
            End
            Begin VB.Label lblCU 
               Caption         =   "> C.U.I.T :"
               Height          =   195
               Index           =   1
               Left            =   3960
               TabIndex        =   33
               Top             =   1720
               Width           =   855
            End
         End
         Begin VB.Label Label2 
            Caption         =   "> Usuario : "
            Height          =   225
            Left            =   30
            TabIndex        =   56
            Top             =   3240
            Width           =   1250
         End
         Begin VB.Label Label1 
            Caption         =   "> Concepto : "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label4 
            Caption         =   "> Comentario :"
            Height          =   255
            Left            =   30
            TabIndex        =   54
            Top             =   3600
            Width           =   1250
         End
      End
      Begin VB.CommandButton cmdBorrar 
         Caption         =   "Borrar"
         Height          =   375
         Left            =   -73920
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Borrar"
         Top             =   5670
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir Listado General"
         Height          =   285
         Index           =   0
         Left            =   -64290
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Imprimir"
         Top             =   3570
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir Gastos por concepto"
         Height          =   315
         Index           =   1
         Left            =   -64290
         TabIndex        =   25
         Top             =   3870
         Visible         =   0   'False
         Width           =   2235
      End
      Begin VB.CommandButton cmdLimpiar 
         Height          =   315
         Left            =   180
         Picture         =   "frmCaja.frx":0172
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Nuevo movimiento"
         Top             =   390
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000004&
         Height          =   945
         Left            =   -74880
         ScaleHeight     =   885
         ScaleWidth      =   9765
         TabIndex        =   17
         Top             =   120
         Width           =   9825
         Begin VB.TextBox txtPersona 
            Height          =   285
            Left            =   1290
            TabIndex        =   67
            Top             =   480
            Width           =   3975
         End
         Begin VB.TextBox txtConceptoBuscarCaja 
            Height          =   285
            Left            =   1290
            TabIndex        =   65
            Top             =   120
            Width           =   3975
         End
         Begin VB.CheckBox chkFechas 
            BackColor       =   &H80000004&
            Caption         =   "Anular fechas"
            Height          =   195
            Left            =   7920
            TabIndex        =   18
            Top             =   120
            Value           =   1  'Checked
            Width           =   1305
         End
         Begin MSComCtl2.DTPicker fdesde 
            Height          =   285
            Left            =   6510
            TabIndex        =   19
            Top             =   90
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   57540609
            CurrentDate     =   38028
         End
         Begin MSComCtl2.DTPicker fhasta 
            Height          =   285
            Left            =   6510
            TabIndex        =   20
            Top             =   360
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   57540609
            CurrentDate     =   38028
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "> Persona: "
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
            Height          =   195
            Left            =   120
            TabIndex        =   68
            Top             =   510
            Width           =   990
         End
         Begin VB.Label lblHasta 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "> Hasta :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   5670
            TabIndex        =   23
            Top             =   390
            Width           =   795
         End
         Begin VB.Label lblDesde 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Caption         =   "> Desde :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   5640
            TabIndex        =   22
            Top             =   90
            Width           =   885
         End
         Begin VB.Label lblConcepto 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "> Concepto: "
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
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   150
            Width           =   1110
         End
      End
      Begin VB.CommandButton cmdIngresarMovimiento 
         Appearance      =   0  'Flat
         Caption         =   "Ingresar movimiento"
         Height          =   495
         Index           =   1
         Left            =   120
         Picture         =   "frmCaja.frx":0274
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   5640
         UseMaskColor    =   -1  'True
         Width           =   1545
      End
      Begin VB.Frame fraFechaImporte 
         Height          =   705
         Left            =   150
         TabIndex        =   11
         Top             =   720
         Width           =   5835
         Begin VB.TextBox vimporte 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1034
               SubFormatType   =   1
            EndProperty
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
            Height          =   315
            Left            =   4320
            TabIndex        =   12
            Top             =   270
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker vfecha 
            Height          =   315
            Left            =   1080
            TabIndex        =   13
            Top             =   270
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   57540609
            CurrentDate     =   38028
         End
         Begin VB.Label lblFecha 
            Caption         =   "> Fecha :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   300
            Width           =   915
         End
         Begin VB.Label lblImporte 
            Caption         =   "> Importe :"
            ForeColor       =   &H00000080&
            Height          =   225
            Left            =   3480
            TabIndex        =   14
            Top             =   300
            Width           =   825
         End
      End
      Begin VB.CommandButton cmdVerResumen 
         Caption         =   "Ver detalle"
         Height          =   495
         Left            =   1680
         Picture         =   "frmCaja.frx":0DAE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   5640
         Width           =   1575
      End
      Begin VB.CommandButton cmdCrearConcepto 
         Caption         =   "ABM Conceptos"
         Height          =   495
         Left            =   3240
         Picture         =   "frmCaja.frx":18A8
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   5640
         Width           =   1575
      End
      Begin MSDataGridLib.DataGrid DgCaja 
         Bindings        =   "frmCaja.frx":71C2
         Height          =   4185
         Left            =   -74880
         TabIndex        =   57
         Top             =   1200
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   7382
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSAdodcLib.Adodc bCliente 
         Height          =   330
         Left            =   3840
         Top             =   6600
         Visible         =   0   'False
         Width           =   2505
         _ExtentX        =   4392
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
         Caption         =   "bcliente"
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
      Begin MSAdodcLib.Adodc bProveedor 
         Height          =   330
         Left            =   4920
         Top             =   6960
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
         Caption         =   "bProveedor"
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
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Saldo Caja :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6750
         TabIndex        =   64
         Top             =   5750
         Width           =   1065
      End
      Begin VB.Label saldo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   7800
         TabIndex        =   63
         Top             =   5700
         Width           =   1605
      End
      Begin VB.Label saldo2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   -67140
         TabIndex        =   62
         Top             =   5700
         Width           =   1725
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         Caption         =   "Saldo Caja :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   -68250
         TabIndex        =   61
         Top             =   5750
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         Caption         =   "Mantenimiento de Caja Diaria"
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
         Height          =   315
         Left            =   0
         TabIndex        =   60
         Top             =   0
         Width           =   10005
      End
      Begin VB.Label vsaldo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   -66960
         TabIndex        =   59
         Top             =   4260
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Crear un nuevo movimiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   570
         TabIndex        =   58
         Top             =   450
         Width           =   2385
      End
   End
   Begin VB.Label Label6 
      Caption         =   "> Comentario :"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1125
   End
End
Attribute VB_Name = "frmCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vModifica As Boolean
Dim vidd As Long
Dim sql, vfilter, vclave As String
Dim vsaldo As Double
Public esIngreso As Boolean
'Public vCodigoCliente, vCodigoProveedor As String ' codigo del cliente
Public vCodigoConcepto As Long
Private Sub calsaldo(vtotal_saldo)
Dim saldoacreditado As Double
    
    On Error Resume Next

    saldoacreditado = 0
    vtotal_saldo = 0
    
    With bcaja
        .RecordSource = "SELECT * FROM caja WHERE 1=1"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
    End With
    
    If Err Then GrabarLog "calsaldo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarGrilla()
    On Error Resume Next
    
    With bcaja
        .RecordSource = "SELECT * FROM caja WHERE 1=1"
        .Refresh
    End With
    
    If Err Then GrabarLog "CargarGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cboTipoIva_Click(Index As Integer)
On Error Resume Next

    cboTipoIva(Index).Tag = TraerDato("TipoIva", "TipoIva = '" & Trim(cboTipoIva(Index).Text) & "'", "idTipoIva")
    
If Err Then GrabarLog "cboTipoIva_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboTipoIva_GotFocus(Index As Integer)
On Error Resume Next
    
    Call CargarComboNew("TipoIva", "TipoIva", cboTipoIva(Index), True)

If Err Then GrabarLog "cboTipoIva_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboUsuario_Click()
    On Error Resume Next

    If Not (vConfigGral.vUser = cboUsuario.Text) Then
    
        cboUsuario.Tag = Trim(TraerDato("usuarios", "Usuario = '" & (cboUsuario.Text) & "'", "Password"))
        cboUsuario.Tag = DesEncriptar(cboUsuario.Tag, LeerConfig(0))
    
    End If
    
    If Err Then GrabarLog "cboUsuario_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboUsuario_GotFocus()
On Error Resume Next

    'Call CargarCombo(pathDBMySQL, "Usuarios", "Usuario", cboUsuario, True)

If Err Then GrabarLog "cboUsuario_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub cmdCrearConcepto_Click()
On Error Resume Next

    MousePointer = vbHourglass
    
    With frmEstructuraCaja
        .mostrarSaldos = False
        .leido = False
        .vModo = Modo.Creacion
        .Show
    End With
    
    MousePointer = Default

If Err Then GrabarLog "cmdBuscarProveedor_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdVerResumen_Click()
On Error Resume Next
    
    MousePointer = vbHourglass
    
    With frmEstructuraCaja
        .mostrarSaldos = True
        .leido = False
        .vModo = Modo.Lectura
        .Show
    End With
    
    MousePointer = Default

If Err Then GrabarLog "cmdVerResumen_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub chkFechas_Click()
    On Error Resume Next

    fdesde.Enabled = CBool(chkFechas.Value - 1)
    fhasta.Enabled = CBool(chkFechas.Value - 1)

    If Err Then GrabarLog "chkFechas_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub chkPagaConCheque_Click()
On Error Resume Next
    
    FraCheque.Visible = CBool(chkPagaConCheque.Value)

If Err Then GrabarLog "chkPagaConCheque_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdBuscarConcepto_Click()
On Error Resume Next

    vVieneConcepto = Me.Name

    With frmEstructuraCaja
        .leido = False
        .vModo = Modo.Seleccion
        .Show
    End With
    
If Err Then GrabarLog "cmdBuscarConcepto_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdIngresarMovimiento_Click(Index As Integer)
Dim valido As Boolean
On Error Resume Next
    
        If (ValidarCampos) = True Then
            valido = True
            With bcaja
                .ConnectionString = pathDBMySQL
                .RecordSource = "SELECT * FROM caja"
                .Refresh
          
            
                If vModifica = False Then .Recordset.AddNew
            
                .Recordset("Fecha").Value = strfechaMySQL(vfecha.Value)
                .Recordset("Importe").Value = Val(Me.vimporte.Text)
                
                If esIngreso = True Then
                    .Recordset("CodigoCliente").Value = EsNulo(txtCliente(0).Text)
                Else
                    .Recordset("CodigoProveedor").Value = EsNulo(txtProveedor(0).Text)
                End If
                
                .Recordset("usuario").Value = Trim(cboUsuario.Text)
                .Recordset("CodigoConcepto").Value = vCodigoConcepto
                .Recordset("comentario").Value = vcomentario.Text
                    
                .Recordset("NroCheque") = txtNroCheque.Text
                .Recordset("FechaDeposito") = dtFechaDeposito.Value
                .Recordset("FechaConfeccion") = dtFechaConfeccion.Value
                
                .Recordset.Update
                
                .Refresh
        
                Limpiar
                
                If vConfigGral.vIncluyeContabilidad = True Then
                    With frmAsientosAlta
                        .Show
                        .ZOrder (0)
                        .txtCuentaVieneDe.Text = Me.Caption
                    End With
                End If
           
           End With
    
    End If

    If Err Then
        GrabarLog "cmdAccion_Click", Err.Number & " " & Err.Description, Me.Name
    Else
        If valido Then
            MsgBox "El movimiento se ha registrado satisfactoriamente", vbInformation
            Unload Me
            Load Me
            Activate
        End If
        
        'Esto pq no actualiza rapido el recordset la caja
        Me.saldo.Visible = False
        Wait 3000
        CalcularSaldo
        Me.saldo.Visible = True
    End If

If Err Then GrabarLog "cmdIngresarMovimiento_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function ValidarCampos() As Boolean
On Error Resume Next

    ValidarCampos = True
    
    If Val(Me.vimporte.Text) = 0 Then
        MsgBox "El importe debe ser un valor mayor o igual que cero", vbInformation
        ValidarCampos = False
        Exit Function
    End If
    
    If Not Len(Me.txtConcepto.Text) > 0 Then
        MsgBox "El campo 'Concepto' es obligatorio", vbInformation
        ValidarCampos = False
        Exit Function
    End If
    
    If Me.esIngreso = True Then
        If Trim(Me.txtCliente(0).Text) = "" Then
            MsgBox "El campo 'Cliente' es obligatorio", vbInformation
        ValidarCampos = False
        Exit Function
        End If
    Else
        If Trim(Me.txtProveedor(0).Text) = "" Then
            MsgBox "El campo 'Proveedor' es obligatorio", vbInformation
        ValidarCampos = False
        Exit Function
        End If
    End If
    
    If Len(cboUsuario.Text) = 0 Then
        MsgBox "El campo 'Usuario' es obligatorio", vbInformation
        ValidarCampos = False
        Exit Function
    End If
    
If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub cmdBorrar_Click()
    On Error Resume Next

    With bcaja
        
        If Not (.Recordset.EOF = True) And Not (.Recordset.BOF = True) Then
             
            If MsgBox("¿ Esta seguro que desea borrar este registro ? ", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
            
                Dim cmdCaja As New ADODB.Command
                Dim sqlCaja As String
                cmdCaja.ActiveConnection = ConnDDBB
                  
                sqlCaja = "SELECT * FROM Caja WHERE (id = " & .Recordset.Fields("id").Value & ")"
                
                Dim rsCaja As New ADODB.Recordset
                  
                If rsCaja.State = 0 Then
                    rsCaja.Open sqlCaja, ConnDDBB, 3, 3
                Else
                    Set rsCaja = ConnDDBB.Execute(sqlCaja)
                End If
                  
                If Not rsCaja.EOF Then
                    rsCaja.MoveFirst
                    rsCaja.Delete
                    rsCaja.Update
                    rsCaja.Requery
                End If
                Unload Mantenimiento
                Load Mantenimiento
                .Refresh
                Buscar
            End If
       
        
        End If
    
    End With

    CalcularSaldo
    
    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    On Error Resume Next

    If Index = 0 Then
                
        With bcajatotales
            If (.Recordset.EOF = True) Then .Recordset.AddNew
            .Recordset("saldo").Value = Val(saldo.Caption)
            .Recordset("id").Value = vclave
            .Recordset.Update
        End With

        Unload Mantenimiento
        Load Mantenimiento
    
        MsgBox "    Prepare la Impresora    ", vbInformation, "Mensaje ..."
    
        With Mantenimiento.rsccaja2
        
            If Not .State = 1 Then
                .Open
                .Close
                .Open
            Else
                .Close
                .Open
            End If

            .Filter = "vtotal >= -10000000000 " & vfilter
    
        End With
    
        With drcaja
            .Sections("TituloEmpresa").Controls("lblSaldoAnterior").Caption = CalcularSaldoAnterior(fdesde.Value)
            .Sections("section5").Controls("vsaldo").Caption = vtotal_saldo.Caption
            .Show
        End With
    
    Else
    
        'frmGastosConcepto.Show
    
    End If

    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdBuscar_Click()
    Buscar
 End Sub
Private Sub Buscar()
On Error Resume Next

    Dim sqlCli, sqlPro, sqlCliProv As String
    
    sql = ""
    vfilter = ""

    
    sqlCli = "cli.Nombre like '%" & txtPersona.Text & "%'"
    
    sqlPro = "p.Nombre like '%" & txtPersona.Text & "%'"
        
    sqlCliProv = " and (" & sqlCli & " or " & sqlPro & ")"
    
    
    If Not Trim(Me.txtConceptoBuscarCaja) = "" Then
        sql = sql + " and co.descripcion like '%" + Trim(Me.txtConceptoBuscarCaja.Text) + "%'"
        vfilter = vfilter + " and co.descripcion like '*" + Trim(Me.txtConceptoBuscarCaja.Text) + "*'"
    End If
        
    If chkFechas.Value = 0 Then
        sql = sql + " and (fecha >= '" & strfechaMySQL(fdesde.Value) + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "')"
        vfilter = vfilter + " and fecha >= '" & strfechaMySQL(fdesde.Value) + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "' "
    End If

    With bcaja
        .RecordSource = "SELECT ca.Fecha, cli.Nombre & p.Nombre as Persona, ca.Importe, co.Descripcion as Concepto, IF(co.IngresoEgreso = True,  'Ingreso', 'Egreso') AS Tipo, ca.id " & _
        " FROM (((caja ca inner join Concepto co on ca.CodigoConcepto = co.Codigo) left join Proveedores p on ca.CodigoProveedor = p.Codigo) " & _
                 " left join Clientes cli on ca.CodigoCliente = cli.Codigo)" & _
        " WHERE 1=1 " + sql + sqlCliProv + " ORDER BY fecha"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst

        vclave = Year(Date) & Hour(Time) & Second(Time)

    End With
    
    If Err Then GrabarLog "cmdBuscar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdLimpiar_Click()
    On Error Resume Next
    
    Limpiar
    
    If Err Then GrabarLog "cmdLimpiar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Activate()
    Activate
    CargarGrilla
    Buscar
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Activate()
    Me.TabCliente.Visible = Me.esIngreso
    Me.chkPagaConCheque.Visible = Me.esIngreso
    Me.TabProveedor.Visible = Not Me.esIngreso
    Me.FraCheque.Visible = False
    TabProveedor.Caption = ""
End Sub
Private Sub Form_Load()
    On Error Resume Next
    
    With bcaja
        .ConnectionString = pathDBMySQL
        .RecordSource = "caja"
        .Refresh
    End With

    With bcajatotales
        .ConnectionString = pathDBMySQL
        .RecordSource = "Cajatotales"
        .Refresh
    End With

    sql = ""
    fdesde.Value = Date
    fhasta.Value = Date
    vfecha.Value = Date
    vModifica = False
    
    cboUsuario.Text = vConfigGral.vUser
    
    TabGeneral.tab = 0
    
    With Me
        .Top = 700
        .Left = 700
        .Width = 10200
        .Height = 6945
    
    End With
    
    CalcularSaldo
    
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()
    On Error Resume Next
    
    vfecha.Value = Date
    vimporte.Text = ""
    Me.txtConcepto.Text = ""
    cboUsuario.Text = ""
    vcomentario.Text = ""
    
    vModifica = False
    vimporte.SetFocus
    
    Dim i As Integer

    For i = 0 To txtCliente.Count - 1
        txtCliente(i).Text = ""
    Next

    txtConcepto.Text = ""
    txtNroCheque.Text = ""
    dtFechaDeposito.Value = Date
    dtFechaConfeccion.Value = Date
    saldo.Caption = ""
    
    If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub DgCaja_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    
    Call OrdenarDataGrid(ColIndex, bcaja.Recordset, DgCaja)
    
    If Err Then GrabarLog "DgCaja_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Function CalcularSaldoAnterior(vfdesde As Date) As Double
    On Error Resume Next

    Dim rsSaldoAnterior As New ADODB.Recordset
    Dim sqlSaldoAnterior As String
    
    If chkFechas.Value = 0 Then
        sqlSaldoAnterior = "SELECT Sum(Caja.Deposito) AS SumaDeDeposito, Sum(Caja.Retiro) AS SumaDeRetiro, [SumaDeDeposito]-[SumadeRetiro] AS Saldo FROM Caja WHERE (Caja.Fecha < '" & Str(vfdesde) & "')"
    Else
        sqlSaldoAnterior = "SELECT Sum(Caja.Deposito) AS SumaDeDeposito, Sum(Caja.Retiro) AS SumaDeRetiro, [SumaDeDeposito]-[SumadeRetiro] AS Saldo FROM Caja WHERE (((Caja.Fecha)<#01/01/2000#))"
    End If
    
    With rsSaldoAnterior
        Call .Open(sqlSaldoAnterior, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF Then
            .MoveFirst
            CalcularSaldoAnterior = .Fields("Saldo").Value
        Else
            CalcularSaldoAnterior = 0
        End If
    
    End With

    sqlSaldoAnterior = ""
    
    rsSaldoAnterior.Close
    Set rsSaldoAnterior = Nothing
    
    If Err Then GrabarLog "CalcularSaldoAnterior", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub pbCarga_Click(Index As Integer)
On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
        
        Case 0 To 10
            frmBusqueda.Show
            
    End Select

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub saldo_Change()
    On Error Resume Next
    
    saldo.Caption = Format(CDbl(saldo.Caption), "#####0.00")
    saldo2.Caption = Format(CDbl(saldo.Caption), "#####0.00")

    If Err Then GrabarLog "limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CalcularSaldo()
On Error Resume Next

    Dim cmd As New ADODB.Command, sql As String
    
    cmd.ActiveConnection = ConnDDBB
    
    Dim sqlCliProv  As String
    Dim sqlFecha As String
    Dim sqlUsu As String
    Dim sqlCon As String
        
    sql = "SELECT fecha, p.Nombre & cli.Nombre AS Nombre, usuario, Descripcion, IF(IngresoEgreso = True,  Importe, 0) AS Ingreso, IF(IngresoEgreso = False,  Importe, 0) AS Egreso, IF(IngresoEgreso = True,  Importe, - Importe) AS Saldo FROM ((caja AS c LEFT JOIN clientes AS cli ON c.CodigoCliente = cli.Codigo) LEFT JOIN Proveedores AS p ON c.CodigoProveedor = p.Codigo) LEFT JOIN Concepto AS co ON co.Codigo = c.CodigoConcepto" _
     & " where c.NroCheque = ''"
     
    Dim rsConsultaCaja As New ADODB.Recordset
      
    If rsConsultaCaja.State = 0 Then
        rsConsultaCaja.Open sql, ConnDDBB, 3, 3
    Else
        Set rsConsultaCaja = ConnDDBB.Execute(sql)
    End If
    
    Dim saldo As Double
    Do While Not rsConsultaCaja.EOF = True
        
        saldo = saldo + Val(Format(rsConsultaCaja("Saldo").Value, "#######0.00"))
       
        rsConsultaCaja.MoveNext
    Loop
        
    Me.saldo.Caption = saldo
       
If Err Then GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
End Sub
