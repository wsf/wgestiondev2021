VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Monito 
   Caption         =   "Monitor del sistema - Control de errores y warnning"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7995
   ScaleWidth      =   13710
   WindowState     =   2  'Maximized
   Begin TabDlg.SSTab SSTab1 
      Height          =   10755
      Left            =   -30
      TabIndex        =   0
      Top             =   30
      Width           =   15225
      _ExtentX        =   26855
      _ExtentY        =   18971
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Arreglos de Valores"
      TabPicture(0)   =   "Monito.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "DataGrid4"
      Tab(0).Control(1)=   "bctacte_dif_true_false"
      Tab(0).Control(2)=   "bctacte"
      Tab(0).Control(3)=   "dbar"
      Tab(0).Control(4)=   "Command3"
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(6)=   "Frame3"
      Tab(0).Control(7)=   "Frame2"
      Tab(0).Control(8)=   "Frame1"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Datos Borrados"
      TabPicture(1)   =   "Monito.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame8"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Tab 2"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.Frame Frame8 
         Caption         =   "Detalles de facturas Borradas:"
         ForeColor       =   &H00000080&
         Height          =   2235
         Left            =   9480
         TabIndex        =   34
         Top             =   600
         Width           =   5145
         Begin VB.CommandButton Command10 
            Caption         =   "Command10"
            Height          =   255
            Left            =   330
            TabIndex        =   40
            Top             =   1920
            Width           =   1545
         End
         Begin VB.PictureBox Picture6 
            Height          =   585
            Left            =   120
            ScaleHeight     =   525
            ScaleWidth      =   4815
            TabIndex        =   36
            Top             =   300
            Width           =   4875
            Begin VB.CommandButton Command11 
               Caption         =   "Command11"
               Height          =   435
               Left            =   2100
               TabIndex        =   41
               Top             =   90
               Width           =   1035
            End
            Begin VB.CommandButton Command9 
               Caption         =   "inpu credito"
               Height          =   405
               Left            =   1020
               TabIndex        =   39
               Top             =   60
               Width           =   915
            End
            Begin VB.CommandButton Command8 
               Caption         =   "inpu debito"
               Height          =   405
               Left            =   60
               TabIndex        =   37
               Top             =   60
               Width           =   915
            End
            Begin ComctlLib.ProgressBar ProgressBar1 
               Height          =   345
               Left            =   1650
               TabIndex        =   38
               Top             =   90
               Width           =   3105
               _ExtentX        =   5477
               _ExtentY        =   609
               _Version        =   327682
               Appearance      =   1
            End
         End
         Begin VB.ListBox l7 
            Height          =   1035
            Left            =   150
            TabIndex        =   35
            Top             =   930
            Width           =   4875
         End
         Begin MSAdodcLib.Adodc Adodc8 
            Height          =   345
            Left            =   750
            Top             =   6120
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Log"
            Caption         =   "Adodc1"
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
      End
      Begin VB.Frame Frame7 
         Caption         =   "Documentos de ventas que tienen artículoscon pagos parciales."
         ForeColor       =   &H00000080&
         Height          =   2655
         Left            =   90
         TabIndex        =   32
         Top             =   5790
         Width           =   14955
         Begin MSAdodcLib.Adodc bfdetalle 
            Height          =   330
            Left            =   6510
            Top             =   1170
            Visible         =   0   'False
            Width           =   2595
            _ExtentX        =   4577
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from factura_fdetalle where pagado = 'PARCIAL' or pagado = 'SOBRANTE' order by factura.codigo"
            Caption         =   "barregla_factura_fdetalle"
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
         Begin MSAdodcLib.Adodc Adodc9 
            Height          =   345
            Left            =   750
            Top             =   6120
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Log"
            Caption         =   "Adodc1"
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
         Begin MSDataGridLib.DataGrid DataGrid6 
            Bindings        =   "Monito.frx":0038
            Height          =   2295
            Left            =   90
            TabIndex        =   33
            Top             =   240
            Width           =   14745
            _ExtentX        =   26009
            _ExtentY        =   4048
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
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
            ColumnCount     =   29
            BeginProperty Column00 
               DataField       =   "Fecha"
               Caption         =   "Fecha"
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
               DataField       =   "Factura.Codigo"
               Caption         =   "Factura.Codigo"
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
            BeginProperty Column02 
               DataField       =   "Nombre"
               Caption         =   "Nombre"
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
            BeginProperty Column03 
               DataField       =   "fdetalle.Codigo"
               Caption         =   "fdetalle.Codigo"
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
            BeginProperty Column04 
               DataField       =   "Descripcion"
               Caption         =   "Descripcion"
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
            BeginProperty Column05 
               DataField       =   "Cantidad"
               Caption         =   "Cantidad"
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
            BeginProperty Column06 
               DataField       =   "Impuesto"
               Caption         =   "Impuesto"
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
            BeginProperty Column07 
               DataField       =   "Factura.repartidor"
               Caption         =   "Factura.repartidor"
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
            BeginProperty Column08 
               DataField       =   "confirmado"
               Caption         =   "confirmado"
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
            BeginProperty Column09 
               DataField       =   "Domicilio"
               Caption         =   "Domicilio"
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
            BeginProperty Column10 
               DataField       =   "Remito"
               Caption         =   "Remito"
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
            BeginProperty Column11 
               DataField       =   "Cventa"
               Caption         =   "Cventa"
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
            BeginProperty Column12 
               DataField       =   "total_cdo"
               Caption         =   "total_cdo"
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
            BeginProperty Column13 
               DataField       =   "Expr1013"
               Caption         =   "Expr1013"
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
            BeginProperty Column14 
               DataField       =   "fdetalle.repartidor"
               Caption         =   "fdetalle.repartidor"
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
            BeginProperty Column15 
               DataField       =   "Ncomprobante"
               Caption         =   "Ncomprobante"
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
            BeginProperty Column16 
               DataField       =   "pago"
               Caption         =   "pago"
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
            BeginProperty Column17 
               DataField       =   "resta"
               Caption         =   "resta"
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
            BeginProperty Column18 
               DataField       =   "Expr1018"
               Caption         =   "Expr1018"
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
            BeginProperty Column19 
               DataField       =   "Pagado"
               Caption         =   "Pagado"
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
            BeginProperty Column20 
               DataField       =   "tipo"
               Caption         =   "tipo"
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
            BeginProperty Column21 
               DataField       =   "id"
               Caption         =   "id"
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
            BeginProperty Column22 
               DataField       =   "Total"
               Caption         =   "Total"
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
            BeginProperty Column23 
               DataField       =   "Totaliva"
               Caption         =   "Totaliva"
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
            BeginProperty Column24 
               DataField       =   "Total_ctacte"
               Caption         =   "Total_ctacte"
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
            BeginProperty Column25 
               DataField       =   "Expr1"
               Caption         =   "Expr1"
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
            BeginProperty Column26 
               DataField       =   "Expr2"
               Caption         =   "Expr2"
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
            BeginProperty Column27 
               DataField       =   "Detalle"
               Caption         =   "Detalle"
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
            BeginProperty Column28 
               DataField       =   "ÚltimoDeTotal"
               Caption         =   "ÚltimoDeTotal"
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
                  ColumnWidth     =   824.882
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   675.213
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1214.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   750.047
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1349.858
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1005.165
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column07 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column08 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column09 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column10 
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column11 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column12 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column13 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column14 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column15 
                  ColumnWidth     =   1124.787
               EndProperty
               BeginProperty Column16 
                  ColumnWidth     =   945.071
               EndProperty
               BeginProperty Column17 
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column18 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column19 
                  ColumnWidth     =   884.976
               EndProperty
               BeginProperty Column20 
                  ColumnWidth     =   764.787
               EndProperty
               BeginProperty Column21 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column22 
                  ColumnWidth     =   959.811
               EndProperty
               BeginProperty Column23 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column24 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column25 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column26 
                  ColumnWidth     =   14.74
               EndProperty
               BeginProperty Column27 
                  ColumnWidth     =   1950.236
               EndProperty
               BeginProperty Column28 
                  ColumnWidth     =   14.74
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "Totales de Facturas que no coinciden con el total de los detalles de las mismas"
         ForeColor       =   &H00000080&
         Height          =   2835
         Left            =   90
         TabIndex        =   25
         Top             =   2910
         Width           =   11775
         Begin VB.PictureBox Picture5 
            Height          =   615
            Left            =   120
            ScaleHeight     =   555
            ScaleWidth      =   11505
            TabIndex        =   26
            Top             =   270
            Width           =   11565
            Begin VB.CommandButton Command7 
               Caption         =   "Generar Arreglos"
               Height          =   345
               Left            =   6210
               TabIndex        =   30
               Top             =   120
               Width           =   1455
            End
            Begin VB.CommandButton Command6 
               Caption         =   "Buscar proceso"
               Height          =   345
               Left            =   4770
               TabIndex        =   28
               Top             =   120
               Width           =   1455
            End
            Begin VB.TextBox Text2 
               Height          =   285
               Left            =   120
               TabIndex        =   27
               Text            =   "Text1"
               Top             =   150
               Width           =   4335
            End
            Begin MSAdodcLib.Adodc barregla_factura_fdetalle 
               Height          =   330
               Left            =   7320
               Top             =   60
               Visible         =   0   'False
               Width           =   2595
               _ExtentX        =   4577
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
               Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "arregla_fdetalle_totalFactura"
               Caption         =   "barregla_factura_fdetalle"
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
            Begin ComctlLib.ProgressBar b2 
               Height          =   345
               Left            =   7800
               TabIndex        =   31
               Top             =   120
               Width           =   3525
               _ExtentX        =   6218
               _ExtentY        =   609
               _Version        =   327682
               Appearance      =   1
            End
         End
         Begin MSAdodcLib.Adodc Adodc7 
            Height          =   345
            Left            =   750
            Top             =   6120
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Log"
            Caption         =   "Adodc1"
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
         Begin MSDataGridLib.DataGrid DataGrid5 
            Bindings        =   "Monito.frx":0050
            Height          =   1785
            Left            =   120
            TabIndex        =   29
            Top             =   990
            Width           =   11595
            _ExtentX        =   20452
            _ExtentY        =   3149
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "Remito"
               Caption         =   "Remito"
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
               DataField       =   "debito"
               Caption         =   "debito"
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
            BeginProperty Column02 
               DataField       =   "Factura"
               Caption         =   "Factura"
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
            BeginProperty Column03 
               DataField       =   "fdetalle"
               Caption         =   "fdetalle"
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
            BeginProperty Column04 
               DataField       =   "ÚltimoDeCodigo"
               Caption         =   "ÚltimoDeCodigo"
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
            BeginProperty Column05 
               DataField       =   "dif"
               Caption         =   "dif"
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
            BeginProperty Column06 
               DataField       =   "ÚltimoDeTotal"
               Caption         =   "ÚltimoDeTotal"
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
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Debitos que se hayan borrado la factura."
         ForeColor       =   &H00000080&
         Height          =   2415
         Left            =   90
         TabIndex        =   19
         Top             =   450
         Width           =   9045
         Begin VB.PictureBox Picture3 
            Height          =   585
            Left            =   120
            ScaleHeight     =   525
            ScaleWidth      =   8745
            TabIndex        =   21
            Top             =   300
            Width           =   8805
            Begin VB.CommandButton Command5 
               Caption         =   "Imprimir"
               Height          =   405
               Left            =   1380
               TabIndex        =   24
               Top             =   30
               Width           =   1335
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Generar Listado"
               Height          =   405
               Left            =   60
               TabIndex        =   22
               Top             =   30
               Width           =   1335
            End
            Begin ComctlLib.ProgressBar b4 
               Height          =   345
               Left            =   2880
               TabIndex        =   23
               Top             =   60
               Width           =   5685
               _ExtentX        =   10028
               _ExtentY        =   609
               _Version        =   327682
               Appearance      =   1
            End
            Begin MSAdodcLib.Adodc bfd 
               Height          =   330
               Left            =   3390
               Top             =   0
               Visible         =   0   'False
               Width           =   4125
               _ExtentX        =   7276
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
               Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "fdetalle"
               Caption         =   "barregla_factura_fdetalle"
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
         End
         Begin VB.ListBox l3 
            Height          =   1425
            Left            =   150
            TabIndex        =   20
            Top             =   900
            Width           =   8775
         End
         Begin MSAdodcLib.Adodc Adodc6 
            Height          =   345
            Left            =   750
            Top             =   6120
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Log"
            Caption         =   "Adodc1"
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tareas realizadas en el sistemas :"
         ForeColor       =   &H00000080&
         Height          =   3255
         Left            =   -74940
         TabIndex        =   14
         Top             =   330
         Width           =   9915
         Begin VB.PictureBox Picture1 
            Height          =   615
            Left            =   120
            ScaleHeight     =   555
            ScaleWidth      =   9615
            TabIndex        =   15
            Top             =   270
            Width           =   9675
            Begin VB.TextBox Text1 
               Height          =   285
               Left            =   120
               TabIndex        =   17
               Text            =   "Text1"
               Top             =   150
               Width           =   4335
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Buscar proceso"
               Height          =   345
               Left            =   4770
               TabIndex        =   16
               Top             =   120
               Width           =   1455
            End
         End
         Begin MSAdodcLib.Adodc blog 
            Height          =   345
            Left            =   750
            Top             =   6120
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Log"
            Caption         =   "Adodc1"
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
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "Monito.frx":0078
            Height          =   2145
            Left            =   120
            TabIndex        =   18
            Top             =   960
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   3784
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "Id"
               Caption         =   "Id"
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
               DataField       =   "hora"
               Caption         =   "hora"
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
            BeginProperty Column02 
               DataField       =   "fecha"
               Caption         =   "fecha"
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
            BeginProperty Column03 
               DataField       =   "proceso"
               Caption         =   "proceso"
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
            BeginProperty Column04 
               DataField       =   "formulario"
               Caption         =   "formulario"
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
            BeginProperty Column05 
               DataField       =   "comentario"
               Caption         =   "comentario"
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
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Detalles de facturas Borradas:"
         ForeColor       =   &H00000080&
         Height          =   3225
         Left            =   -64980
         TabIndex        =   9
         Top             =   360
         Width           =   5145
         Begin VB.ListBox nofdetalle 
            Height          =   2010
            Left            =   150
            TabIndex        =   13
            Top             =   900
            Width           =   4875
         End
         Begin VB.PictureBox Picture2 
            Height          =   585
            Left            =   120
            ScaleHeight     =   525
            ScaleWidth      =   4815
            TabIndex        =   10
            Top             =   300
            Width           =   4875
            Begin VB.CommandButton Command1 
               Caption         =   "Generar Listado"
               Height          =   405
               Left            =   60
               TabIndex        =   12
               Top             =   60
               Width           =   1335
            End
            Begin ComctlLib.ProgressBar bar 
               Height          =   345
               Left            =   1650
               TabIndex        =   11
               Top             =   90
               Width           =   3105
               _ExtentX        =   5477
               _ExtentY        =   609
               _Version        =   327682
               Appearance      =   1
            End
            Begin MSAdodcLib.Adodc bfactura 
               Height          =   330
               Left            =   -150
               Top             =   -60
               Visible         =   0   'False
               Width           =   2625
               _ExtentX        =   4630
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
               Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "Factura"
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
            Begin MSAdodcLib.Adodc bagrupa_fdetalle 
               Height          =   330
               Left            =   480
               Top             =   120
               Visible         =   0   'False
               Width           =   2385
               _ExtentX        =   4207
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
               Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "agrupa_fdetalle"
               Caption         =   "bagrupa_fdetalle"
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
         End
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   345
            Left            =   750
            Top             =   6120
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Log"
            Caption         =   "Adodc1"
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
      End
      Begin VB.Frame Frame3 
         Caption         =   "Meses desbalancados:"
         ForeColor       =   &H00000080&
         Height          =   2535
         Left            =   -74910
         TabIndex        =   7
         Top             =   3630
         Width           =   15075
         Begin MSAdodcLib.Adodc bpagopormes 
            Height          =   435
            Left            =   3300
            Top             =   2100
            Visible         =   0   'False
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   767
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from pagopormes where saldo < 0"
            Caption         =   "bpagopormes"
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
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   345
            Left            =   750
            Top             =   6120
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Log"
            Caption         =   "Adodc1"
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
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "Monito.frx":008B
            Height          =   1965
            Left            =   150
            TabIndex        =   8
            Top             =   360
            Width           =   14865
            _ExtentX        =   26220
            _ExtentY        =   3466
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
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
            ColumnCount     =   7
            BeginProperty Column00 
               DataField       =   "ÚltimodeFecha"
               Caption         =   "ÚltimodeFecha"
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
               DataField       =   "Codigo"
               Caption         =   "Codigo"
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
            BeginProperty Column02 
               DataField       =   "SumaDeDebito"
               Caption         =   "SumaDeDebito"
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
            BeginProperty Column03 
               DataField       =   "SumaDeCredito"
               Caption         =   "SumaDeCredito"
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
            BeginProperty Column04 
               DataField       =   "saldo"
               Caption         =   "saldo"
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
            BeginProperty Column05 
               DataField       =   "anomes"
               Caption         =   "anomes"
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
            BeginProperty Column06 
               DataField       =   "ÚltimoDeComentario"
               Caption         =   "ÚltimoDeComentario"
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
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column05 
                  ColumnWidth     =   915.024
               EndProperty
               BeginProperty Column06 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Saldo que no coinciden con el saldo final del detalle de los movimientos:"
         ForeColor       =   &H00000080&
         Height          =   2475
         Left            =   -74880
         TabIndex        =   4
         Top             =   6240
         Width           =   10395
         Begin VB.PictureBox Picture4 
            Height          =   315
            Left            =   120
            ScaleHeight     =   255
            ScaleWidth      =   10125
            TabIndex        =   5
            Top             =   270
            Width           =   10185
            Begin MSAdodcLib.Adodc bsaldo_error 
               Height          =   330
               Left            =   6300
               Top             =   60
               Visible         =   0   'False
               Width           =   2385
               _ExtentX        =   4207
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
               Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
               OLEDBFile       =   ""
               DataSourceName  =   ""
               OtherAttributes =   ""
               UserName        =   ""
               Password        =   ""
               RecordSource    =   "saldo_error"
               Caption         =   "bsaldo_error"
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
         End
         Begin MSAdodcLib.Adodc Adodc5 
            Height          =   345
            Left            =   750
            Top             =   6120
            Width           =   3765
            _ExtentX        =   6641
            _ExtentY        =   609
            ConnectMode     =   0
            CursorLocation  =   3
            IsolationLevel  =   -1
            ConnectionTimeout=   15
            CommandTimeout  =   30
            CursorType      =   3
            LockType        =   3
            CommandType     =   2
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
            Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Log"
            Caption         =   "Adodc1"
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
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "Monito.frx":00A5
            Height          =   1755
            Left            =   120
            TabIndex        =   6
            Top             =   630
            Width           =   10125
            _ExtentX        =   17859
            _ExtentY        =   3096
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   3
            FormatLocked    =   -1  'True
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
            ColumnCount     =   5
            BeginProperty Column00 
               DataField       =   "Codigo"
               Caption         =   "Codigo"
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
               DataField       =   "Nombre"
               Caption         =   "Nombre"
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
            BeginProperty Column02 
               DataField       =   "Saldos_Clientes.Expr1"
               Caption         =   "Saldos_Clientes.Expr1"
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
            BeginProperty Column03 
               DataField       =   "Saldo_Ficha.Expr1"
               Caption         =   "Saldo_Ficha.Expr1"
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
            BeginProperty Column04 
               DataField       =   "Error"
               Caption         =   "Error"
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
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1739.906
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1739.906
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "arreglar"
         Height          =   405
         Left            =   -63600
         TabIndex        =   2
         Top             =   8370
         Width           =   3105
      End
      Begin MSComctlLib.ProgressBar dbar 
         Height          =   285
         Left            =   -64350
         TabIndex        =   1
         Top             =   8070
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSAdodcLib.Adodc bctacte 
         Height          =   435
         Left            =   -63750
         Top             =   8550
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   767
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "cuentascorrientes"
         Caption         =   "bctacte"
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
      Begin MSAdodcLib.Adodc bctacte_dif_true_false 
         Height          =   435
         Left            =   -63840
         Top             =   7260
         Visible         =   0   'False
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   767
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\192.168.0.1\c\vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "ctacte_dif_true_false"
         Caption         =   "bctacte_dif_true_false"
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
      Begin MSDataGridLib.DataGrid DataGrid4 
         Bindings        =   "Monito.frx":00C0
         Height          =   1725
         Left            =   -64410
         TabIndex        =   3
         Top             =   6240
         Width           =   4515
         _ExtentX        =   7964
         _ExtentY        =   3043
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
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
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            DataField       =   "SumaDeCredito"
            Caption         =   "SumaDeCredito"
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
         BeginProperty Column02 
            DataField       =   "cFalse"
            Caption         =   "cFalse"
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
         BeginProperty Column03 
            DataField       =   "cTrue"
            Caption         =   "cTrue"
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
         BeginProperty Column04 
            DataField       =   "dif"
            Caption         =   "dif"
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
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   810.142
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1035.213
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "Monito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
bfactura.Refresh
bagrupa_fdetalle.Refresh

bar.Max = bfactura.Recordset.RecordCount + 1
bar.Value = 0

Do Until bfactura.Recordset.EOF
    
    bagrupa_fdetalle.RecordSource = "select * from agrupa_fdetalle where remito = " + Str(bfactura.Recordset("remito"))
    bagrupa_fdetalle.Refresh
    
    If bagrupa_fdetalle.Recordset.RecordCount = 0 Then
       
        nofdetalle.AddItem ("Remito Nro: " + Str(bfactura.Recordset("remito")) + "  " + Str(bfactura.Recordset("fecha")) + "  " + Str(bfactura.Recordset("codigo")) + " " + Str(bfactura.Recordset("total")))
    End If
    
    bar.Value = bar.Value + 1
    
    bfactura.Recordset.MoveNext
Loop


End Sub

Private Sub Command10_Click()
bctacte.RecordSource = "select * from cuentascorrientes where credito > 0"
bctacte.Refresh
Dim sfecha, afecha  As String


Do Until bctacte.Recordset.EOF
    sfecha = Mid(Str(bctacte.Recordset("fechainput")), 7, 4) + Mid(Str(bctacte.Recordset("fechainput")), 4, 2)
afecha = "01/" + Mid(Str(bctacte.Recordset("anomes")), 6, 2) + "/" + Mid(Str(bctacte.Recordset("anomes")), 2, 4)
    If Not Val(sfecha) = bctacte.Recordset("anomes") And bctacte.Recordset("anomes") > 0 Then
        
        
        l7.AddItem (bctacte.Recordset("codigo") + "   id:" + Str(bctacte.Recordset("id")))
    End If
    bctacte.Recordset.MoveNext
Loop
End Sub

Private Sub Command11_Click()
bctacte.RecordSource = "select * from cuentascorrientes where credito >0"
bctacte.Refresh
Dim sfecha As String

' cambia la anomes por lo que tiene en la fecha de inputación


Do Until bctacte.Recordset.EOF
    sfecha = Mid(Format(bctacte.Recordset("fechainput"), "dd/mm/yyyy"), 7, 4) + Mid(Format(bctacte.Recordset("fechainput"), "dd/mm/yyyy"), 4, 2)
    
    If Not Val(sfecha) = bctacte.Recordset("anomes") Then
        l7.AddItem (sfecha + " " + Str(bctacte.Recordset("anomes")) + " " + bctacte.Recordset("codigo"))
        bctacte.Recordset("anomes") = Val(sfecha)
        'bctacte.Recordset("fechainput") = bctacte.Recordset("fecha")
        bctacte.Recordset.Update
    End If
    bctacte.Recordset.MoveNext
Loop
End Sub

Private Sub Command3_Click()

bctacte_dif_true_false.Refresh
dbar.Max = bctacte_dif_true_false.Recordset.RecordCount
bctacte.Refresh
Do Until bctacte_dif_true_false.Recordset.EOF
    
    If Val(Format(bctacte_dif_true_false.Recordset, "000000000000.00")) <> 0 Then
            bctacte.Recordset.AddNew
            bctacte.Recordset("fecha") = date
            bctacte.Recordset("codigo") = bctacte_dif_true_false.Recordset("codigo")
            bctacte.Recordset("credito") = bctacte_dif_true_false.Recordset("dif")
            bctacte.Recordset("noimputar") = True
            bctacte.Recordset("Comentario") = "arreglo de créditos faltantes/sobrantes"
            bctacte.Recordset.Update
    End If
    
    bctacte_dif_true_false.Recordset.MoveNext
    dbar.Value = dbar.Value + 1
    
Loop

End Sub

Private Sub Command4_Click()
bctacte.Refresh


bctacte.RecordSource = "select * from cuentascorrientes where debito > 0 and remito > 0"
bctacte.Refresh

b4.Max = bctacte.Recordset.RecordCount


Do Until bctacte.Recordset.EOF
    bfd.RecordSource = "select * from fdetalle where remito = " + Str(bctacte.Recordset("remito"))
    bfd.Refresh
    
    If bfd.Recordset.EOF Then
        l3.AddItem (bctacte.Recordset("codigo") + "  " + bctacte.Recordset("nombre") + "  " + Str(bctacte.Recordset("debito")) + "  " + Format(bctacte.Recordset("comentario"), "###################################################################"))
    End If

    bctacte.Recordset.MoveNext
    b4.Value = b4.Value + 1
Loop


End Sub

Private Sub Command7_Click()
Dim i As Long
i = 0
barregla_factura_fdetalle.Refresh

b2.Max = barregla_factura_fdetalle.Recordset.RecordCount + 1


Do Until barregla_factura_fdetalle.Recordset.EOF
                
                If barregla_factura_fdetalle.Recordset("remito") = 8187 Then
                Print "0"
                End If
                
    
    
    If Val(Format(barregla_factura_fdetalle.Recordset("dif"), "0000000.00")) > 0 Then
        If Val(Format(barregla_factura_fdetalle.Recordset("factura"), "00000000.00")) = Val(Format(barregla_factura_fdetalle.Recordset("debito"), "000000.00")) Then
            
                i = i + 1
                

                '----- arrreglo debito de cta cte ---------------
                bctacte.RecordSource = "select * from cuentascorrientes where remito = " + Str(barregla_factura_fdetalle.Recordset("remito"))
                bctacte.Refresh
                
                bctacte.Recordset("debito") = barregla_factura_fdetalle.Recordset("fdetalle")
                bctacte.Recordset.Update
                '----------------------------------------------
                
                '------ arreglo total de factuta --------------
                bfactura.RecordSource = "select * from Factura where remito = " + Str(barregla_factura_fdetalle.Recordset("remito"))
                bfactura.Refresh
                
                bfactura.Recordset("total") = barregla_factura_fdetalle.Recordset("fdetalle")
                bfactura.Recordset("totaliva") = barregla_factura_fdetalle.Recordset("fdetalle")
                
                If bfactura.Recordset("total_ctacte") > 0 Then
                    bfactura.Recordset("total_ctacte") = barregla_factura_fdetalle.Recordset("fdetalle")
                End If
                
                bfactura.Recordset.Update
                '----------------------------------------------
        End If
    End If
    
    barregla_factura_fdetalle.Recordset.MoveNext
    b2.Value = b2.Value + 1
Loop

MsgBox "Se modificaron un total de registros: " + Str(i), vbInformation, "Mensaje..."

b2.Value = 0
End Sub

Private Sub Command8_Click()
bctacte.RecordSource = "select * from cuentascorrientes where debito > 0"
bctacte.Refresh
Dim sfecha As String

Do Until bctacte.Recordset.EOF
    sfecha = Mid(Str(bctacte.Recordset("fecha")), 7, 4) + Mid(Str(bctacte.Recordset("fecha")), 4, 2)
    
    If Not Val(sfecha) = bctacte.Recordset("anomes") Then
        bctacte.Recordset("anomes") = Val(sfecha)
        bctacte.Recordset("fechainput") = bctacte.Recordset("fecha")
        bctacte.Recordset.Update
    End If
    bctacte.Recordset.MoveNext
Loop

End Sub

Private Sub Command9_Click()
bctacte.RecordSource = "select * from cuentascorrientes where credito > 0"
bctacte.Refresh
Dim sfecha, afecha  As String

Do Until bctacte.Recordset.EOF
    sfecha = Mid(Format(bctacte.Recordset("fechainput"), "dd/mm/yyyy"), 7, 4) + Mid(Format(bctacte.Recordset("fechainput"), "dd/mm/yyyy"), 4, 2)
afecha = "01/" + Mid(Format(bctacte.Recordset("anomes"), "dd/mm/yyyy"), 6, 2) + "/" + Mid(Format(bctacte.Recordset("anomes"), "dd/mm/yyyy"), 2, 4)
    If Not Val(sfecha) = bctacte.Recordset("anomes") And bctacte.Recordset("anomes") > 0 Then
        l7.AddItem (Str(bctacte.Recordset("id")) + "  " + bctacte.Recordset("codigo"))
    End If
    bctacte.Recordset.MoveNext
Loop
End Sub

Private Sub List1_Click()


End Sub

