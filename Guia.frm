VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmGuia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guía Telefónica"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   13095
   Begin MSAdodcLib.Adodc bguia 
      Height          =   405
      Left            =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   714
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
      Caption         =   "bguia"
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
      Height          =   6600
      Left            =   0
      TabIndex        =   14
      ToolTipText     =   "Configuración de parámetros. Porcentaje de ganancia"
      Top             =   -360
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   11642
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ingresar datos"
      TabPicture(0)   =   "Guia.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdFormatoPlanilla"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdNuevo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdSiguiente"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBuscar_Alta"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdPrevio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "v1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "bot"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "v1(1)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "v1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "v1(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "v1(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "v1(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "v1(8)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "iva"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "v1(9)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "v1(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblTitulo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(6)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(3)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(4)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(7)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(8)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(9)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "Forma Planilla"
      TabPicture(1)   =   "Guia.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "DgGuia"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraBusqueda"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraOrdena"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraBotonera"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame fraBotonera 
         Height          =   705
         Left            =   8520
         TabIndex        =   41
         Top             =   5520
         Width           =   2625
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "Borrar"
            Height          =   495
            Left            =   1320
            Picture         =   "Guia.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Borrar datos"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir"
            Height          =   495
            Left            =   60
            Picture         =   "Guia.frx":013A
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Generar reporte para imprimir"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame fraOrdena 
         Caption         =   "Ordenado por :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   8520
         TabIndex        =   38
         Top             =   4950
         Width           =   2565
         Begin VB.OptionButton op1 
            Caption         =   "Nombre"
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   330
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton op2 
            Caption         =   "Código"
            Height          =   195
            Left            =   1170
            TabIndex        =   39
            Top             =   330
            Width           =   945
         End
      End
      Begin VB.CommandButton cmdFormatoPlanilla 
         Caption         =   "Formato Planilla"
         Height          =   495
         Left            =   -71460
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ir a planilla de contáctos"
         Top             =   4860
         Width           =   1245
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   -72270
         Picture         =   "Guia.frx":023C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Nuevo contácto"
         Top             =   4860
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton cmdSiguiente 
         BackColor       =   &H80000004&
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -68700
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4830
         Width           =   525
      End
      Begin VB.Frame fraBusqueda 
         Height          =   1215
         Left            =   180
         TabIndex        =   32
         Top             =   5040
         Width           =   6945
         Begin VB.CommandButton cmdbuscar 
            Caption         =   "Buscar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   5760
            TabIndex        =   46
            Top             =   180
            Width           =   1095
         End
         Begin VB.TextBox Txtocupacion 
            Height          =   285
            Left            =   1080
            TabIndex        =   44
            ToolTipText     =   "Presionar enter para ejectutar la consulta. Filtra por código y por descripción simultaneamente"
            Top             =   800
            Width           =   4605
         End
         Begin VB.TextBox Txtlocalidad 
            Height          =   285
            Left            =   1080
            TabIndex        =   35
            ToolTipText     =   "Presionar enter para ejectutar la consulta. Filtra por código y por descripción simultaneamente"
            Top             =   500
            Width           =   4605
         End
         Begin VB.TextBox txtNombre 
            Height          =   285
            Left            =   1080
            TabIndex        =   34
            ToolTipText     =   "Presionar enter para ejectutar la consulta. Filtra por código y por descripción simultaneamente"
            Top             =   180
            Width           =   4605
         End
         Begin VB.Label lblOcupación 
            AutoSize        =   -1  'True
            Caption         =   "Ocupación :"
            Height          =   195
            Left            =   120
            TabIndex        =   45
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Localidad :"
            Height          =   195
            Left            =   150
            TabIndex        =   37
            Top             =   525
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Contácto :"
            Height          =   195
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame3 
         Height          =   45
         Left            =   -75180
         TabIndex        =   31
         Top             =   780
         Width           =   11865
      End
      Begin VB.Frame Frame1 
         Height          =   45
         Left            =   -75240
         TabIndex        =   30
         Top             =   4740
         Width           =   9555
      End
      Begin VB.CommandButton cmdBuscar_Alta 
         Caption         =   "Buscar"
         Height          =   495
         Left            =   -73050
         Picture         =   "Guia.frx":033E
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Buscar contácto"
         Top             =   4860
         UseMaskColor    =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton cmdPrevio 
         BackColor       =   &H80000004&
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   -69210
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4830
         Width           =   525
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   7
         Left            =   -72795
         TabIndex        =   7
         Top             =   3510
         Width           =   2670
      End
      Begin VB.Frame bot 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   -74610
         TabIndex        =   16
         Top             =   4830
         Visible         =   0   'False
         Width           =   1575
         Begin VB.CommandButton cmdBorrar_Alta 
            Caption         =   "Borrar"
            Height          =   495
            Left            =   780
            Picture         =   "Guia.frx":0440
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Borrar datos"
            Top             =   30
            UseMaskColor    =   -1  'True
            Width           =   795
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "Guardar"
            Height          =   495
            Left            =   0
            Picture         =   "Guia.frx":0542
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Guardar datos"
            Top             =   30
            UseMaskColor    =   -1  'True
            Width           =   795
         End
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   1
         Left            =   -72795
         TabIndex        =   1
         Top             =   1320
         Width           =   4635
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   0
         Left            =   -72795
         TabIndex        =   0
         Top             =   960
         Width           =   2685
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   2
         Left            =   -72795
         TabIndex        =   2
         Top             =   1695
         Width           =   1875
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   3
         Left            =   -72795
         TabIndex        =   3
         Top             =   2055
         Width           =   1875
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   4
         Left            =   -72795
         TabIndex        =   4
         Top             =   2415
         Width           =   1875
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   8
         Left            =   -72795
         TabIndex        =   8
         Top             =   3870
         Width           =   2655
      End
      Begin VB.ComboBox iva 
         Height          =   315
         ItemData        =   "Guia.frx":0644
         Left            =   -72795
         List            =   "Guia.frx":0651
         TabIndex        =   5
         Text            =   "Responsable Inscripto"
         Top             =   2820
         Width           =   2610
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   9
         Left            =   -72795
         TabIndex        =   9
         Top             =   4230
         Width           =   2655
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   6
         Left            =   -72795
         TabIndex        =   6
         Top             =   3150
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid DgGuia 
         Bindings        =   "Guia.frx":068B
         Height          =   4575
         Left            =   0
         TabIndex        =   15
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777152
         HeadLines       =   2
         RowHeight       =   15
         FormatLocked    =   -1  'True
         AllowDelete     =   -1  'True
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
         ColumnCount     =   13
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
            DataField       =   "Direccion"
            Caption         =   "Direccion"
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
            DataField       =   "Localidad"
            Caption         =   "Localidad"
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
            DataField       =   "Telefono"
            Caption         =   "Telefono"
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
            DataField       =   "Mutual"
            Caption         =   "Mutual"
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
            DataField       =   "Iva"
            Caption         =   "Iva"
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
            DataField       =   "Cuit"
            Caption         =   "Cuit"
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
            DataField       =   "Correo"
            Caption         =   "Correo"
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
            DataField       =   "Ocupación"
            Caption         =   "Ocupación"
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
            DataField       =   "Ibrutos"
            Caption         =   "Ibrutos"
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
            DataField       =   "Saldo"
            Caption         =   "Saldo"
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
         SplitCount      =   1
         BeginProperty Split0 
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnWidth     =   854.929
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
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Ingresar Contácto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   315
         Left            =   -74685
         TabIndex        =   27
         Top             =   420
         Width           =   6585
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Correo Electrónico :"
         Height          =   285
         Index           =   6
         Left            =   -74970
         TabIndex        =   26
         Top             =   3885
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Tipo de I.V.A. :"
         Height          =   255
         Index           =   2
         Left            =   -74970
         TabIndex        =   25
         Top             =   2805
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Nombre y Apellido :"
         Height          =   345
         Index           =   1
         Left            =   -74970
         TabIndex        =   24
         Top             =   1335
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Código  del Contácto :"
         Height          =   405
         Index           =   0
         Left            =   -74970
         TabIndex        =   23
         Top             =   960
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Dirección :"
         Height          =   225
         Index           =   3
         Left            =   -74970
         TabIndex        =   22
         Top             =   1695
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Teléfono :"
         Height          =   405
         Index           =   4
         Left            =   -74970
         TabIndex        =   21
         Top             =   2445
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Localidad :"
         Height          =   285
         Index           =   5
         Left            =   -74970
         TabIndex        =   20
         Top             =   2070
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Ocupación  :"
         Height          =   315
         Index           =   7
         Left            =   -74970
         TabIndex        =   19
         Top             =   3195
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Ing. Bruto :"
         Height          =   345
         Index           =   8
         Left            =   -74970
         TabIndex        =   18
         Top             =   4200
         Width           =   2150
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> C.U.I.T.  :"
         Height          =   315
         Index           =   9
         Left            =   -74970
         TabIndex        =   17
         Top             =   3540
         Width           =   2150
      End
   End
End
Attribute VB_Name = "frmGuia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const t = 9
Dim sqlp As String
Dim vfilter As String
Dim ordenado As String

Private Sub cmdBorrar_Alta_Click()
    On Error Resume Next
    

    If Err Then GrabarLog "cmdBorrar_Alta_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdBorrar_Click()
    On Error Resume Next
    
    Borrar bguia.object, True
    
    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdBuscar_Alta_Click()
On Error Resume Next

    With frmBuscarContacto
        .o = 0
        .Show
    End With

If Err Then GrabarLog "cmdBuscar_Alta_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdBuscar_Click()
    On Error Resume Next

    If op1.Value = True Then
        ordenado = "Nombre"
    Else
        ordenado = "Codigo"
    End If

    sqlp = ""
    vfilter = ""

    If Not txtLocalidad = "" Then
        sqlp = sqlp + " and Localidad Like '%" + Trim(txtLocalidad) + "%'"
        vfilter = vfilter + " and ([Localidad] Like '*" + Trim(txtLocalidad) + "*')"
    End If

    If Not Txtocupacion = "" Then
        sqlp = sqlp + " and Ocupación Like '%" + Trim(Txtocupacion) + "%'"
        vfilter = vfilter + " and ([Ocupación] Like '*" + Trim(Txtocupacion) + "*')"
    End If

    If Not txtNombre = "" Then
        sqlp = sqlp + " and (nombre Like '%" + Trim(txtNombre) + "%') or (codigo = '" + Trim(txtNombre) + "')"
        vfilter = vfilter + " and ([nombre] like '*" + Trim(txtNombre) + "*') or ([codigo] = '" + Trim(txtNombre) + "')"
    End If

    With bguia
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Guia WHERE 1=1 " + sqlp + " ORDER BY " + ordenado
        .Refresh
        If Not .Recordset.EOF = True Then .Recordset.MoveLast
    End With
    
    If Err Then GrabarLog "cmdbuscar_Click", Err.Number & "-" & Err.Description, Me.Name
End Sub

Private Sub cmdFormatoPlanilla_Click()
On Error Resume Next

    TabGeneral.Tab = 1

    If Err Then GrabarLog "cmdBorrar_Alta_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdGuardar_Click()
    On Error Resume Next
    
    With bguia
        .Recordset.MoveFirst
        .Recordset.Find ("codigo = '" + Trim(v1(0).Text) + "'")
    
        If bguia.Recordset.EOF Then .Recordset.AddNew
    
        .Recordset("codigo").Value = v1(0).Text
        .Recordset(1).Value = v1(1).Text      ' nombre
        .Recordset(2).Value = v1(2).Text      ' direccion
        .Recordset(3).Value = v1(3).Text      ' localidad
        .Recordset(4).Value = v1(4).Text      ' telefono
        .Recordset(5).Value = iva.Text   ' iva
        .Recordset(6).Value = v1(6).Text      ' cuit
        .Recordset(7).Value = v1(7).Text      ' credito
        .Recordset(8).Value = v1(8).Text      ' responsable
        .Recordset(9).Value = v1(9).Text      ' responsable
        
        .Recordset.Update
    End With

    Limpia ' limpia los campos de entrada

    If vConfigGral.vIncluyeContabilidad = True Then
        With frmAsientosAlta
            .Show
            .ZOrder (0)
            .txtCuentaVieneDe.Text = Me.Caption
        End With
    End If
    
If Err Then GrabarLog "cmdGuardar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdImprimir_Click()
    On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    With Mantenimiento.rsguia
        
        If .State = 0 Then
            .Open
            .Close
            .Open
        Else
            .Close
            .Open
        End If
        
        .filter = "([id] > 0)" + vfilter
        .Sort = ordenado
    End With
    With drguia
        .Show
    End With
    
    If Err Then GrabarLog "cmdImprimir_Click", Err.Number & "-" & Err.Description, Me.Name
End Sub

Private Sub cmdNuevo_Click()
On Error Resume Next

    Limpia
    
If Err Then GrabarLog "cmdNuevo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdPrevio_Click()
    On Error Resume Next

    ' frmPrincipal.display.Caption = "Un artículo fue guardado exitosamente"
    With bguia
        .Recordset.MovePrevious
    End With

    mostrar

    If Err Then
        bguia.Recordset.MoveFirst
        mostrar
        Exit Sub
    End If

End Sub
Private Sub cmdSiguiente_Click()
    On Error Resume Next

    'frmPrincipal.display.Caption = "Un artículo fue guardado exitosamente"
    With bguia
        .Recordset.MoveNext
    End With

    mostrar

    If Err Then
        bguia.Recordset.MoveLast
        mostrar
        Exit Sub
    End If

End Sub

Private Sub DgGuia_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    OrdenarDataGrid ColIndex, bguia.Recordset, DgGuia

If Err Then GrabarLog "DgGuia_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next

    With bguia
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Guia"
        .Refresh
    End With
    
    With Me
        .Top = 500
        .Left = 2900
        .Width = 7400
        .Height = 5460
    End With
    
    TabGeneral.Tab = 0

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpia()
    On Error Resume Next
    Dim i As Integer

    For i = 0 To t
        If i = 5 Then i = 6
        v1(i).Text = ""
    Next

    v1(0).SetFocus

    If Err Then GrabarLog "limpia", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub mostrar()
    On Error Resume Next
    Dim i As Integer

    With bguia

        For i = 0 To t

            If i = 5 Then i = 6
            v1(i).Text = .Recordset(i).Value
        Next

    End With

    If Err Then GrabarLog "mostrar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub TabGeneral_Click(PreviousTab As Integer)
On Error Resume Next

    Select Case TabGeneral.Tab

        Case 0
            With Me
                .Top = 600
                .Left = 1500
                .Height = 5460
                .Width = 7400
            End With
        Case 1
            With Me
                .Top = 400
                .Left = 650
                .Height = 6800
                .Width = 11300
            End With
            
            cmdBuscar_Click
        
        Case 2
    
    End Select

If Err Then GrabarLog "TabGeneral_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtLocalidad_Change()
On Error Resume Next

    cmdBuscar_Click
    
If Err Then GrabarLog "cmdBuscar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtNombre_Change()
On Error Resume Next
    
    cmdBuscar_Click

If Err Then GrabarLog "txtNombre_Change", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Txtocupacion_Change()
On Error Resume Next

    cmdBuscar_Click
    
If Err Then GrabarLog "Txtocupacion_Change", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub v1_Change(Index As Integer)
On Error Resume Next

    If Not v1(0) = "" Then
        bot.Visible = True
    Else
        bot.Visible = False
    End If
    
If Err Then GrabarLog "v1_Change", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub v1_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        If Index >= t Then
            cmdGuardar.SetFocus
        Else

            If Index = 4 Then Index = 5
            v1(Index + 1).SetFocus
        End If

    End If
    
If Err Then GrabarLog "v1_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub v1_LostFocus(Index As Integer)
    On Error Resume Next

    If Index = 0 Then
        With bguia
            .RecordSource = "Select * from guia"
            .Refresh
            If .Recordset.EOF = True Then Exit Sub
            
            .Recordset.MoveFirst
            .Recordset.Find ("codigo = '" + Trim(v1(0).Text) + "'")
    
            If Not .Recordset.EOF Then
                MsgBox "Código existente !", vbInformation, "Mensaje..."
                v1(0) = ""
                v1(0).SetFocus
            End If
        End With
    End If

    If Err Then GrabarLog "v1_LostFocus", Err.Number & " " & Err.Description, Me.Name
End Sub


