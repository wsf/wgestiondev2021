VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComentario 
   Caption         =   "Módulo de Alta y Modificacion de Comentarios a Clientes y Empleados"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16620
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   16620
   Begin MSAdodcLib.Adodc bComentario 
      Height          =   375
      Left            =   120
      Top             =   8760
      Visible         =   0   'False
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   661
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
      Caption         =   "bcomentario"
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
   Begin TabDlg.SSTab TabComentarios 
      Height          =   8655
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   14925
      _ExtentX        =   26326
      _ExtentY        =   15266
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   10842463
      TabCaption(0)   =   "< Alta >"
      TabPicture(0)   =   "frmComentario.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblTitulos(1)"
      Tab(0).Control(1)=   "lblTitulos(3)"
      Tab(0).Control(2)=   "lblTitulos(4)"
      Tab(0).Control(3)=   "lblTitulos(2)"
      Tab(0).Control(4)=   "lblTitulos(0)"
      Tab(0).Control(5)=   "lblTituloGral"
      Tab(0).Control(6)=   "txtlocalidad"
      Tab(0).Control(7)=   "txtcomentario"
      Tab(0).Control(8)=   "txtcliente"
      Tab(0).Control(9)=   "cboRepartidor"
      Tab(0).Control(10)=   "cboReparto"
      Tab(0).Control(11)=   "Frame1"
      Tab(0).Control(12)=   "Frame4"
      Tab(0).Control(13)=   "Frame5"
      Tab(0).Control(14)=   "fram"
      Tab(0).Control(15)=   "cmdVerCliente"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "< Mantenimiento >"
      TabPicture(1)   =   "frmComentario.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DgComentario"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).ControlCount=   3
      Begin VB.CommandButton cmdVerCliente 
         Height          =   495
         Left            =   -61320
         TabIndex        =   42
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame fram 
         Height          =   45
         Left            =   -74460
         TabIndex        =   41
         Top             =   1440
         Width           =   13065
      End
      Begin VB.Frame Frame5 
         Height          =   45
         Left            =   -74280
         TabIndex        =   40
         Top             =   3270
         Width           =   13065
      End
      Begin VB.Frame Frame4 
         Height          =   1965
         Left            =   -73620
         TabIndex        =   32
         Top             =   5160
         Width           =   12525
         Begin VB.TextBox txtrcantidad 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   5640
            TabIndex        =   33
            Text            =   "1"
            Top             =   1500
            Width           =   885
         End
         Begin MSComCtl2.DTPicker vfdesde 
            Height          =   315
            Left            =   2265
            TabIndex        =   34
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   51576833
            CurrentDate     =   38573
         End
         Begin MSComCtl2.DTPicker vfhasta 
            Height          =   315
            Left            =   4665
            TabIndex        =   35
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Format          =   51576833
            CurrentDate     =   38573
         End
         Begin VB.Label lblPeríodoQue 
            Caption         =   "> Período que desea que salga el comentario"
            Height          =   195
            Index           =   0
            Left            =   630
            TabIndex        =   39
            Top             =   360
            Width           =   3540
         End
         Begin VB.Label lblFechaDesde 
            Alignment       =   1  'Right Justify
            Caption         =   "> Fecha Desde:"
            Height          =   195
            Index           =   5
            Left            =   870
            TabIndex        =   38
            Top             =   825
            Width           =   1350
         End
         Begin VB.Label lblFHasta 
            Alignment       =   1  'Right Justify
            Caption         =   "> F. Hasta:"
            Height          =   195
            Index           =   6
            Left            =   3585
            TabIndex        =   37
            Top             =   825
            Width           =   1020
         End
         Begin VB.Label lblCRepaticiones 
            Caption         =   "> Cantidad de reepaticiones que desea que se imprima el comentario:"
            Height          =   195
            Index           =   7
            Left            =   570
            TabIndex        =   36
            Top             =   1530
            Width           =   5130
         End
      End
      Begin VB.Frame Frame3 
         Height          =   855
         Left            =   60
         TabIndex        =   28
         Top             =   7410
         Width           =   14775
         Begin VB.CommandButton cmdAcciones 
            Caption         =   "Imprimir"
            Height          =   615
            Index           =   6
            Left            =   2190
            Picture         =   "frmComentario.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdAcciones 
            Caption         =   "Modificar"
            Height          =   615
            Index           =   5
            Left            =   1110
            Picture         =   "frmComentario.frx":056A
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
         Begin VB.CommandButton cmdAcciones 
            Caption         =   "Borrar"
            Height          =   615
            Index           =   4
            Left            =   30
            Picture         =   "frmComentario.frx":0A9C
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame1 
         Height          =   795
         Left            =   -74970
         TabIndex        =   24
         Top             =   7470
         Width           =   14805
         Begin VB.CommandButton cmdAcciones 
            Caption         =   "Formato Planilla"
            Height          =   585
            Index           =   2
            Left            =   1830
            Picture         =   "frmComentario.frx":0FCE
            TabIndex        =   25
            ToolTipText     =   "Imprimir"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   1155
         End
         Begin VB.CommandButton cmdAcciones 
            Caption         =   "Limpiar"
            Height          =   585
            Index           =   1
            Left            =   925
            Picture         =   "frmComentario.frx":10D0
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Ejecutar búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   915
         End
         Begin VB.CommandButton cmdAcciones 
            Caption         =   "Guardar"
            Height          =   585
            Index           =   0
            Left            =   30
            Picture         =   "frmComentario.frx":1602
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Ejecutar búsqueda"
            Top             =   150
            UseMaskColor    =   -1  'True
            Width           =   915
         End
      End
      Begin MSDataGridLib.DataGrid DgComentario 
         Bindings        =   "frmComentario.frx":1704
         Height          =   6255
         Left            =   90
         TabIndex        =   6
         Top             =   1080
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   11033
         _Version        =   393216
         BackColor       =   15856113
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
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
         ColumnCount     =   12
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
         BeginProperty Column03 
            DataField       =   "Cod_Reparto"
            Caption         =   "Cod_Reparto"
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
            DataField       =   "Reparto"
            Caption         =   "Reparto"
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
            DataField       =   "Cod_Repartidor"
            Caption         =   "Cod_Repartidor"
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
            DataField       =   "Repartidor"
            Caption         =   "Repartidor"
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
            DataField       =   "Comentario"
            Caption         =   "Comentario"
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
            DataField       =   "Repeticiones"
            Caption         =   "Repeticiones"
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
            DataField       =   "FDesde"
            Caption         =   "FDesde"
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
            DataField       =   "FHasta"
            Caption         =   "FHasta"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2415.118
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   3839.811
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   0
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox cboReparto 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5715F&
         Height          =   360
         Left            =   -73560
         TabIndex        =   2
         Top             =   2040
         Width           =   6765
      End
      Begin VB.ComboBox cboRepartidor 
         BackColor       =   &H80000016&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5715F&
         Height          =   360
         Left            =   -73560
         TabIndex        =   3
         Top             =   2490
         Width           =   6765
      End
      Begin VB.TextBox txtcliente 
         Height          =   315
         Left            =   -73440
         TabIndex        =   0
         Top             =   720
         Width           =   12045
      End
      Begin VB.TextBox txtcomentario 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1455
         Left            =   -73680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   3690
         Width           =   12555
      End
      Begin VB.TextBox txtlocalidad 
         Height          =   315
         Left            =   -73560
         TabIndex        =   1
         Top             =   1680
         Width           =   3705
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   90
         TabIndex        =   7
         Top             =   30
         Width           =   14715
         Begin VB.CommandButton cmdAcciones 
            Caption         =   "Filtrar!"
            Height          =   345
            Index           =   3
            Left            =   11580
            Picture         =   "frmComentario.frx":171E
            TabIndex        =   15
            Top             =   540
            UseMaskColor    =   -1  'True
            Width           =   3015
         End
         Begin VB.ComboBox cbocReparto 
            Height          =   315
            Left            =   7080
            TabIndex        =   13
            Top             =   120
            Width           =   4335
         End
         Begin VB.ComboBox cbocRepartidor 
            Height          =   315
            Left            =   7080
            TabIndex        =   14
            Top             =   540
            Width           =   4335
         End
         Begin VB.TextBox txtcComentario 
            Height          =   315
            Left            =   1320
            TabIndex        =   12
            Top             =   540
            Width           =   4635
         End
         Begin VB.TextBox txtcNombre 
            Height          =   315
            Left            =   1320
            TabIndex        =   11
            Top             =   180
            Width           =   4635
         End
         Begin MSComCtl2.DTPicker vFecha 
            Height          =   315
            Left            =   11580
            TabIndex        =   23
            Top             =   180
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   556
            _Version        =   393216
            Format          =   51576833
            CurrentDate     =   38573
         End
         Begin VB.Label lblTitulos 
            Alignment       =   1  'Right Justify
            Caption         =   "> Repartidor :"
            Height          =   195
            Index           =   11
            Left            =   6045
            TabIndex        =   22
            Top             =   600
            Width           =   1005
         End
         Begin VB.Label lblTitulos 
            Alignment       =   1  'Right Justify
            Caption         =   "> Reparto :"
            Height          =   195
            Index           =   10
            Left            =   6045
            TabIndex        =   21
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lblTitulos 
            Alignment       =   1  'Right Justify
            Caption         =   "> Comentario :"
            Height          =   195
            Index           =   9
            Left            =   30
            TabIndex        =   20
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label lblTitulos 
            Alignment       =   1  'Right Justify
            Caption         =   "> Cliente :"
            Height          =   195
            Index           =   8
            Left            =   30
            TabIndex        =   19
            Top             =   240
            Width           =   1200
         End
      End
      Begin VB.Label lblTituloGral 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00A5715F&
         Caption         =   "Alta de Comentarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   -75000
         TabIndex        =   18
         Top             =   0
         Width           =   14895
      End
      Begin VB.Label lblTitulos 
         Alignment       =   1  'Right Justify
         Caption         =   "> Cliente :"
         Height          =   195
         Index           =   0
         Left            =   -74835
         TabIndex        =   17
         Top             =   780
         Width           =   1350
      End
      Begin VB.Label lblTitulos 
         Alignment       =   1  'Right Justify
         Caption         =   "> Reparto :"
         Height          =   195
         Index           =   2
         Left            =   -75000
         TabIndex        =   16
         Top             =   2160
         Width           =   1400
      End
      Begin VB.Label lblTitulos 
         Alignment       =   1  'Right Justify
         Caption         =   "> Comentario :"
         Height          =   195
         Index           =   4
         Left            =   -75000
         TabIndex        =   10
         Top             =   3630
         Width           =   1350
      End
      Begin VB.Label lblTitulos 
         Alignment       =   1  'Right Justify
         Caption         =   "> Repartidor :"
         Height          =   195
         Index           =   3
         Left            =   -75000
         TabIndex        =   9
         Top             =   2535
         Width           =   1400
      End
      Begin VB.Label lblTitulos 
         Alignment       =   1  'Right Justify
         Caption         =   "> Localidad :"
         Height          =   195
         Index           =   1
         Left            =   -75000
         TabIndex        =   8
         Top             =   1715
         Width           =   1400
      End
   End
End
Attribute VB_Name = "frmComentario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo, vlocalidad As String
Dim vModifica As Boolean
Dim vBandera As Long
Dim sql As String
Private Function BuscarCliente() As Boolean
On Error Resume Next
    
    Dim ConnClientes As New ADODB.Connection
    Dim rsClientes As New ADODB.Recordset
    Dim sqlClientes As String
    
    With ConnClientes
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlClientes = "SELECT * FROM clientes WHERE (codigo = '" + Trim(txtCliente.Text) + "')"
    
    With rsClientes
        Call .Open(sqlClientes, ConnClientes, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            Limpiar
            txtCliente.Text = .Fields("Nombre").Value
            vCodigo = .Fields("Codigo").Value
            txtlocalidad.Text = .Fields("Localidad").Value
            
            cboRepartidor.Text = BuscarDato("Empleados WHERE codigo = '" & .Fields("Repartidor").Value & "'", "Nombre")
            cboReparto.Text = BuscarDato("clireparto WHERE nreparto = '" & .Fields("Reparto").Value & "'", "descrip")
            txtComentario.SetFocus
        Else
            
            'cmdVerCliente_Click
            frmBuscarCliente.txtClientes.SetFocus
            frmBuscarCliente.o = 10
            frmBuscarCliente.txtClientes.Text = txtCliente.Text
            frmBuscarCliente.Show
                                       
        End If

    End With

    sqlClientes = ""
    
    rsClientes.Close
    Set rsClientes = Nothing
    
    ConnClientes.Close
    Set ConnClientes = Nothing
    
If Err Then GrabarLog "BuscarCliente", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub BorrarRegistro()
On Error Resume Next

    Borrar bComentario.object, True

If Err Then GrabarLog "BorrarRegistro", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cbocRepartidor_GotFocus()
On Error Resume Next
    
    Call CargarCombo("Empleados", "Nombre", cbocRepartidor, False, "Codigo")

If Err Then GrabarLog "cboRepartidor_Filtro_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboRepartidor_GotFocus()
On Error Resume Next
    
    Call CargarCombo("Empleados", "Nombre", cboRepartidor, False, "Codigo")
    
If Err Then GrabarLog "cboRepartidor_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cbocReparto_GotFocus()
On Error Resume Next
    
    Call CargarCombo("clireparto", "descrip", cbocReparto, False, "nreparto")

If Err Then GrabarLog "cboRepartoFiltro_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboreparto_GotFocus()
On Error Resume Next
    
    Call CargarCombo("clireparto", "descrip", cboReparto, False, "nreparto")
    
If Err Then GrabarLog "cboReparto_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdAcciones_Click(Index As Integer)
On Error Resume Next
    Select Case Index
    
        Case 0
            
            Guardar
        Case 1
            Limpiar
        Case 2
            FormatoPlanilla
        Case 3
            Filtrar
        Case 4
            BorrarRegistro
        Case 5
            Modificar
        Case 6
            Imprimir
    
    End Select

If Err Then GrabarLog "cmdAcciones_Click :" & Index, Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdVerCliente_Click()
On Error Resume Next

    With frmBuscarCliente
        .txtClientes.Text = (txtCliente.Text)
        .Show
    End With
    
If Err Then GrabarLog "cmdVerCliente_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub DgComentario_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
    
    OrdenarDataGrid ColIndex, bComentario.Recordset, DgComentario
    
If Err Then GrabarLog "DgComentario_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Filtrar()
On Error Resume Next
    
    sql = ""
    
    If Not Trim(txtcNombre.Text) = "" Then
        sql = sql + " and (Nombre like '%" & Trim(txtcNombre.Text) & "%')"
    End If
    
    If Not Trim(txtcComentario.Text) = "" Then
        sql = sql + " and (Comentario like '%" & Trim(txtcComentario.Text) & "%')"
    End If
    
    If Not Trim(cbocReparto.Text) = "" Then
        sql = sql + " and (Cod_Reparto = '" & Trim(cbocReparto.ItemData(cbocReparto.ListIndex)) & "')"
    End If
    
    If Not Trim(cbocRepartidor.Text) = "" Then
        sql = sql + " and (Cod_Repartidor = '" & Trim(cbocRepartidor.ItemData(cbocRepartidor.ListIndex)) & "')"
    End If
    
    With bComentario
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM comentarios WHERE 1=1" & sql & " ORDER BY id"
        .Refresh
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
    
    End With
    
If Err Then GrabarLog "Filtrar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then SendKeys "{TAB}"
            
If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next

    With bComentario
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM comentarios ORDER BY id ASC"
        .Refresh
    End With
    
    vfdesde.Value = Date
    vfhasta.Value = Date

    With Me
        .Height = 9330
        .Width = 15090
        .KeyPreview = True
        .TabComentarios.Tab = 1
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoPlanilla()
On Error Resume Next

    TabComentarios.Tab = 1

If Err Then GrabarLog "FormatoPlanilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Function BuscarDato(vTabla As String, vCampo As String) As String
    On Error Resume Next
    
    Dim connDato As New ADODB.Connection
    Dim rsDato As New ADODB.Recordset
    Dim sqlDato As String
    
    With connDato
        .ConnectionString = pathDBMySQL
        .Open
    End With
        
    sqlDato = "SELECT * FROM " & vTabla & ""
            
    With rsDato
        Call .Open(sqlDato, connDato, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            BuscarDato = .Fields(vCampo).Value
        Else
            BuscarDato = ""
        End If
        
    End With
    
    sqlDato = ""
    
    rsDato.Close
    Set rsDato = Nothing

    connDato.Close
    Set connDato = Nothing
    
    If Err Then GrabarLog "BuscarDato", Err.Number & "-" & Err.Description, Me.Name
End Function
Private Sub Guardar()
On Error Resume Next
   
    If Val(txtrcantidad) = 0 Then
        MsgBox "Debe poner ingresar nro de repeticiones del anuncio", vbCritical
        Exit Sub
    End If

    If (vCodigo = "") Then
        MsgBox "Por Favor Elija un CLIENTE para Asignarle un comentario", vbInformation, "Mensaje ..."
        Exit Sub
    End If
    
    If (Val(txtrcantidad.Text) = 0) And (vfdesde.Enabled = False) Then
        MsgBox "Por Favor Elija un Método de Repetición para este comentario", vbInformation, "Mensaje ..."
        Exit Sub
    End If
    
    With bComentario
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM comentarios ORDER BY id ASC"
        .Refresh
        
        If vModifica = True Then
            If Not .Recordset.EOF = True Then .Recordset.MoveFirst
            .Recordset.Find ("id = " & vBandera & "")
            If .Recordset.EOF = True Then .Recordset.AddNew
            
        Else
        
            .Recordset.AddNew
        
        End If
        
        .Recordset("Codigo").Value = vCodigo
        .Recordset("Nombre").Value = Left(txtCliente.Text, 50)
        .Recordset("Localidad").Value = Left(txtlocalidad.Text, 50)
        
        If vModifica = Not True Then
            .Recordset("Cod_Reparto").Value = Trim(cboReparto.ItemData(cboReparto.ListIndex))
            .Recordset("Reparto").Value = Left(cboReparto.Text, 50)
                
            .Recordset("Cod_Repartidor").Value = Trim(cboRepartidor.ItemData(cboRepartidor.ListIndex))
            .Recordset("Repartidor").Value = cboRepartidor.Text
        
        End If

        .Recordset("Comentario").Value = Left(txtComentario.Text, 254)
        
        If Val(txtrcantidad.Text) = 0 Then
            .Recordset("Fdesde").Value = vfdesde.Value
            .Recordset("FHasta").Value = vfhasta.Value
            .Recordset("Repeticiones").Value = 0
        Else
            .Recordset("Repeticiones").Value = Val(txtrcantidad.Text)
        End If
        
        .Recordset.Update
    End With
    
    Limpiar

If Err Then GrabarLog "cmdAceptar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Imprimir()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "    Prepare la Impresora    ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsComentarios
        If .State = 1 Then .Close
        
        .Source = bComentario.RecordSource
        
        If .State = 0 Then .Open
        .Close
        .Open
    End With
    
    With drComentario
        .Sections(2).Controls("sfecha").Caption = vfecha.Value
        .Sections(2).Controls("sreparto").Caption = cbocReparto.Text
        .Sections(2).Controls("srepartidor").Caption = cbocRepartidor.Text
        .Show
    End With


If Err Then GrabarLog "Imprimir", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()
On Error Resume Next
      
    vBandera = 0
    vModifica = False
    txtlocalidad.Text = ""
    cboReparto.Text = ""
    cboRepartidor.Text = ""
    txtComentario.Text = ""
    vfdesde.Value = Date
    vfhasta.Value = Date
    txtrcantidad.Text = ""
    txtCliente.SetFocus
    
If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Modificar()
On Error Resume Next

    With bComentario
        If Not (.Recordset.EOF = True) And Not (.Recordset.BOF = True) Then
        
        vCodigo = .Recordset("Codigo").Value
        txtCliente.Text = .Recordset("Nombre").Value
        txtlocalidad.Text = .Recordset("Localidad").Value
        cboReparto.Text = .Recordset("Reparto").Value
        cboRepartidor.Text = .Recordset("Repartidor").Value
        txtComentario.Text = .Recordset("Comentario").Value
        
        If .Recordset("Repeticiones").Value = 0 Then
            vfdesde.Value = .Recordset("Fdesde").Value
            vfhasta.Value = .Recordset("FHasta").Value
            txtrcantidad.Text = 0
        Else
             txtrcantidad.Text = .Recordset("Repeticiones").Value
        End If

        vBandera = .Recordset("Id").Value
        vModifica = True
        
        TabComentarios.Tab = 0
        
        End If
    
    End With
  

If Err Then GrabarLog "Modificar", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtCliente_Keypress(KeyAscii As Integer)
On Error Resume Next
        
        If KeyAscii = 13 Then
            If Not Trim(txtCliente.Text) = "" Then
                BuscarCliente
            End If
        End If
    
If Err Then GrabarLog "txtcliente_Keypress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtcComentario_KeyPress(KeyAscii As Integer)
On Error Resume Next
        
        If KeyAscii = 13 Then Filtrar
        
If Err Then GrabarLog "txtComentario_Filtro_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtComentario_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then txtrcantidad.SetFocus

If Err Then GrabarLog "txtComentario_Filtro_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtcNombre_KeyPress(KeyAscii As Integer)
On Error Resume Next
        
        If KeyAscii = 13 Then Filtrar
        
If Err Then GrabarLog "txtNombre_Filtro_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtrcantidad_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then cmdAcciones(0).SetFocus

If Err Then GrabarLog "txtrcantidad_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

