VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEmpleados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento de Empleados"
   ClientHeight    =   6855
   ClientLeft      =   60
   ClientTop       =   225
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab TabEmpleados 
      Height          =   6585
      Left            =   -60
      TabIndex        =   14
      ToolTipText     =   "Configuración de parámetros. Porcentaje de ganancia"
      Top             =   -330
      Width           =   11205
      _ExtentX        =   19764
      _ExtentY        =   11615
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ingresar datos"
      TabPicture(0)   =   "Empleados.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lblAlta(6)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblAlta(9)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAlta(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAlta(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblAlta(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblAlta(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblAlta(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblAlta(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblAlta(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblAlta(7)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblTitulo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Shape2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblCodigoExistente(0)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblCodigoExistente(1)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "v1(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "v1(9)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "iva"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "v1(8)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "v1(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "v1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "v1(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "v1(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "v1(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "bot"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "v1(7)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdMover(0)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdBuscar(0)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Frame3"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "cmdMover(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "cmdNuevo"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdPlanilla"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "Forma Planilla"
      TabPicture(1)   =   "Empleados.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "dgEmpleados"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "fraBusqueda"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "fraOrden"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "fraImprimir"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.Frame fraImprimir 
         Height          =   705
         Left            =   8400
         TabIndex        =   41
         Top             =   5400
         Width           =   2625
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "Borrar"
            Height          =   535
            Index           =   1
            Left            =   1680
            Picture         =   "Empleados.frx":0038
            Style           =   1  'Graphical
            TabIndex        =   42
            ToolTipText     =   "Borrar datos"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   855
         End
         Begin VB.CommandButton cmdModificar 
            Caption         =   "Modificar"
            Height          =   535
            Left            =   870
            Picture         =   "Empleados.frx":05C2
            Style           =   1  'Graphical
            TabIndex        =   43
            ToolTipText     =   "Click Sobre algun Empleado lo Modifica"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   825
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir"
            Height          =   535
            Left            =   60
            Picture         =   "Empleados.frx":6E14
            Style           =   1  'Graphical
            TabIndex        =   44
            ToolTipText     =   "Generar reporte para imprimir"
            Top             =   120
            UseMaskColor    =   -1  'True
            Width           =   825
         End
      End
      Begin VB.Frame fraOrden 
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
         Height          =   705
         Left            =   6120
         TabIndex        =   38
         Top             =   5400
         Width           =   2205
         Begin VB.OptionButton opOrden 
            Caption         =   "Nombre"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   40
            Top             =   330
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opOrden 
            Caption         =   "Código"
            Height          =   195
            Index           =   1
            Left            =   1170
            TabIndex        =   39
            Top             =   330
            Width           =   945
         End
      End
      Begin VB.CommandButton cmdPlanilla 
         Caption         =   "Formato Planilla"
         Height          =   495
         Left            =   -71460
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ir a planilla de empleados"
         Top             =   4860
         Width           =   1245
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   495
         Left            =   -72270
         Picture         =   "Empleados.frx":739E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Nuevo Empleado"
         Top             =   4860
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.CommandButton cmdMover 
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
         Index           =   1
         Left            =   -68700
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4830
         Width           =   525
      End
      Begin VB.Frame fraBusqueda 
         Height          =   975
         Left            =   180
         TabIndex        =   32
         Top             =   5250
         Width           =   5865
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   615
            Index           =   1
            Left            =   4800
            Picture         =   "Empleados.frx":74A0
            Style           =   1  'Graphical
            TabIndex        =   45
            ToolTipText     =   "Generar reporte para imprimir"
            Top             =   240
            UseMaskColor    =   -1  'True
            Width           =   945
         End
         Begin VB.TextBox txtLocalidad 
            Height          =   315
            Left            =   1050
            TabIndex        =   35
            ToolTipText     =   "Presionar enter para ejectutar la consulta. Filtra por código y por descripción simultaneamente"
            Top             =   540
            Width           =   3525
         End
         Begin VB.TextBox txtEmpleado 
            Height          =   315
            Left            =   1050
            TabIndex        =   34
            ToolTipText     =   "Presionar enter para ejectutar la consulta. Filtra por código y por descripción simultaneamente"
            Top             =   180
            Width           =   3525
         End
         Begin VB.Label lblBusqueda 
            Alignment       =   1  'Right Justify
            Caption         =   "Localidad :"
            Height          =   195
            Index           =   1
            Left            =   30
            TabIndex        =   37
            Top             =   600
            Width           =   1000
         End
         Begin VB.Label lblBusqueda 
            Alignment       =   1  'Right Justify
            Caption         =   "Empleado :"
            Height          =   195
            Index           =   0
            Left            =   30
            TabIndex        =   36
            Top             =   270
            Width           =   1000
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
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Enabled         =   0   'False
         Height          =   495
         Index           =   0
         Left            =   -73050
         Picture         =   "Empleados.frx":79D2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Buscar Empleado"
         Top             =   4860
         UseMaskColor    =   -1  'True
         Width           =   795
      End
      Begin VB.CommandButton cmdMover 
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
         Index           =   0
         Left            =   -69210
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4830
         Width           =   525
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   7
         Left            =   -72850
         TabIndex        =   7
         Top             =   3480
         Width           =   1125
      End
      Begin VB.Frame bot 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   -74610
         TabIndex        =   16
         Top             =   4830
         Visible         =   0   'False
         Width           =   1575
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "Borrar"
            Height          =   495
            Index           =   0
            Left            =   780
            Picture         =   "Empleados.frx":7AD4
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
            Picture         =   "Empleados.frx":7BD6
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
         Left            =   -72850
         TabIndex        =   1
         Top             =   1290
         Width           =   4635
      End
      Begin VB.TextBox v1 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   -72850
         TabIndex        =   0
         Top             =   930
         Width           =   1470
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   2
         Left            =   -72850
         TabIndex        =   2
         Top             =   1665
         Width           =   1875
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   3
         Left            =   -72850
         TabIndex        =   3
         Top             =   2025
         Width           =   1875
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   4
         Left            =   -72850
         TabIndex        =   4
         Top             =   2385
         Width           =   1875
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   8
         Left            =   -72850
         TabIndex        =   8
         Top             =   3840
         Width           =   2655
      End
      Begin VB.ComboBox iva 
         Height          =   315
         ItemData        =   "Empleados.frx":7CD8
         Left            =   -72850
         List            =   "Empleados.frx":7CE5
         TabIndex        =   5
         Text            =   "Responsable Inscripto"
         Top             =   2750
         Width           =   2625
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   9
         Left            =   -72850
         TabIndex        =   9
         Top             =   4200
         Width           =   2655
      End
      Begin VB.TextBox v1 
         Height          =   315
         Index           =   6
         Left            =   -72850
         TabIndex        =   6
         Top             =   3120
         Width           =   2655
      End
      Begin MSDataGridLib.DataGrid dgEmpleados 
         Bindings        =   "Empleados.frx":7D1F
         Height          =   4785
         Left            =   180
         TabIndex        =   15
         Top             =   420
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   8440
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         HeadLines       =   2
         RowHeight       =   15
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label lblCodigoExistente 
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   -69840
         TabIndex        =   47
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblCodigoExistente 
         Caption         =   "Codigo Usado Por :"
         Height          =   195
         Index           =   0
         Left            =   -71280
         TabIndex        =   46
         Top             =   960
         Width           =   1380
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   615
         Left            =   -69450
         Top             =   4860
         Width           =   1275
      End
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         Caption         =   "Ingreso de Datos de Empleados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Crédito máximo :"
         Height          =   285
         Index           =   7
         Left            =   -74970
         TabIndex        =   26
         Top             =   3510
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Tipo de I.V.A. :"
         Height          =   255
         Index           =   5
         Left            =   -74970
         TabIndex        =   25
         Top             =   2805
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Nombre de Empleado :"
         Height          =   345
         Index           =   1
         Left            =   -74970
         TabIndex        =   24
         Top             =   1335
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Código  del Empleado:"
         Height          =   405
         Index           =   0
         Left            =   -74970
         TabIndex        =   23
         Top             =   960
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Dirección :"
         Height          =   225
         Index           =   2
         Left            =   -74970
         TabIndex        =   22
         Top             =   1695
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Teléfono :"
         Height          =   405
         Index           =   4
         Left            =   -74970
         TabIndex        =   21
         Top             =   2445
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Localidad :"
         Height          =   285
         Index           =   3
         Left            =   -74970
         TabIndex        =   20
         Top             =   2070
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Responsable  :"
         Height          =   315
         Index           =   8
         Left            =   -74970
         TabIndex        =   19
         Top             =   3870
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> Ing. Bruto :"
         Height          =   345
         Index           =   9
         Left            =   -74970
         TabIndex        =   18
         Top             =   4200
         Width           =   2000
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         Caption         =   "> C.U.I.T.  :"
         Height          =   315
         Index           =   6
         Left            =   -74970
         TabIndex        =   17
         Top             =   3150
         Width           =   2000
      End
   End
   Begin MSAdodcLib.Adodc bEmpleados 
      Height          =   405
      Left            =   0
      Top             =   6480
      Visible         =   0   'False
      Width           =   11055
      _ExtentX        =   19500
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
      Caption         =   "bEmpleados"
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
Attribute VB_Name = "frmEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Buscar()
On Error Resume Next
    
    Dim vSQL As String, vOrdenado As String

    If opOrden(0).Value = True Then
        vOrdenado = "Nombre"
    Else
        vOrdenado = "Codigo"
    End If

    vSQL = ""

    If Not txtEmpleado = "" Then vSQL = vSQL + " AND (nombre LIKE '%" + Trim(txtEmpleado) + "%' OR codigo = '" + Trim(txtEmpleado) + "')"
    If Not txtLocalidad = "" Then vSQL = vSQL + " AND (Localidad LIKE '%" + Trim(txtLocalidad) + "%')"

    With bempleados
        .RecordSource = "SELECT * FROM empleados WHERE 1=1 " + vSQL + " ORDER BY " & vOrdenado
        .Refresh
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
    End With

    FormatoGrilla
    
If Err Then GrabarLog "Buscar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla()
On Error Resume Next
    
        With dgEmpleados
            .HeadLines = 2
            
            .Columns(0).width = 0
            .Columns(1).width = 750
            .Columns(2).width = 2500
            .Columns(3).width = 1500
            .Columns(4).width = 1500
            .Columns(5).width = 1500
            .Columns(6).width = 1500
            .Columns(7).width = 1000
            .Columns(8).width = 0
            .Columns(9).width = 0
            .Columns(10).width = 0
            .Columns(11).width = 0
        
        End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdBuscar_Click(Index As Integer)
    On Error Resume Next

    If Index = 0 Then
        frmBuscarEmpleado.o = 0
        frmBuscarEmpleado.Show
    Else
        Buscar
    End If

    If Err Then GrabarLog "cmdBuscar_Click", Err.Number & "-" & Err.Description, Me.Name
End Sub
Private Sub cmdModificar_Click()
On Error Resume Next

    Limpiar

    With bempleados
        If (.Recordset.EOF = True) Or (.Recordset.BOF = True) Then
            MsgBox "Debe seleccionar un Empleado para poder Modificarlo!!!", vbExclamation, "Mensaje ..."
            Exit Sub
        End If
    End With
    
    MostrarRegistro
    
    TabEmpleados.Tab = 0
    v1(0).Tag = "Modificando"
    v1(0).Enabled = False
    v1(1).SetFocus
    

If Err Then GrabarLog "cmdModificar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdNuevo_Click()
On Error Resume Next
    
    Limpiar

If Err Then GrabarLog "cmdNuevo_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdGuardar_Click()
    On Error Resume Next
    
    Dim rsEmpleado As New ADODB.Recordset, sqlEmpleado As String
    
    sqlEmpleado = "SELECT * FROM Empleados WHERE (codigo = '" + Trim(v1(0).Text) + "')"
    
    With rsEmpleado
        .CursorLocation = adUseClient
    
        Call .Open(sqlEmpleado, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .EOF = True Then
            .AddNew
        Else
            If Not MsgBox("El Empleado que esta ingresando ya existe, desea modificar sus datos ?", vbYesNo + vbInformation, "Mensaje ...") = vbYes Then
                Exit Sub
            End If
        End If
    
        .Fields("codigo") = v1(0).Text
        .Fields("Nombre").Value = v1(1).Text
        .Fields("Direccion").Value = v1(2).Text
        .Fields("Localidad").Value = v1(3).Text
        .Fields("Telefono").Value = v1(4).Text
        .Fields("Iva").Value = iva.Text
        .Fields("Cuit").Value = v1(6).Text
        .Fields("Credito").Value = Val(v1(7).Text)
        .Fields("Responsable").Value = v1(8).Text
        .Fields("IBrutos").Value = v1(9).Text
        
        .Update
    End With
    
    sqlEmpleado = ""
    
    If rsEmpleado.State = 1 Then
        rsEmpleado.Close
        Set rsEmpleado = Nothing
    End If
    
    Limpiar

If Err Then GrabarLog "cmdGuardar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdBorrar_Click(Index As Integer)
    On Error Resume Next
    
    If Index = 0 Then
        MsgBox "No Habilitado"
    Else
        Call Borrar(bempleados, True)
    End If

    If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdPlanilla_Click()
On Error Resume Next

    TabEmpleados.Tab = 1

If Err Then GrabarLog "cmdPlanilla_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdMover_Click(Index As Integer)
    On Error Resume Next

    If Index = 0 Then
        'Va para atras
    Else
        'Va para adelante
    End If
    
    MostrarRegistro
    
If Err Then GrabarLog "cmdMover_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdImprimir_Click()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora !!!!", vbInformation, "Mensaje ..."

    With Mantenimiento.rsempleados
        If .State = 1 Then .Close
        
        .Source = bempleados.RecordSource
    
        If .State = 0 Then .Open
        .Close
        .Open
        
        If .EOF = True Then Exit Sub
    End With
    
    With drEmpleados
        .Show
    End With

If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next
        
    With Me
        .Show
        .TabEmpleados.Tab = 0
    End With
    
    With bempleados
        .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Empleados ORDER BY Nombre ASC"
        .Refresh
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()
    On Error Resume Next
    Dim i As Integer

    For i = 0 To 9
        If Not i = 5 Then v1(i) = ""
    Next

    lblCodigoExistente(1).Caption = ""
    v1(0).Tag = ""
    v1(0).Enabled = True
    v1(0).SetFocus

    If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub MostrarRegistro()
    On Error Resume Next
    
    With bempleados
        v1(0).Text = EsNulo(.Recordset("Codigo").Value)
        v1(1).Text = EsNulo(.Recordset("Nombre").Value)
        v1(2).Text = EsNulo(.Recordset("Direccion").Value)
        v1(3).Text = EsNulo(.Recordset("Localidad").Value)
        v1(4).Text = EsNulo(.Recordset("Telefono").Value)
        iva.Text = EsNulo(.Recordset("Iva").Value)
        v1(6).Text = EsNulo(.Recordset("Cuit").Value)
        v1(7).Text = EsNulo(.Recordset("Credito").Value)
        v1(8).Text = EsNulo(.Recordset("Responsable").Value)
        v1(9).Text = EsNulo(.Recordset("IBrutos").Value)
    End With

    If Err Then GrabarLog "MostrarRegistro", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub TabEmpleados_Click(PreviousTab As Integer)
On Error Resume Next

    Select Case TabEmpleados.Tab

        Case 0
            With Me
                .Top = 100
                .width = 7155
                .Left = 2000
                .height = 5745
                .v1(0).SetFocus
            End With

        Case 1
            With Me
                .Top = 100
                .width = 11180
                .Left = 500
                .height = 6990
                .txtEmpleado.SetFocus
            End With
        Case 2
    
    End Select

If Err Then GrabarLog "TabEmpleados_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtEmpleado_Change()
On Error Resume Next

    Buscar

If Err Then GrabarLog "txtEmpleado_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtLocalidad_Change()
On Error Resume Next

    Buscar

If Err Then GrabarLog "txtLocalidad_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub v1_Change(Index As Integer)
On Error Resume Next

    If Not Trim(v1(0).Tag) = "Modificando" Then
        If Not v1(0) = "" Then
            lblCodigoExistente(1).Caption = ControlarExistente(v1(0).Text)
            If (lblCodigoExistente(1).Caption) = "" Then
                bot.Visible = True
            Else
                bot.Visible = False
            End If
        Else
            lblCodigoExistente(1).Caption = ""
            bot.Visible = False
        End If
    Else
        lblCodigoExistente(1).Caption = "Modificando..."
        bot.Visible = True
    End If
    
If Err Then GrabarLog "v1_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub v1_KeyPress(Index As Integer, _
                        KeyAscii As Integer)

    If KeyAscii = 13 Then
    
        If Index >= 9 Then
            If bot.Visible = True Then cmdGuardar.SetFocus
        Else

            If Index = 4 Then Index = 5
            v1(Index + 1).SetFocus
        End If

    End If

End Sub
Private Function ControlarExistente(vCodigoNuevo As String) As String
On Error Resume Next

    ControlarExistente = TraerDato("Empleados", "Codigo = '" & vCodigoNuevo & "'", "Nombre")

If Err Then GrabarLog "ControlarExistente", Err.Number & " " & Err.Description, Me.Name
End Function
