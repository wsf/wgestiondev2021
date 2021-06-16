VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmClienteRepartidor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Asignación de porcentaje de ganancia Cliente/Repartidor"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   10545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   10545
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.ProgressBar barra 
      Height          =   225
      Left            =   90
      TabIndex        =   20
      Top             =   4470
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   397
      _Version        =   327682
      Appearance      =   0
   End
   Begin MSAdodcLib.Adodc bclirep2 
      Height          =   330
      Left            =   3000
      Top             =   5280
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
      Caption         =   "bclirep2"
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   2025
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3572
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Altas / Busquedas"
      TabPicture(0)   =   "Cliente_Repartidor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "vrepartidor"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "vcliente"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdFiltrar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "vporcentaje"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "vrubro"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Asociar Rubro a los Rubros Generales"
      TabPicture(1)   =   "Cliente_Repartidor.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label5"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(3)=   "Text1"
      Tab(1).Control(4)=   "vrubro2"
      Tab(1).Control(5)=   "cmdCargaNuevoRubro"
      Tab(1).Control(6)=   "cmdBorrarRubros"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdBorrarRubros 
         Caption         =   "Borrar Rubros Generales"
         Height          =   375
         Left            =   -67680
         TabIndex        =   21
         Top             =   1470
         Width           =   2535
      End
      Begin VB.CommandButton cmdCargaNuevoRubro 
         Caption         =   "Cargar nuevo rubro y %"
         Height          =   375
         Left            =   -67680
         TabIndex        =   18
         Top             =   1050
         Width           =   2535
      End
      Begin VB.ComboBox vrubro2 
         Height          =   315
         ItemData        =   "Cliente_Repartidor.frx":0038
         Left            =   -72780
         List            =   "Cliente_Repartidor.frx":003F
         TabIndex        =   16
         Text            =   "General"
         Top             =   1140
         Width           =   4335
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   -72780
         TabIndex        =   14
         Top             =   1530
         Width           =   1125
      End
      Begin VB.ComboBox vrubro 
         Height          =   315
         ItemData        =   "Cliente_Repartidor.frx":004C
         Left            =   2010
         List            =   "Cliente_Repartidor.frx":0053
         TabIndex        =   12
         Text            =   "General"
         Top             =   1110
         Width           =   4335
      End
      Begin VB.TextBox vporcentaje 
         Height          =   285
         Left            =   2010
         TabIndex        =   10
         Top             =   1590
         Width           =   1125
      End
      Begin VB.Frame Frame4 
         Height          =   555
         Left            =   7920
         TabIndex        =   7
         Top             =   1320
         Width           =   2055
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir"
            Height          =   345
            Left            =   1020
            TabIndex        =   8
            Top             =   150
            Width           =   975
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "Agregar"
            Height          =   345
            Left            =   60
            TabIndex        =   9
            Top             =   150
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdFiltrar 
         Caption         =   "Filtrar "
         Height          =   585
         Left            =   8790
         MaskColor       =   &H8000000F&
         Picture         =   "Cliente_Repartidor.frx":0060
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Filtrar datos con valores ingresados."
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   855
      End
      Begin VB.TextBox vcliente 
         Height          =   285
         Left            =   2010
         TabIndex        =   2
         Top             =   450
         Width           =   5355
      End
      Begin VB.TextBox vrepartidor 
         Height          =   285
         Left            =   2010
         TabIndex        =   1
         Top             =   780
         Width           =   5385
      End
      Begin VB.Label Label7 
         Caption         =   "Ingresar RUBRO y PORCENTAJE de ganancia para aquellos Clientes/Empleados que tengan rubro GENERAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -74640
         TabIndex        =   17
         Top             =   540
         Width           =   9585
      End
      Begin VB.Label Label6 
         Caption         =   "> % de Ganancia :"
         Height          =   195
         Left            =   -74550
         TabIndex        =   15
         Top             =   1560
         Width           =   1395
      End
      Begin VB.Label Label5 
         Caption         =   "> Ingresar Rubro:"
         Height          =   225
         Left            =   -74550
         TabIndex        =   13
         Top             =   1170
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "> % de Ganancia :"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1620
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "> Ingresar Cliente :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "> Ingresar Repartidor :"
         Height          =   225
         Left            =   240
         TabIndex        =   4
         Top             =   810
         Width           =   1995
      End
      Begin VB.Label Label4 
         Caption         =   "> Ingresar Rubro:"
         Height          =   225
         Left            =   240
         TabIndex        =   3
         Top             =   1140
         Width           =   1995
      End
   End
   Begin MSAdodcLib.Adodc brubros 
      Height          =   330
      Left            =   480
      Top             =   5640
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
      Caption         =   "brubros"
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
   Begin MSAdodcLib.Adodc bempleados 
      Height          =   330
      Left            =   480
      Top             =   5280
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
      Caption         =   "bempleados"
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
   Begin MSAdodcLib.Adodc bclirep 
      Height          =   330
      Left            =   3000
      Top             =   4920
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
      Caption         =   "bclirep"
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
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   330
      Left            =   480
      Top             =   4920
      Visible         =   0   'False
      Width           =   2500
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Cliente_Repartidor.frx":0162
      Height          =   2085
      Left            =   90
      TabIndex        =   19
      Top             =   2340
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   3678
      _Version        =   393216
      AllowUpdate     =   -1  'True
      Appearance      =   0
      BackColor       =   -2147483624
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Cod_Empleado"
         Caption         =   "Cod_Empleado"
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
         DataField       =   "Empleado"
         Caption         =   "Empleado"
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
         DataField       =   "Cod_Cliente"
         Caption         =   "Cod_Cliente"
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
         DataField       =   "Cliente"
         Caption         =   "Cliente"
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
         DataField       =   "Cod_Rubro"
         Caption         =   "Cod_Rubro"
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
         DataField       =   "Rubro"
         Caption         =   "Rubro"
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
         DataField       =   "Porcentaje"
         Caption         =   "Porcentaje"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2789.858
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2610.142
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   2894.74
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1184.882
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClienteRepartidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vcod_cliente, sql, vcod_empleado, vcodigo_rubro, vcodigo_rubro2 As String

Private Sub cmdAgregar_Click()
    On Error Resume Next
    
    With bclirep
        .Recordset.AddNew
        
        .Recordset("cod_cliente") = vcod_cliente
        .Recordset("cod_empleado") = vcod_empleado
        .Recordset("cliente") = vcliente
        .Recordset("empleado") = vrepartidor
        .Recordset("porcentaje") = vporcentaje
        .Recordset("cod_rubro") = vcodigo_rubro
        .Recordset("rubro") = vrubro.Text
    
        .Recordset.Update

    End With
    Limpiar
    vcliente.SetFocus
    
    If Err Then
        MsgBox "Los datos NO fueron guardados !", vbInformation, "Error..."
        GrabarLog "cmdAgregar_Click", Err.Number & " " & Err.Description, Me.Name
    End If

End Sub

Private Sub cmdBorrarRubros_Click()
On Error Resume Next

    If MsgBox("Confirma el borrado de los Rubros 'Generales'", vbYesNo + vbInformation, "Mensaje ...") = vbYes Then

        With bclirep2
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from clirep where rubro = 'General'"
            .Refresh
        
            Do Until .Recordset.EOF = True
                .Recordset.Delete
                .Recordset.MoveNext
            Loop

            bclirep.Refresh
            
        End With
    
    End If
    
If Err Then GrabarLog "cmdBorrarRubros_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdCargaNuevoRubro_Click()
Dim i As Integer
On Error Resume Next
    
    With bclirep2
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "select * from clirep where rubro = 'General'"
        .Refresh

        If .Recordset.EOF = True Then Exit Sub
        
        Barra.Max = .Recordset.RecordCount
        Barra.Value = 0
    

        Do Until .Recordset.EOF
            
            bclirep.Recordset.AddNew

            For i = 0 To 6
                bclirep.Recordset(i).Value = .Recordset(i).Value
            Next

            bclirep.Recordset("rubro") = vrubro2.Text
            bclirep.Recordset("cod_rubro") = vcodigo_rubro2
            bclirep.Recordset("porcentaje") = Val(vporcentaje)
           
            bclirep.Recordset.Update
        
            Barra.Value = Barra.Value + 1
        
            .Recordset.MoveNext
        
        Loop
    End With

If Err Then GrabarLog "cmdCargaNuevoRubro_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdFiltrar_Click()
On Error Resume Next
    
    sql = ""

    If Not vrepartidor.Text = "" Then sql = sql + " and Empleado like '%" + Trim(vrepartidor) + "%'"
    If Not vcliente.Text = "" Then sql = sql + " and Cliente like '%" + Trim(vcliente) + "%'"
    If Not vrubro.Text = "" Then sql = sql + " and rubro like '%" + Trim(vrubro) + "%'"
    If Not vporcentaje.Text = "" Then sql = sql + " and porcentaje = " + Trim(vporcentaje) + ""


    With bclirep
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "select * from clirep where 1=1" + sql
        .Refresh
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
    End With

If Err Then GrabarLog "cmdFiltrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdImprimir_Click()
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    With Mantenimiento.rscliente_repartidor

        If .State = 1 Then
            .Close
            .Open
        Else
            .Open
            .Close
            .Open
        End If

    
        If Not vrepartidor.Text = "" Then sql = sql + "and Empleado like '%" + Trim(vrepartidor) + "%' "
        If Not vcliente.Text = "" Then sql = sql + "and Cliente like '%" + Trim(vcliente) + "%' "
        

        .filter = "Cliente <> Null " + sql
    
    End With
    
    MsgBox "     Prepare la impresora     ", vbInformation, "Mensaje ..."
    
    With drclirep
        .Show
    End With
    
If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next


    With bclirep
        .ConnectionString = pathDBMySQL
        .RecordSource = "clirep"
        .Refresh
    End With

    With bclirep2
        .ConnectionString = pathDBMySQL
        .RecordSource = "clirep"
        .Refresh
    End With

   
    With Me
        .Height = 5265
        .Width = 10665
        .Top = 100
        .Left = 300
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()
On Error Resume Next

    vcliente.Text = ""
    vrepartidor.Text = ""
    vrubro.Text = ""
    vporcentaje.Text = ""

If Err Then GrabarLog "limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vcliente_Keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        With bcliente
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from clientes where codigo = '" + Trim(vcliente.Text) + "' or nombre like '%" + Trim(vcliente.Text) + "%'"
            .Refresh
    
            If Not .Recordset.EOF = True Then
        
                vcliente.Text = .Recordset("nombre")
                vcod_cliente = .Recordset("codigo")
                vrepartidor.SetFocus
        
            Else
        
                MsgBox "El Cliente no fue encontrado.", vbInformation, "Mensaje ..."
        
            End If

        End With
    End If

If Err Then GrabarLog "vcliente_Keypress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vporcentaje_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then cmdAgregar.SetFocus
    
If Err Then GrabarLog "vporcentaje_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vrepartidor_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        With bempleados
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from empleados where codigo = '" + Trim(vrepartidor.Text) + "' or nombre like '%" + Trim(vrepartidor.Text) + "%'"
            .Refresh
    
            If Not .Recordset.EOF = True Then
                vrepartidor.Text = .Recordset("nombre")
                vcod_empleado = .Recordset("codigo")
                vrubro.SetFocus
            Else
                MsgBox "El Empleado no fue encontrado.", vbInformation, "Mensaje ..."
            End If
        
        End With

    End If
If Err Then GrabarLog "vrepartidor_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vrubro2_KeyPress(KeyAscii As Integer)
On Error Resume Next


    If KeyAscii = 13 Then
        
        With brubros
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from rubros where codigo = '" + Trim(vrubro2) + "' or nombre like '%" + Trim(vrubro2) + "%'"
            .Refresh
    
            If Not .Recordset.EOF = True Then
                vrubro2.Text = .Recordset("nombre")
                vcodigo_rubro2 = .Recordset("codigo")
                vporcentaje.SetFocus
            Else
            
                MsgBox "El rubro no fue encontrado !.", vbInformation, "Mensaje..."
            End If
  
        End With
    
    End If

If Err Then GrabarLog "vrubro2_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vrubro_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        With brubros
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from rubros where codigo = '" + Trim(vrubro) + "' or nombre like '%" + Trim(vrubro) + "%'"
            .Refresh
    
            If Not .Recordset.EOF Then
                vrubro.Text = .Recordset("nombre")
                vcodigo_rubro = .Recordset("codigo")
                vporcentaje.SetFocus
            Else
                MsgBox "El rubro no fue encontrado !.", vbInformation, "Mensaje..."
            End If
  
        End With
    End If
    
If Err Then GrabarLog "vrubro_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
