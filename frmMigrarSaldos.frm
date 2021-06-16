VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmMigrarSaldos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Migra Porcentaje de Ganancia a Clientes Especiales"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9945
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   9945
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir los Saldos"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   7680
      Width           =   9855
   End
   Begin VB.Frame fraSaldo 
      Caption         =   "Guardar en:"
      Height          =   975
      Left            =   8280
      TabIndex        =   5
      Top             =   0
      Width           =   1575
      Begin VB.OptionButton op_saldo 
         Caption         =   "S. Migrado"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   540
         Width           =   1335
      End
      Begin VB.OptionButton op_saldo 
         Caption         =   "S. Actual"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.CheckBox cmdSoloDiferencias 
      BackColor       =   &H80000013&
      Caption         =   "SOLO SALDOS DIFERENTES"
      Height          =   495
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7080
      Width           =   9855
   End
   Begin MSDataGridLib.DataGrid DgSaldos 
      Bindings        =   "frmMigrarSaldos.frx":0000
      Height          =   5295
      Left            =   0
      TabIndex        =   3
      Top             =   1680
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   9340
      _Version        =   393216
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
      ColumnCount     =   6
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
         DataField       =   "Saldo_Actual"
         Caption         =   "Saldo_Actual"
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
         DataField       =   "Saldo_Migrado"
         Caption         =   "Saldo_Migrado"
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
         BeginProperty Column00 
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   4110.236
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1454.74
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1470.047
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   0
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCalcularSaldos 
      Caption         =   "Calcular Saldo de Clientes"
      Height          =   855
      Left            =   4200
      TabIndex        =   2
      Top             =   120
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc bsaldos_clientes 
      Height          =   330
      Left            =   10095
      Top             =   0
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bsaldos_clientes"
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
   Begin MSAdodcLib.Adodc barticulos 
      Height          =   330
      Left            =   10095
      Top             =   720
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "barticulos"
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
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSAdodcLib.Adodc bfdetalle 
      Height          =   330
      Left            =   10095
      Top             =   2160
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bfdetalle"
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
   Begin VB.CommandButton cmdMigrarPorcentaje 
      Caption         =   "Migrar Porcentaje de Ganacia"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
   End
   Begin MSAdodcLib.Adodc bclientes 
      Height          =   330
      Left            =   10095
      Top             =   1800
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bclientes"
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
   Begin MSAdodcLib.Adodc bfacturas 
      Height          =   330
      Left            =   10095
      Top             =   1080
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bfacturas"
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
   Begin MSAdodcLib.Adodc bclientes_ganancia 
      Height          =   330
      Left            =   10095
      Top             =   1440
      Visible         =   0   'False
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bclientes_ganancia"
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
   Begin MSAdodcLib.Adodc bSaldos 
      Height          =   330
      Left            =   10095
      Top             =   360
      Width           =   3000
      _ExtentX        =   5292
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
      Caption         =   "bSaldos"
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
Attribute VB_Name = "frmMigrarSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vganancia As Double
Function Busca_Rubro(vcodigo_articulo As String) As String
On Error Resume Next
    
    'A partir de un Articulo, averiguo el RUBRO
    With barticulos
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from Articulos where Codigo = '" & vcodigo_articulo & "'"
        .Refresh
        
        If Not .Recordset.EOF = True Then 'Controlo que haya encontrado el articulo
            If Not IsNull(.Recordset("Rubro")) Then 'Controlo que el campo rubro no sea NULO
                Busca_Rubro = .Recordset("Rubro").Value
            Else
                Busca_Rubro = ""
            End If
        End If
    End With

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub cmdCalcularSaldos_Click()
On Error Resume Next
    
    BorrarBase "Saldos", pathDBMySQL
    
    With bclientes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from clientes where Pasivo = 'NO' order by Codigo_Num"
        .Refresh
        
        If .Recordset.EOF = True Then Exit Sub
        
        
        .Recordset.MoveFirst
        Barra.Value = 0
        Barra.Max = .Recordset.RecordCount
        
        Do Until .Recordset.EOF = True
            DoEvents
            Guardo_Saldo fsaldoclientes(.Recordset("Codigo").Value)
        
            .Recordset.MoveNext
            Barra.Value = Barra.Value + 1
        Loop

    End With

If Err Then GrabarLog "cmdCalcularSaldos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdImprimir_Click()
On Error Resume Next
    
    With bSaldos
        If (.Recordset.EOF = True) Or (.Recordset.BOF = True) Then Exit Sub
        
        .Recordset.MoveFirst
        
        Do Until .Recordset.EOF = True
            .Recordset("Diferencia").Value = Val(Format(.Recordset("Saldo_Actual").Value - .Recordset("Saldo_Migrado").Value, "######0.00"))
            .Recordset.MoveNext
        Loop
    
    End With
    
    Unload Mantenimiento
    Load Mantenimiento
    
    With Mantenimiento.rsArreglos
    
        If Not .State = 1 Then
            .Open
            .Close
            .Open
        Else
            .Close
            .Open
        End If
    
        If MsgBox("¿¿¿ Imprimir SOLAMENTE saldos con DIFENCIAS ???", vbYesNo, "Mensaje ...") = vbYes Then
            
            .filter = "[Diferencia] > 0"
        
        Else
            
            .filter = "[id] > 0"
        
        End If

    End With
    
    With drArreglos
        .Show
    End With


If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdMigrarPorcentaje_Click()
On Error Resume Next

    With bfdetalle
        'Filtro todos los detalles que tienen Ganancia en 0 y no es el articulo 99999
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Fdetalle WHERE (Ganancia = 0) and not (Codigo = '99999')"
        .Refresh
        
        Barra.Value = 0
        Barra.Max = .Recordset.RecordCount
        '
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
        
        Do Until .Recordset.EOF = True
            DoEvents
            'Controlo la Factura del Remito
            Controlo_Factura (.Recordset("Remito").Value), Busca_Rubro(.Recordset("Codigo").Value), (.Recordset("Codigo").Value)
            
            .Recordset("Ganancia") = vganancia
            .Recordset("Sueldo") = (vganancia / 100) * .Recordset("Cantidad").Value * .Recordset("Precio")
            
            .Recordset.MoveNext
            Barra.Value = Barra.Value + 1
        
        Loop
    
    End With
    
If Err Then GrabarLog "cmdMigrarPorcentaje_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdSoloDiferencias_Click()
On Error Resume Next
        
    
    With bSaldos
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            
        If cmdSoloDiferencias.Value = 1 Then
            .RecordSource = "SELECT * FROM saldos WHERE [saldo_actual]-[saldo_migrado] > 0"
            cmdSoloDiferencias.Caption = "TODOS LOS SALDOS"
        Else
            .RecordSource = "SELECT * FROM saldos"
            cmdSoloDiferencias.Caption = "SOLO SALDOS DIFERENTES"
        End If
        .Refresh
    End With

If Err Then GrabarLog "cmdSoloDiferencias_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Controlo_Factura(vnumero_Remito As Long, vcodigo_rubro As String, vcodigo_articulo As String)
On Error Resume Next
    
    With bfacturas
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from Factura where remito = " & vnumero_Remito
        .Refresh
        
        If Not .Recordset.EOF = True Then
            
            Controlo_Ganancia (.Recordset("Codigo").Value), (vcodigo_rubro), (vcodigo_articulo)
            
        Else
            
            'PANIC "No Tiene Factura OJO"
        
        End If
        
    End With

If Err Then GrabarLog "Controlo_Factura", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Controlo_Ganancia(vcodigo_cliente As String, vcodigo_rubro As String, vcodigo_articulo As String)
On Error Resume Next

    With bclientes_ganancia
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from CliRep where Cod_Cliente = '" & vcodigo_cliente & "' and COd_rubro = '" + vcodigo_rubro + "'"
        .Refresh
        
        If Not .Recordset.EOF = True Then
                    
            vganancia = .Recordset("Porcentaje").Value
            
        Else
            With barticulos
                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "Select * from articulos where codigo = '" & vcodigo_articulo & "'"
                .Refresh
                
                If Not .Recordset.EOF = True Then vganancia = .Recordset("Ganancia").Value
                
            End With
        
        End If
         

    End With

If Err Then GrabarLog "Controlo_Ganancia", Err.Number & " " & Err.Description, Me.Name
End Sub
Function fsaldoclientes(vCodigo As String) As Double 'Calcula el saldo a cualquier FECHA (CONSULTA Saldos_clientes)
    On Error Resume Next

    With bsaldos_clientes
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT Max(cuentascorrientes.Fecha) AS Fecha, cuentascorrientes.Codigo, Sum(cuentascorrientes.Debito) AS SumaDeDebito, Sum(cuentascorrientes.Credito) AS SumaDeCredito, ([SumaDeDebito]-[SumaDeCredito]) AS Saldo FROM cuentascorrientes INNER JOIN Clientes ON cuentascorrientes.Codigo = Clientes.Codigo WHERE (((cuentascorrientes.Noimputar) <> True)) GROUP BY cuentascorrientes.Codigo HAVING (((cuentascorrientes.Codigo) = '" + vCodigo + "'))"
        .Refresh
        
        If Not .Recordset.EOF = True Then
            fsaldoclientes = Val(Format(.Recordset("Saldo").Value, "#######0.00"))
        Else
            fsaldoclientes = 0
        End If

    End With

    If Err Then GrabarLog "fsaldoclientes ", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub Guardo_Saldo(vsaldo As Double)
On Error Resume Next

    With bSaldos
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM saldos ORDER BY id ASC"
        .Refresh
        
        .Recordset.AddNew
        
        
        .Recordset("Codigo").Value = bclientes.Recordset("Codigo").Value
        .Recordset("Nombre").Value = bclientes.Recordset("Nombre").Value
        .Recordset("Localidad").Value = bclientes.Recordset("Localidad").Value
        
        If op_saldo(0).Value = True Then
            .Recordset("Saldo_Actual").Value = vsaldo
        Else
            .Recordset("Saldo_Migrado").Value = vsaldo
        End If
        
        .Recordset.Update
    End With


If Err Then GrabarLog "Guardo_Saldo", Err.Number & " " & Err.Description, Me.Name
End Sub

