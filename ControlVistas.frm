VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmControlVistas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de Control de Vistas"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleWidth      =   13590
   Begin VB.CommandButton cmdImprime 
      Caption         =   "Imprime Listado"
      Height          =   495
      Left            =   11880
      TabIndex        =   3
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "Ejecutar Listado"
      Height          =   495
      Left            =   10320
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
   End
   Begin ComctlLib.ProgressBar barra 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5280
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin MSAdodcLib.Adodc bsaldos_clientes 
      Height          =   330
      Left            =   6480
      Top             =   7320
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
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   330
      Left            =   3960
      Top             =   6600
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
   Begin MSAdodcLib.Adodc bvista 
      Height          =   330
      Left            =   3960
      Top             =   6960
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
      Caption         =   "bvista"
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
   Begin MSDataGridLib.DataGrid DgVistas 
      Bindings        =   "ControlVistas.frx":0000
      Height          =   5055
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   8916
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
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "Id"
         Caption         =   "Id"
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
         DataField       =   "Codigo"
         Caption         =   "Codigo"
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
      BeginProperty Column02 
         DataField       =   "Cliente"
         Caption         =   "Cliente"
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
      BeginProperty Column03 
         DataField       =   "Ficha"
         Caption         =   "Ficha"
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
      BeginProperty Column04 
         DataField       =   "Mes"
         Caption         =   "Mes"
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
      BeginProperty Column05 
         DataField       =   "Factura"
         Caption         =   "Factura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "DifFicha"
         Caption         =   "DifFicha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """$"" #,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "DifFactura"
         Caption         =   "DifFactura"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """$"" #,##0.00"
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
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3000.189
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
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc bfacturas 
      Height          =   330
      Left            =   6480
      Top             =   6960
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
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   6480
      Top             =   6600
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
End
Attribute VB_Name = "frmControlVistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vDetener As Boolean

Private Sub cmdEjecutar_Click()
Dim vmensual, vfactura, vficha As Double
On Error Resume Next

    If cmdEjecutar.Caption = " Detener " Then
        vDetener = True
    Else
        vDetener = False
    End If
    With bcliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * From Clientes"
        .Refresh
        
        If .Recordset.EOF = True Then Exit Sub
        
        barra.Max = .Recordset.RecordCount
        barra.Value = 0
    
    End With
    
    cmdEjecutar.Caption = " Detener "
    BorrarBase "vista", pathDBMySQL

    With bvista
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Vista"
        .Refresh


        Do Until bcliente.Recordset.EOF = True
            DoEvents
            
            If vDetener = True Then
                cmdEjecutar.Caption = " Ejecutar Listado "
                Exit Sub
            End If
            
            vmensual = mensual(Trim(bcliente.Recordset("Codigo").Value))
            vfactura = facturas(Trim(bcliente.Recordset("codigo").Value))
            vficha = Ficha(Trim(bcliente.Recordset("codigo").Value))
    
    
            If Not ((vmensual = vfactura) And (vmensual = vficha) And (vfactura = vficha)) Then
    
                .Recordset.AddNew
        
                .Recordset("Codigo") = bcliente.Recordset("codigo")
                .Recordset("Cliente") = bcliente.Recordset("nombre")
                .Recordset("Ficha") = vficha
                .Recordset("Mes") = vmensual
                .Recordset("Factura") = vfactura
                .Recordset("DifFicha") = vmensual - vficha
                .Recordset("DifFactura") = vmensual - vfactura
        
                .Recordset.Update
        
            End If
    
            bcliente.Recordset.MoveNext
            barra.Value = barra.Value + 1
        Loop
    
        cmdEjecutar.Caption = " Ejecutar Listado "
    End With

If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdImprime_Click()
On Error Resume Next
    
        Imprime

If Err Then GrabarLog "cmdImprime_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Function facturas(vcli As String) As Double
Dim vsaldo, vfsaldo As Double
On Error Resume Next

    vsaldo = 0

    With bfacturas
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "Select * from Factura_ctacte where Codigo = '" + Trim(vcli) + "' order by Fecha ASC, remito ASC, ID ASC"
        .Refresh
        
        
        If Not .Recordset.RecordCount = 0 Then
        
            Do Until .Recordset.EOF = True
    
                If Not Val(Format(.Recordset("credito"), "########0.00")) > 0 Then
            
                    vsaldo = vsaldo + Val(Format(.Recordset("diferencia"), "#######0.00"))
                    vfsaldo = vsaldo
        
                End If
    
                .Recordset.MoveNext
            Loop
            
            facturas = vsaldo
        Else
        
            facturas = 0
                        
        End If
    End With
    
If Err.Number Then GrabarLog "facturas : " & vcli, Err.Number & " " & Err.Description, Me.Name
End Function
Function Ficha(vcli As String) As Double
Dim vsaldo As Double

On Error Resume Next
 
    With bccliente
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "select * from cuentascorrientes where codigo = '" + vcli + "' and (Noimputar = true or debito > 0) order by fecha ASC, id ASC"
        .Refresh

        
        If Not .Recordset.RecordCount = 0 Then
            .Recordset.MoveFirst
            
            Do Until .Recordset.EOF = True
                DoEvents
                vsaldo = vsaldo - Val(Format(.Recordset("Credito"), "#######0.00")) + Val(Format(.Recordset("debito"), "#######0.00"))
                
                .Recordset.MoveNext
        
            Loop '
        
            Ficha = Val(Format(vsaldo, "#######0.00"))
        
        Else
        
            Ficha = 0
        
        End If
    
    End With
    
If Err Then GrabarLog "Ficha", Err.Number & " " & Err.Description, Me.Name
End Function

Private Sub Form_Load()
On Error Resume Next
    
    With bvista
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM vista"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveFirst
    
    End With

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Imprime()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento

    With Mantenimiento.rsVistas
        
        If .State = 1 Then
            
            .Close
            .Open
        
        Else
        
            .Open
            .Close
            .Open
        
        End If
    
        .filter = "(IdVista > 0)"
        .Sort = "Codigo ASC, IdVista ASC"
    
    End With
    
    With drVistas
        .Show
    End With
    
If Err Then GrabarLog "Imprime", Err.Number & " " & Err.Description, Me.Name
End Sub
Function mensual(vcli As String) As Double
On Error Resume Next

        With bsaldos_clientes
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "Saldos_Clientes"
            .Refresh
        
            If Not .Recordset.EOF = True Then .Recordset.MoveFirst
                
            .Recordset.Find ("codigo = '" + vcli + "'")

            If Not .Recordset.EOF = True Then
                mensual = Val(Format(bsaldos_clientes.Recordset("expr1"), "#######0.00"))
            Else
                mensual = 0
            End If
        
        End With
                
If Err Then GrabarLog "mensual : " & vcli, Err.Number & " " & Err.Description, Me.Name
End Function
