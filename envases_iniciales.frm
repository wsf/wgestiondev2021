VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEnvasesIniciales 
   Caption         =   "Mantenimiento de Envases Iniciales"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   12930
   Begin VB.CommandButton cmdMostrar 
      Caption         =   "Mostrar Todo"
      Height          =   375
      Left            =   3210
      TabIndex        =   14
      Top             =   4170
      Width           =   1485
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar el Cliente y Artículo Seleccionado"
      Height          =   435
      Left            =   8940
      TabIndex        =   13
      Top             =   5100
      Width           =   3735
   End
   Begin MSAdodcLib.Adodc bdevol 
      Height          =   330
      Left            =   240
      Top             =   3800
      Width           =   12555
      _ExtentX        =   22146
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
      Caption         =   "bdevol"
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
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   1740
      TabIndex        =   12
      Top             =   4170
      Width           =   1485
   End
   Begin VB.CommandButton cmdBorrar 
      Caption         =   "Borrar"
      Height          =   375
      Left            =   270
      TabIndex        =   9
      Top             =   4170
      Width           =   1485
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   270
      TabIndex        =   3
      Top             =   6120
      Width           =   1485
   End
   Begin VB.TextBox venvases 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2940
      TabIndex        =   2
      Top             =   5700
      Width           =   1605
   End
   Begin VB.TextBox varticulo 
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   5310
      Width           =   6495
   End
   Begin VB.TextBox vcliente 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   4980
      Width           =   6495
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "envases_iniciales.frx":0000
      Height          =   3675
      Left            =   240
      TabIndex        =   4
      Top             =   60
      Width           =   12555
      _ExtentX        =   22146
      _ExtentY        =   6482
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
      ColumnCount     =   12
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "SumaDeCantidad"
         Caption         =   "SumaDeCantidad"
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
         DataField       =   "SumaDeImpuesto1"
         Caption         =   "SumaDeImpuesto1"
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
      BeginProperty Column04 
         DataField       =   "Clientes.Codigo"
         Caption         =   "Clientes.Codigo"
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
         DataField       =   "ÚltimoDeFecha"
         Caption         =   "ÚltimoDeFecha"
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
         DataField       =   "ÚltimoDeNombre"
         Caption         =   "ÚltimoDeNombre"
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
         DataField       =   "Mora"
         Caption         =   "Mora"
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
         DataField       =   "SumaDeenvases"
         Caption         =   "SumaDeenvases"
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
         DataField       =   "ÚltimoDerepartidor"
         Caption         =   "ÚltimoDerepartidor"
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
         DataField       =   "ÚltimoDeEnvase"
         Caption         =   "ÚltimoDeEnvase"
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
         DataField       =   "MoraEmpleado"
         Caption         =   "MoraEmpleado"
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
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   599.811
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   2264.882
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   989.858
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3089.764
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   1244.976
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc barticulo 
      Height          =   405
      Left            =   3240
      Top             =   6120
      Visible         =   0   'False
      Width           =   3555
      _ExtentX        =   6271
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
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   405
      Left            =   3240
      Top             =   6120
      Visible         =   0   'False
      Width           =   3555
      _ExtentX        =   6271
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
   Begin MSAdodcLib.Adodc bart_cli_envases 
      Height          =   405
      Left            =   3240
      Top             =   6120
      Visible         =   0   'False
      Width           =   3555
      _ExtentX        =   6271
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
      Caption         =   "bart_cli_envases"
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
   Begin VB.Label varticulo_codigo 
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   7740
      TabIndex        =   11
      Top             =   5310
      Width           =   1005
   End
   Begin VB.Label vcliente_codigo 
      BackColor       =   &H00404040&
      ForeColor       =   &H00C0FFFF&
      Height          =   315
      Left            =   7740
      TabIndex        =   10
      Top             =   4980
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "> Cantidad de Envaces prestados:"
      Height          =   255
      Left            =   300
      TabIndex        =   8
      Top             =   5730
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "> Artículo:"
      Height          =   255
      Left            =   300
      TabIndex        =   7
      Top             =   5340
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "> Cliente:"
      Height          =   255
      Left            =   300
      TabIndex        =   6
      Top             =   4980
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Mantenimiento de datos:"
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
      Height          =   255
      Left            =   300
      TabIndex        =   5
      Top             =   4620
      Width           =   12495
   End
End
Attribute VB_Name = "frmEnvasesIniciales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
On Error Resume Next

    With bart_cli_envases
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "select * from art_cli_envases where articulo = '" + Trim(varticulo_codigo.Caption) + "' and cliente = '" + Trim(vcliente_codigo.Caption) + "'"
        .Refresh

        If .Recordset.EOF Then .Recordset.AddNew

        .Recordset("articulo") = varticulo_codigo.Caption
        .Recordset("cliente") = vcliente_codigo.Caption
        .Recordset("envases") = venvases.Text
    
        .Recordset.Update

        vcliente.SetFocus
    
        If vConfigGral.vIncluyeContabilidad = True Then
            With frmAsientosAlta
                .Show
                .ZOrder (0)
                .txtCuentaVieneDe.Text = Me.Caption
            End With
        End If
    End With

If Err Then GrabarLog "cmdAceptar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdBorrar_Click()
On Error Resume Next

    Borrar bdevol.object, True

If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdBuscar_Click()
On Error Resume Next

    With bdevol
        .RecordSource = "select * from Devol where DetalleCodigo = '" + Trim(varticulo_codigo.Caption) + "' and Codigo = '" + Trim(vcliente_codigo.Caption) + "'"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveLast
    End With
    
If Err Then GrabarLog "cmdBuscar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdMostrar_Click()
On Error Resume Next
    
    With bdevol
        .RecordSource = "select * from devol"
        .Refresh
        
        If Not .Recordset.EOF = True Then .Recordset.MoveLast
    End With
        
If Err Then GrabarLog "cmdBuscar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    With bdevol
        .ConnectionString = pathDBMySQL
        .RecordSource = "Devol"
        .Refresh
    End With
    
    
    With Me
        .Width = 13050
        .Height = 7245
    End With

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub varticulo_Change()
On Error Resume Next

    If varticulo.Text = "" Then varticulo_codigo.Caption = ""


If Err Then GrabarLog "varticulo_Change", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub varticulo_keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        With barticulo
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from articulos where (codigo = '" + varticulo.Text + "')"
            .Refresh

            If .Recordset.EOF = True Then
                .RecordSource = "Select * from articulos where (descrip like '%" + varticulo.Text + "%')"
                .Refresh
                If Not .Recordset.EOF = True Then
                    varticulo = .Recordset("descrip")
                    varticulo_codigo.Caption = .Recordset("codigo")
                End If
            Else
                varticulo = .Recordset("descrip")
                varticulo_codigo.Caption = .Recordset("codigo")
            End If

            cmdBuscar_Click
        End With
    End If

If Err Then GrabarLog "varticulo_keypress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vcliente_Change()
On Error Resume Next

    If vcliente.Text = "" Then vcliente_codigo.Caption = ""

If Err Then GrabarLog "vcliente_Change", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub vcliente_Keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        With bCliente
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "select * from clientes where codigo = '" + vcliente.Text + "' or nombre like '%" + vcliente.Text + "%'"
            .Refresh

            If Not .Recordset.EOF = True Then
                vcliente = .Recordset("nombre")
                vcliente_codigo.Caption = .Recordset("codigo")
            End If
        
        End With
    
    End If
    
If Err Then GrabarLog "vcliente_Keypress", Err.Number & " " & Err.Description, Me.Name
End Sub
