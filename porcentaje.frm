VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmPorcentaje 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento de Porcentaje en Articulos para libreta"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdimprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8160
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   4080
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc bporcentaje 
      Height          =   330
      Left            =   0
      Top             =   4800
      Width           =   9135
      _ExtentX        =   16113
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
      Caption         =   "bporcentaje"
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
   Begin MSDataGridLib.DataGrid dgporcentaje 
      Bindings        =   "porcentaje.frx":0000
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   5530
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
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "codigo"
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
         DataField       =   "descrip"
         Caption         =   "Articulo"
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
         DataField       =   "pventa1"
         Caption         =   "Precio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "porcentaje"
         Caption         =   "% de aumento"
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
      BeginProperty Column04 
         DataField       =   "pventat"
         Caption         =   "Precio Final"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   1
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   2
         Size            =   2
         BeginProperty Column00 
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3569.953
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1590.236
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1319.811
         EndProperty
      EndProperty
   End
   Begin VB.Frame fralinea 
      Height          =   15
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   9615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimiento de Porcentaje en Articulos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9105
   End
End
Attribute VB_Name = "frmPorcentaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub cmdImprimir_Click()
On Error Resume Next
    With drporcentaje
        Mantenimiento.rsporcentaje.Sort = "Codigonum ASC"
        Unload Mantenimiento
        
        Load Mantenimiento
        .Show
    End With
If Err Then GrabarLog "cmdimprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub dgporcentaje_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
    With bporcentaje
        Select Case ColIndex
            Case 0, 1 'Codigo de Articulo - Descrip de Articulo
                .Recordset.Update
            Case 2, 3 'Pventa1 - 'Porcentaje de Aumento
                .Recordset("Pventat") = Val(Format(.Recordset("pventa1"), "#######0.00")) + (Val(Format(.Recordset("pventa1"), "######0.00")) * Val(Format(.Recordset("porcentaje"), "######0.00"))) / 100
                .Recordset.Update
            Case 4 'Total con aumento
                .Recordset("Pventa1") = .Recordset("pventat") / Val("1." & .Recordset("porcentaje"))
                .Recordset.Update
        End Select
    End With
If Err Then GrabarLog "dgporcentaje_AfterColEdit: " & ColIndex, Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next
    With bporcentaje
        .ConnectionString = pathDBMySQL
        .RecordSource = "Select codigo, descrip, pventa1, porcentaje, pventat  from articulos"
        .Refresh
    End With
    Height = 5025
    Width = 9450
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub

