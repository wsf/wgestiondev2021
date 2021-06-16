VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmClientesRepartos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Repartos"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton csalir 
      Caption         =   "Salir"
      Height          =   495
      Left            =   2505
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton cimprimir 
      Caption         =   "Imprimir"
      Height          =   495
      Left            =   1725
      Picture         =   "clirepartos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Imprimir"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   795
   End
   Begin VB.CommandButton cborrar 
      Appearance      =   0  'Flat
      Caption         =   "Borrar"
      Height          =   495
      Left            =   915
      Picture         =   "clirepartos.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Borrar datos"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   825
   End
   Begin VB.CommandButton cguardar 
      Appearance      =   0  'Flat
      Caption         =   "Guardar"
      Height          =   495
      Left            =   105
      Picture         =   "clirepartos.frx":0204
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guardar datos"
      Top             =   5160
      UseMaskColor    =   -1  'True
      Width           =   825
   End
   Begin MSAdodcLib.Adodc bclirepartos 
      Height          =   330
      Left            =   3630
      Top             =   5310
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "bclirepartos"
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
      Bindings        =   "clirepartos.frx":0306
      Height          =   5085
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8969
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483624
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
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
      ColumnCount     =   3
      BeginProperty Column00 
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
      BeginProperty Column01 
         DataField       =   "nreparto"
         Caption         =   "nreparto"
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
         DataField       =   "descrip"
         Caption         =   "descrip"
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
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1739.906
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   4215.118
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmClientesRepartos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cborrar_Click()
    Borrar bclirepartos, True
End Sub

Private Sub cguardar_Click()
    bclirepartos.Recordset.Update
End Sub

Private Sub cimprimir_Click()
    drclirepartos.Show
End Sub

Private Sub csalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    With bclirepartos
        .ConnectionString = pathDBMySQL
        .RecordSource = "clireparto"
        .Refresh
    End With

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
