VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Monitor 
   Caption         =   "Monitor del sistema - Control de errores y warnning"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   13710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12555
   ScaleWidth      =   17160
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Debitos en CtaCte Borrados"
      Height          =   3795
      Left            =   990
      TabIndex        =   6
      Top             =   5970
      Width           =   11595
      Begin VB.CommandButton Command1 
         Caption         =   "Ver Errores"
         Height          =   375
         Left            =   10350
         TabIndex        =   13
         Top             =   2910
         Width           =   1005
      End
      Begin ComctlLib.ProgressBar b 
         Height          =   255
         Left            =   210
         TabIndex        =   12
         Top             =   3390
         Width           =   11265
         _ExtentX        =   19870
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSAdodcLib.Adodc bfdetalle_debito 
         Height          =   495
         Left            =   5880
         Top             =   2760
         Visible         =   0   'False
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   873
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
         Caption         =   "bdifFacFd"
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
      Begin MSAdodcLib.Adodc bctacte_debito 
         Height          =   495
         Left            =   870
         Top             =   2760
         Visible         =   0   'False
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   873
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
         Caption         =   "bdifFacFd"
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
      Begin VB.ListBox l 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2910
         Left            =   150
         TabIndex        =   11
         Top             =   390
         Width           =   11235
      End
   End
   Begin VB.Frame Frame11 
      Height          =   5715
      Left            =   6840
      TabIndex        =   3
      Top             =   0
      Width           =   6945
      Begin VB.CommandButton Command2 
         Caption         =   "actualizar"
         Height          =   375
         Left            =   5520
         TabIndex        =   14
         Top             =   5280
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc berroVista 
         Height          =   495
         Left            =   960
         Top             =   4680
         Visible         =   0   'False
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   873
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
         Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
         OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Vbprog\La Surgente\Datos\Felu.mdb;Persist Security Info=False"
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "errorvistas"
         Caption         =   "berroVista"
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
      Begin MSDataGridLib.DataGrid DataGrid9 
         Bindings        =   "Monitor.frx":0000
         Height          =   4005
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7064
         _Version        =   393216
         BackColor       =   16761024
         HeadLines       =   1
         RowHeight       =   15
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label vvmal 
         Alignment       =   2  'Center
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   615
         Left            =   2160
         TabIndex        =   10
         Top             =   4920
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Cantidad de Registros:"
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   5130
         Width           =   2175
      End
      Begin VB.Label Label2 
         Caption         =   "Clientes con diferencias entre saldo Vista Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   435
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame Frame10 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6585
      Begin VB.CommandButton Command3 
         Caption         =   "actuaizar"
         Height          =   375
         Left            =   5160
         TabIndex        =   15
         Top             =   5280
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc bDifFactFDetalle 
         Height          =   495
         Left            =   780
         Top             =   4050
         Visible         =   0   'False
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   873
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
         Caption         =   "bdifFacFd"
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
      Begin MSDataGridLib.DataGrid DataGrid8 
         Bindings        =   "Monitor.frx":0019
         Height          =   4035
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   7117
         _Version        =   393216
         BackColor       =   8454143
         HeadLines       =   1
         RowHeight       =   15
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
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Label vdborrados 
         Alignment       =   2  'Center
         Caption         =   "Label4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   615
         Left            =   1800
         TabIndex        =   8
         Top             =   4920
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Cantidad de Registros:"
         Height          =   285
         Left            =   150
         TabIndex        =   7
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Verifique las artículos facturados en las siguientes fechas correspondiente a cada CtaCte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   435
         Left            =   180
         TabIndex        =   2
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
bebitos_borrados
End Sub

Private Sub Command2_Click()
vvmal.Caption = Trim(Me.berroVista.Recordset.RecordCount)
End Sub

Private Sub Command3_Click()
vdborrados.Caption = Trim(bDifFactFDetalle.Recordset.RecordCount)
End Sub

Private Sub Form_Load()
On Error Resume Next

With bDifFactFDetalle
        .ConnectionString = pathDB
        .RecordSource = "DifFactFDetalle"
        .Refresh
vdborrados.Caption = Trim(.Recordset.RecordCount)

End With

 With bErrorVista
        .ConnectionString = pathDB
        .RecordSource = "ErrorVistas"
        .Refresh
vvmal.Caption = Trim(.Recordset.RecordCount)
End With
    
With bctacte_debito
        .ConnectionString = pathDB
        .RecordSource = "ctacte_debito"
        .Refresh
End With

 With bfdetalle_debito
        .ConnectionString = pathDB
        .RecordSource = "fdetalle_debito"
        .Refresh
End With
    
    


End Sub

Private Sub bebitos_borrados()
On Error Resume Next
vdborrados.Caption = Trim(bDifFactFDetalle.Recordset.RecordCount)
bfdetalle_debito.Refresh

b.Max = bfdetalle_debito.Recordset.RecordCount
b.Value = 0

Do Until bfdetalle_debito.Recordset.EOF
        bctacte_debito.RecordSource = "select * from ctacte_debito where remito = " + Trim(bfdetalle_debito.Recordset("remito"))
        bctacte_debito.Refresh
        
        b.Value = b.Value + 1
        If bctacte_debito.Recordset.EOF Then
            l.AddItem ("> " + bfdetalle_debito.Recordset("codigo") + "  " + bfdetalle_debito.Recordset("nombre") + "  " + Trim(bfdetalle_debito.Recordset("fecha")))
        End If
                
        bfdetalle_debito.Recordset.MoveNext
Loop


If Err < 0 Then Exit Sub
End Sub
