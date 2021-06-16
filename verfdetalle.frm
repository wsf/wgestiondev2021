VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form verfdetalle 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Detalle de Factura"
   ClientHeight    =   4125
   ClientLeft      =   10260
   ClientTop       =   1485
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkpagados 
      Caption         =   "Articulos Pagados"
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   11
      Top             =   3840
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.CheckBox chksinpagar 
      Caption         =   "Articulos sin pagar"
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   3495
   End
   Begin VB.ListBox lstdetalles 
      Height          =   1410
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   2040
      Width           =   3735
   End
   Begin VB.Frame fradetalle 
      Height          =   15
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   5055
   End
   Begin MSAdodcLib.Adodc bfacturas 
      Height          =   330
      Left            =   2040
      Top             =   4440
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin MSAdodcLib.Adodc bfdetalle 
      Height          =   330
      Left            =   120
      Top             =   4440
      Visible         =   0   'False
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin VB.Label lblrepartidor 
      AutoSize        =   -1  'True
      DataField       =   "Repartidor"
      DataSource      =   "bfacturas"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   9
      Top             =   1440
      Width           =   45
   End
   Begin VB.Label lblcondicion 
      AutoSize        =   -1  'True
      DataField       =   "Cventa"
      DataSource      =   "bfacturas"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      DataField       =   "Fecha"
      DataSource      =   "bfacturas"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "> Detalle :"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label lblrepartidor 
      AutoSize        =   -1  'True
      Caption         =   "> Repartidor :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label lblcondicion 
      AutoSize        =   -1  'True
      Caption         =   "> Condición :"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   930
   End
   Begin VB.Label lblfecha 
      AutoSize        =   -1  'True
      Caption         =   "> Fecha :"
      Height          =   195
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   675
   End
   Begin VB.Label lbltitulos 
      AutoSize        =   -1  'True
      Caption         =   "Detalles de la Factura"
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
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   1890
   End
End
Attribute VB_Name = "verfdetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vfremito As Long

Private Sub Form_Load()
    On Error Resume Next

    With bfacturas
        .ConnectionString = pathDB
        .RecordSource = "select * from factura where remito = " & Val(vfremito)
        '.RecordSource = "select * from Factura"
        .Refresh
    End With

    With bfdetalle
        .ConnectionString = pathDB
        .RecordSource = "Select * from fdetalle where remito = " & Val(vfremito)
        .Refresh

        If Not .Recordset.RecordCount = 0 Then .Recordset.MoveFirst
        lstdetalles.Clear
        Dim i As Integer
        i = 0

        Do Until .Recordset.EOF = True
            
            If .Recordset("Pagado") = "SI" Then
                lstdetalles.AddItem .Recordset("cantidad") & "  -  " & .Recordset("detalle") & "  -  $" & .Recordset("Precio")
                lstdetalles.Selected(i) = True
            Else
                lstdetalles.AddItem .Recordset("cantidad") & "  -  " & .Recordset("detalle") & "  -  $" & .Recordset("Precio")
                lstdetalles.Selected(i) = False
            End If

            .Recordset.MoveNext
            i = i + 1
        Loop

    End With

    If Err Then FeLu.graba_log "Form_load", Err.Number & " " & Err.Description, Me.caption
End Sub

