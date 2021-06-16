VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmEstadisticaCliente 
   Caption         =   "Estadística sobre Clientes"
   ClientHeight    =   7875
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7875
   ScaleWidth      =   12870
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "Area"
      Height          =   315
      Left            =   3810
      TabIndex        =   22
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "3D"
      Height          =   315
      Left            =   2880
      TabIndex        =   21
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Barra"
      Height          =   315
      Left            =   1950
      TabIndex        =   20
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Linea"
      Height          =   315
      Left            =   1020
      TabIndex        =   19
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Torta"
      Height          =   315
      Left            =   90
      TabIndex        =   18
      Top             =   6720
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1995
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   7155
      Begin VB.TextBox varticulo 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   2790
         TabIndex        =   0
         Top             =   480
         Width           =   4335
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Volumen de venta"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   330
         TabIndex        =   10
         Top             =   1260
         Width           =   1905
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Volumen de ganancia"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   330
         TabIndex        =   9
         Top             =   1590
         Width           =   1905
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Relación Costo/Beneficio"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2430
         TabIndex        =   8
         Top             =   1230
         Width           =   2145
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Opción"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2430
         TabIndex        =   7
         Top             =   1590
         Width           =   2145
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Opción"
         Height          =   255
         Left            =   4770
         TabIndex        =   6
         Top             =   1230
         Width           =   1095
      End
      Begin VB.CheckBox Check6 
         Caption         =   "Opción"
         Height          =   255
         Left            =   4800
         TabIndex        =   5
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Ejecutar"
         Height          =   405
         Left            =   5910
         TabIndex        =   4
         Top             =   1320
         Width           =   915
      End
      Begin VB.CheckBox o 
         Caption         =   "Anular Fecha "
         Height          =   255
         Left            =   780
         TabIndex        =   3
         Top             =   900
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   285
         Left            =   1440
         TabIndex        =   11
         Top             =   180
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   73662465
         CurrentDate     =   38029
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   285
         Left            =   1440
         TabIndex        =   12
         Top             =   510
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   73662465
         CurrentDate     =   38029
      End
      Begin VB.Label Label7 
         Caption         =   "> Fecha Hasta :"
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   570
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "> Fecha Desde :"
         Height          =   225
         Left            =   0
         TabIndex        =   14
         Top             =   210
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Caption         =   "Nombre del Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2790
         TabIndex        =   13
         Top             =   240
         Width           =   4335
      End
   End
   Begin MSAdodcLib.Adodc barticulo 
      Height          =   330
      Left            =   6720
      Top             =   6720
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
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
   Begin MSAdodcLib.Adodc bgrafico 
      Height          =   330
      Left            =   4800
      Top             =   6720
      Visible         =   0   'False
      Width           =   1965
      _ExtentX        =   3466
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
      Caption         =   "bgrafico"
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
   Begin MSChart20Lib.MSChart g 
      Bindings        =   "EstCliente.frx":0000
      Height          =   4485
      Left            =   90
      OleObjectBlob   =   "EstCliente.frx":001C
      TabIndex        =   1
      Top             =   2160
      Width           =   11625
   End
   Begin RichTextLib.RichTextBox display 
      Height          =   1755
      Left            =   7350
      TabIndex        =   16
      Top             =   300
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   3096
      _Version        =   393217
      TextRTF         =   $"EstCliente.frx":289B
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Caption         =   "Speach "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   7350
      TabIndex        =   17
      Top             =   60
      Width           =   4335
   End
End
Attribute VB_Name = "frmEstadisticaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vcodigoart As String

Private Sub buscaart()
    barticulo.RecordSource = "select * from clientes where (nombre = '" + varticulo + "') or (codigo = '" + varticulo + "')"
    barticulo.Refresh

    If barticulo.Recordset.EOF Then
        frmBuscarCliente.o = 11
        frmBuscarCliente.txtClientes.Text = varticulo.Text
        frmBuscarCliente.txtClientes_keypress (13)
        frmBuscarCliente.txtClientes.SetFocus
        frmBuscarCliente.Show

    Else
        varticulo = barticulo.Recordset(1).Value
        vcodigoart = barticulo.Recordset(0).Value
    End If

End Sub

Private Sub Command1_Click()

    On Error Resume Next

    Dim i As Integer

    If o.Value = 1 Then
        bgrafico.RecordSource = "select factura.fecha, fdetalle.fecha, factura.remito, factura.total, fdetalle.remito, fdetalle.cantidad from factura,fdetalle where factura.codigo =  '" + vcodigoart + "' and factura.remito = fdetalle.remito" ' group by fdetalle.remito"
        bgrafico.Refresh
    Else
        '   bgrafico.RecordSource = "select * from factura where fecha >= '" & strfechaMySQl(fdesde) + "' and fecha <= '" & strfechaMySQl(fhasta) + "' codigo =  '" + vcodigoart + "'"
        bgrafico.Refresh
    End If

    bgrafico.Recordset.MoveFirst

    g.ColumnCount = 2
    g.RowCount = bgrafico.Recordset.RecordCount

    i = 0
    display.Text = "Fecha        Cantidad           Total" + Chr(13)
    display.Text = display.Text + "----------------------------------------------------" + Chr(13)
    
    Do Until bgrafico.Recordset.EOF
        i = i + 1
        g.Row = i
        '----------------------------------------
        g.Column = 1
        g.Data = bgrafico.Recordset("total")
        g.RowLabel = Str(bgrafico.Recordset("fecha"))
        display.Text = display.Text + Str(bgrafico.Recordset("factura.fecha")) + "                  " + Format(bgrafico.Recordset("total"), "$ ########.00") + Chr(13)
        '---------------------------------------
    
        '----------------------------------------
        g.Column = 2
        g.Data = bgrafico.Recordset("cantidad")
        display.Text = display.Text + Str(bgrafico.Recordset("fdetalle.fecha")) + "     " + Str(bgrafico.Recordset("cantidad")) + Chr(13)
        '---------------------------------------
    
        bgrafico.Recordset.MoveNext
    Loop

    If Err Then Exit Sub

End Sub

Private Sub Command2_Click()
    g.chartType = 14
End Sub

Private Sub Command3_Click()
    g.chartType = 3
End Sub

Private Sub Command4_Click()
    g.chartType = 9
End Sub

Private Sub Command5_Click()
    g.chartType = 2
End Sub

Private Sub Command6_Click()
    g.chartType = 4
End Sub

Private Sub Form_Load()

    With barticulo
        .ConnectionString = pathDBMySQL
        .RecordSource = "Clientes"
        .Refresh
    End With

    With bgrafico
        .ConnectionString = pathDBMySQL
        .RecordSource = "fdetalle"
        .Refresh
    End With

End Sub

Private Sub o_Click()

    If fdesde.Enabled = True Then fdesde.Enabled = False
    If fdesde.Enabled = False Then fdesde.Enabled = True
    If fhasta.Enabled = True Then fhasta.Enabled = False
    If fhasta.Enabled = False Then fhasta.Enabled = True
End Sub

Public Sub varticulo_keypress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        buscaart
    End If

End Sub

