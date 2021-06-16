VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmSaldos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Saldos Totales"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   11010
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   11010
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc bcheques 
      Height          =   330
      Left            =   8160
      Top             =   5880
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Caption         =   "bcheques"
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
   Begin MSAdodcLib.Adodc bcaja 
      Height          =   330
      Left            =   3960
      Top             =   5880
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
      Caption         =   "bcaja"
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
      Left            =   2040
      Top             =   5880
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
   Begin MSChart20Lib.MSChart g 
      Height          =   3105
      Left            =   240
      OleObjectBlob   =   "frmSaldos.frx":0000
      TabIndex        =   9
      Top             =   2640
      Width           =   10605
   End
   Begin VB.PictureBox r 
      Height          =   480
      Left            =   9480
      ScaleHeight     =   420
      ScaleWidth      =   1140
      TabIndex        =   16
      Top             =   5160
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Ejecutar"
      Height          =   375
      Left            =   7290
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1290
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Visualizar Reporte"
      Height          =   375
      Left            =   180
      MaskColor       =   &H8000000F&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5880
      UseMaskColor    =   -1  'True
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Caption         =   "Saldos :"
      ForeColor       =   &H00000080&
      Height          =   2595
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   10785
      Begin VB.Frame Frame3 
         Caption         =   "Saldo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   825
         Left            =   6540
         TabIndex        =   25
         Top             =   1620
         Width           =   4125
         Begin VB.Label saldo 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H007EE9FC&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   150
            TabIndex        =   26
            Top             =   270
            Width           =   3855
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   7140
         TabIndex        =   20
         Top             =   120
         Width           =   3015
         Begin MSComCtl2.DTPicker fdesde 
            Height          =   315
            Left            =   1410
            TabIndex        =   21
            Top             =   180
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   58851329
            CurrentDate     =   38028
         End
         Begin MSComCtl2.DTPicker fhasta 
            Height          =   315
            Left            =   1410
            TabIndex        =   22
            Top             =   630
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   58851329
            CurrentDate     =   38028
         End
         Begin VB.Label Label11 
            Caption         =   "> Fecha Hasta :"
            Height          =   225
            Left            =   90
            TabIndex        =   24
            Top             =   660
            Width           =   1635
         End
         Begin VB.Label Label9 
            Caption         =   "> Fecha Desde :"
            Height          =   225
            Left            =   60
            TabIndex        =   23
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command8 
         Height          =   285
         Left            =   5820
         MaskColor       =   &H8000000F&
         Picture         =   "frmSaldos.frx":30CA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Inspeccionar"
         Top             =   2010
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton Command7 
         Height          =   285
         Left            =   5820
         MaskColor       =   &H8000000F&
         Picture         =   "frmSaldos.frx":31CC
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Inspeccionar"
         Top             =   1290
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton Command6 
         Height          =   285
         Left            =   5820
         MaskColor       =   &H8000000F&
         Picture         =   "frmSaldos.frx":32CE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Inspeccionar"
         Top             =   1650
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton Command5 
         Height          =   285
         Left            =   5820
         MaskColor       =   &H8000000F&
         Picture         =   "frmSaldos.frx":33D0
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Inspeccionar"
         Top             =   900
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton Command4 
         Height          =   285
         Left            =   5820
         MaskColor       =   &H8000000F&
         Picture         =   "frmSaldos.frx":34D2
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Inspeccionar"
         Top             =   510
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.Label Label12 
         Caption         =   "> Cheques Recibidos :"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   690
         TabIndex        =   19
         Top             =   2040
         Width           =   1785
      End
      Begin VB.Label srecibido 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "WST_Engl"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2580
         TabIndex        =   18
         Top             =   2040
         Width           =   3195
      End
      Begin VB.Line Line5 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   2
         X1              =   180
         X2              =   570
         Y1              =   2190
         Y2              =   2190
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF00FF&
         BorderWidth     =   2
         X1              =   180
         X2              =   570
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   180
         X2              =   570
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Line Line2 
         BorderColor     =   &H0000FF00&
         BorderWidth     =   2
         X1              =   180
         X2              =   570
         Y1              =   990
         Y2              =   990
      End
      Begin VB.Line Line1 
         BorderColor     =   &H000000FF&
         BorderWidth     =   2
         X1              =   180
         X2              =   600
         Y1              =   630
         Y2              =   630
      End
      Begin VB.Label semitido 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "WST_Engl"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2580
         TabIndex        =   8
         Top             =   1710
         Width           =   3165
      End
      Begin VB.Label scaja 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "WST_Engl"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2580
         TabIndex        =   7
         Top             =   1320
         Width           =   3165
      End
      Begin VB.Label sproveedor 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "WST_Engl"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2580
         TabIndex        =   6
         Top             =   900
         Width           =   3165
      End
      Begin VB.Label sclientes 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "WST_Engl"
            Size            =   9
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2580
         TabIndex        =   5
         Top             =   510
         Width           =   3165
      End
      Begin VB.Label Label4 
         Caption         =   ">Cheques emitidos :"
         ForeColor       =   &H00404040&
         Height          =   225
         Left            =   690
         TabIndex        =   4
         Top             =   1680
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "> Caja :"
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   690
         TabIndex        =   3
         Top             =   1290
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "> Cta. Cte.  Proveedores :"
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   660
         TabIndex        =   2
         Top             =   900
         Width           =   1965
      End
      Begin VB.Label Label3 
         Caption         =   "> Cta.Cte. Clientes :"
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   690
         TabIndex        =   1
         Top             =   540
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   330
      Left            =   6120
      Top             =   5880
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
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
End
Attribute VB_Name = "frmSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub calsaldocaja()
    On Error Resume Next
    Dim i As Long
    i = 0
    g.Column = 3

    Dim vsaldo As Double

    bcaja.Recordset.MoveFirst

    Do Until bcaja.Recordset.EOF
        vsaldo = vsaldo + bcaja.Recordset("deposito") - bcaja.Recordset("retiro")
    
        i = i + 1
        g.Row = i
        g.Data = vsaldo
    
        bcaja.Recordset.MoveNext
    Loop

    bcaja.Recordset.MoveLast
    scaja.Caption = Str(vsaldo)

    If Err Then Exit Sub
End Sub

Private Sub calsaldocheque()
    On Error Resume Next
    Dim i, j As Long
    i = 0
    j = 0

    Dim vemitido, vrecibido As Double

    bcheques.Refresh
    bcheques.Recordset.MoveFirst

    Do Until bcheques.Recordset.EOF
    
        If bcheques.Recordset("cp") = "p" Then
            vemitido = vemitido + bcheques.Recordset("monto")
            g.Column = 5
            j = j + 1
            g.Row = j
            g.Data = vemitido
        Else
            vrecibido = vrecibido + bcheques.Recordset("monto")
            g.Column = 4
            i = i + 1
            g.Row = i
            g.Data = vrecibido
        End If
   
        bcheques.Recordset.MoveNext
    Loop

    semitido.Caption = Str(vemitido)
    srecibido.Caption = Str(vrecibido)

    If Err Then Exit Sub
End Sub

Public Sub calsaldocliente()
    On Error Resume Next

    Dim i As Long
    i = 0
    g.Column = 1

    Dim vsaldo As Double

    bccliente.Refresh
    bccliente.Recordset.MoveFirst

    Do Until bccliente.Recordset.EOF
        vsaldo = vsaldo - bccliente.Recordset("Credito") + bccliente.Recordset("debito")
        i = i + 1
    
        g.Row = i
        g.Data = vsaldo
    
        bccliente.Recordset.MoveNext
    
    Loop

    bccliente.Recordset.MoveLast
    sclientes.Caption = Str(vsaldo)

    If Err Then Exit Sub
End Sub

Private Sub calsaldoproveedor()
    On Error Resume Next

    Dim i As Long
    i = 0
    g.Column = 2

    Dim vsaldo As Double

    bccliente.Refresh
    bccliente.Recordset.MoveFirst

    Do Until bccliente.Recordset.EOF
        vsaldo = vsaldo + bccliente.Recordset("Credito") - bccliente.Recordset("debito")
        i = i + 1
        g.Row = i
        g.Data = vsaldo
        bccliente.Recordset.MoveNext
    Loop

    bccliente.Recordset.MoveLast
    sproveedor.Caption = Str(vsaldo)

    If Err Then Exit Sub
End Sub

Private Sub Command1_Click()

    filtrocliente
    calsaldocliente

    filtroproveedor
    calsaldoproveedor

    filtrocaja
    calsaldocaja

    filtrocheque
    calsaldocheque

    saldo.Caption = Str(Val(sclientes.Caption) - Val(sproveedor.Caption) + Val(scaja.Caption) + Val(srecibido.Caption) - Val(semitido.Caption))

End Sub

'----------
Private Sub filtrocaja()
    On Error Resume Next
    bcaja.RecordSource = "select * from caja where fecha >= '" & strfechaMySQL(fdesde.Value) + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "'"
    bcaja.Refresh
    calsaldocaja

    If Err Then Exit Sub
End Sub

Private Sub filtrocheque()
    On Error Resume Next
    bcheques.RecordSource = "select * from cheques where fecha >= '" & strfechaMySQL(fdesde.Value) + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "'  order by fecha"
   
    bcheques.Refresh
    calsaldocheque

    If Err Then Exit Sub
End Sub

'----------

Private Sub filtrocliente()
    On Error Resume Next
    bccliente.RecordSource = "select * from cuentascorrientes where fecha >= '" & strfechaMySQL(fdesde.Value) + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "'  order by fecha "
    bccliente.Refresh
    calsaldocliente

    If Err Then Exit Sub
End Sub

'--------

Private Sub filtroproveedor()
    On Error Resume Next
    bccliente.RecordSource = "select * from pcuentascorrientes where fecha >= '" & strfechaMySQL(fdesde.Value) + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "'  order by fecha "
    bccliente.Refresh
    calsaldoproveedor

    If Err Then Exit Sub
End Sub

Private Sub Form_Load()

    With bccliente
        .ConnectionString = pathDBMySQL
        .RecordSource = "cuentascorrientes"
        .Refresh
    End With

    With bcheques
        .ConnectionString = pathDBMySQL
        .RecordSource = "cheques"
        .Refresh
    End With

    With bcliente
        .ConnectionString = pathDBMySQL
        .RecordSource = "select nombre,saldo  from Clientes"
        .Refresh
    End With

    With bcaja
        .ConnectionString = pathDBMySQL
        .RecordSource = "caja"
        .Refresh
    End With
    
    Me.Top = 200
    Me.Left = 500
    Me.Width = 11280
    Me.Height = 6660
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub saldo_Change()
    saldo.Caption = Format(saldo.Caption, "#############.00")
End Sub

Private Sub scaja_Change()
    scaja.Caption = Format(scaja.Caption, "#############.00")
End Sub

Private Sub sclientes_Change()
    sclientes.Caption = Format(sclientes.Caption, "#############.00")
End Sub

Private Sub semitido_Change()
    semitido.Caption = Format(semitido.Caption, "#############.00")
End Sub

Private Sub sproveedor_Change()
    sproveedor.Caption = Format(sproveedor.Caption, "#############.00")
End Sub

Private Sub srecibido_Change()
    srecibido.Caption = Format(srecibido.Caption, "#############.00")
End Sub

