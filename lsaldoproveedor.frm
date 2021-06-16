VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSaldosProveedor 
   Caption         =   "Listado de Saldos Proveedor "
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3630
   ScaleWidth      =   4455
   Begin VB.CommandButton Command7 
      BackColor       =   &H80000004&
      Caption         =   "Imprimir"
      Height          =   525
      Left            =   2850
      Picture         =   "lsaldoproveedor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Generar reporte para imprimir"
      Top             =   2130
      UseMaskColor    =   -1  'True
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Height          =   1245
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   2385
      Begin VB.OptionButton Option4 
         Caption         =   "Todos los Clientes"
         Height          =   225
         Left            =   90
         TabIndex        =   12
         Top             =   930
         Width           =   2265
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Saldos a favor de Clientes"
         Height          =   225
         Left            =   90
         TabIndex        =   11
         Top             =   690
         Width           =   2205
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Saldados"
         Height          =   225
         Left            =   90
         TabIndex        =   10
         Top             =   450
         Width           =   2025
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Deudores"
         Height          =   225
         Left            =   90
         TabIndex        =   9
         Top             =   210
         Width           =   2025
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Con detalles"
      Height          =   285
      Left            =   2760
      TabIndex        =   7
      Top             =   1530
      Width           =   1545
   End
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   330
      Left            =   240
      Top             =   3120
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
   Begin VB.Frame Frame1 
      Caption         =   "Proveedor :"
      ForeColor       =   &H00000080&
      Height          =   1275
      Left            =   180
      TabIndex        =   0
      Top             =   30
      Width           =   4035
      Begin VB.TextBox vlocalidad 
         Height          =   315
         Left            =   1020
         TabIndex        =   5
         Top             =   900
         Width           =   2865
      End
      Begin VB.TextBox vcdesde 
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   240
         Width           =   2865
      End
      Begin VB.TextBox vchasta 
         Height          =   315
         Left            =   1020
         TabIndex        =   1
         Top             =   570
         Width           =   2865
      End
      Begin VB.Label Label3 
         Caption         =   "Localidad :"
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   930
         Width           =   825
      End
      Begin VB.Label Label1 
         Caption         =   "Desde :"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   300
         Width           =   705
      End
      Begin VB.Label Label2 
         Caption         =   "Hasta :"
         Height          =   225
         Left            =   210
         TabIndex        =   3
         Top             =   600
         Width           =   675
      End
   End
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   240
      Top             =   2760
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
Attribute VB_Name = "frmSaldosProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vcodigodesde, vcodigohasta As String

Function calsaldo(vcodigo As String)
    On Error Resume Next
    Dim vtotal As Double
    bccliente.RecordSource = "select * from pcuentascorrientes where codigo = '" + vcodigo + "'"
    bccliente.Refresh
    vtotal = 0
    Do Until bccliente.Recordset.EOF
        vtotal = vtotal - bccliente.Recordset("debito") + bccliente.Recordset("credito")
        bccliente.Recordset.MoveNext
    Loop
    calsaldo = vtotal
    If Err Then Exit Function
End Function


Private Sub buscaprov(vprov As String, _
                      dh As String)
    bcliente.RecordSource = "select * from proveedores where (nombre = '" + vprov + "') or (codigo = '" + vprov + "')"
    bcliente.Refresh
    If bcliente.Recordset.EOF Then
        frmBuscaProveedor.Show
        If dh = "d" Then
            frmBuscaProveedor.o = 8
        Else
            frmBuscaProveedor.o = 9
        End If
        frmBuscaProveedor.Show
        frmBuscaProveedor.varticulo = vprov
        'frmbuscacliente.varticulo_KeyPress (13)
        frmBuscaProveedor.Show
        frmBuscaProveedor.varticulo.SetFocus
    Else
        Dim j As Integer
        If dh = "d" Then
            vcdesde = bcliente.Recordset("nombre")
            vcodigodesde = bcliente.Recordset(0)
            vchasta.SetFocus
        Else
            vchasta = bcliente.Recordset("nombre")
            vcodigohasta = bcliente.Recordset(0)
            vlocalidad.SetFocus
        End If
    End If
End Sub

Private Sub Limpiar()
    vcdesde.Text = ""
    vchasta.Text = ""
    vlocalidad.Text = ""
    vcodigodesde = ""
    vcodigohasta = ""
End Sub

Public Sub vcdesde_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        buscaprov vcdesde, "d"
    End If
End Sub

Public Sub vchasta_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        buscaprov vchasta, "h"
    End If
End Sub

Public Sub vlocalidad_KeyPress(Keyascii As Integer)
    If Keyascii = 13 Then
        Command7.SetFocus
    End If
End Sub

Private Sub Command7_Click()
    CambiarImpresora 1
    On Error Resume Next
    Dim vFiltro As String
    vFiltro = ""
    MousePointer = vbHourglass
    If Not (vcodigodesde = "" And vcodigohasta = "") Then
        vFiltro = vFiltro + " and codigo >= '" + vcodigodesde + "' and codigo <= '" + vcodigohasta + "'"
    End If
    If Not vlocalidad = "" Then
        vFiltro = vFiltro + " and localidad = '" + vlocalidad + "'"
    End If
    bcliente.Refresh
    Do Until bcliente.Recordset.EOF
        bcliente.Recordset("Saldo") = calsaldo(bcliente.Recordset("codigo"))
        bcliente.Recordset.Update
        bcliente.Recordset.MoveNext
    Loop
    If Option2 Then mantenimiento.rsclprov.Filter = "saldo = 0" + vFiltro
    If Option3 Then mantenimiento.rsclprov.Filter = "saldo < 0" + vFiltro
    If Option1 Then mantenimiento.rsclprov.Filter = "saldo > 0" + vFiltro
    If Option4 Then mantenimiento.rsclprov.Filter = "credito >= 0 " + vFiltro
    If Not mantenimiento.rsclprov.State = 1 Then
        mantenimiento.rsclprov.Open
        mantenimiento.rsclprov.Close
        mantenimiento.rsclprov.Open
    Else
        mantenimiento.rsclprov.Close
        mantenimiento.rsclprov.Open
    End If
    mantenimiento.rsclprov.Sort = "nombre"
    MsgBox "Prepare la impresora !", vbInformation, "Mensaje ..."
    drproveedor.Refresh
    drproveedor.Show
    Limpiar
    MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    With bccliente
        .ConnectionString = pathDB
        .RecordSource = "pcuentascorrientes"
        .Refresh
    End With

    With bcliente
        .ConnectionString = pathDB
        .RecordSource = "Proveedores"
        .Refresh
    End With
    With Me
        .Top = 1000
        .Left = 1300
        .Width = 4800
        .Height = 3090
    End With
End Sub

