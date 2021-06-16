VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmSaldosProveedores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Saldos Proveedor "
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkDetalles 
      Caption         =   "Con detalles"
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   1560
      Width           =   1545
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Proveedor :"
      ForeColor       =   &H00000080&
      Height          =   1395
      Left            =   180
      TabIndex        =   3
      Top             =   30
      Width           =   4875
      Begin VB.TextBox txtLocalidad 
         Height          =   315
         Left            =   1020
         TabIndex        =   2
         Top             =   900
         Width           =   3500
      End
      Begin VB.TextBox txtProveedor 
         Height          =   315
         Index           =   0
         Left            =   1020
         TabIndex        =   0
         Top             =   240
         Width           =   3500
      End
      Begin VB.TextBox txtProveedor 
         Height          =   315
         Index           =   1
         Left            =   1020
         TabIndex        =   1
         Top             =   570
         Width           =   3500
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "Localidad :"
         Height          =   195
         Index           =   2
         Left            =   30
         TabIndex        =   6
         Top             =   930
         Width           =   950
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "Desde :"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   5
         Top             =   300
         Width           =   950
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "Hasta :"
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   4
         Top             =   600
         Width           =   950
      End
   End
   Begin MSAdodcLib.Adodc bccliente 
      Height          =   330
      Left            =   2040
      Top             =   90
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
   Begin XtremeSuiteControls.PushButton cmdEjecutar 
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   2040
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Listado General"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmSaldosProveedores.frx":0000
   End
   Begin VB.Frame Frame2 
      Height          =   1245
      Left            =   180
      TabIndex        =   8
      Top             =   1380
      Width           =   2745
      Begin VB.OptionButton OpTipoSaldo 
         Caption         =   "Todos los Proveedores"
         Height          =   225
         Index           =   3
         Left            =   30
         TabIndex        =   12
         Top             =   930
         Value           =   -1  'True
         Width           =   2650
      End
      Begin VB.OptionButton OpTipoSaldo 
         Caption         =   "Saldos a favor de Proveedores"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   30
         TabIndex        =   11
         Top             =   690
         Width           =   2650
      End
      Begin VB.OptionButton OpTipoSaldo 
         Caption         =   "Saldados"
         Enabled         =   0   'False
         Height          =   225
         Index           =   1
         Left            =   30
         TabIndex        =   10
         Top             =   450
         Width           =   2650
      End
      Begin VB.OptionButton OpTipoSaldo 
         Caption         =   "Deudores"
         Enabled         =   0   'False
         Height          =   225
         Index           =   0
         Left            =   30
         TabIndex        =   9
         Top             =   210
         Width           =   2650
      End
   End
End
Attribute VB_Name = "frmSaldosProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub BuscarProveedor(vproveedor As String, Index As Integer)
On Error Resume Next

    Dim rsProveedores As New ADODB.Recordset, sqlProveedores As String
    
    sqlProveedores = "SELECT * FROM proveedores WHERE (nombre = '" & vproveedor & "') or (codigo = '" & vproveedor & "')"

    With rsProveedores
        Call .Open(sqlProveedores, ConnDDBB, adOpenStatic, adLockReadOnly)

        If .EOF Then
            frmBuscarProveedor.Show

            If Index = 0 Then
                frmBuscarProveedor.o = 8
            Else
                frmBuscarProveedor.o = 9
            End If
    
            frmBuscarProveedor.Show
            frmBuscarProveedor.txtProveedor = vproveedor
            'frmBuscarProveedor.TXTPROVEEDOR_KeyPress (13)
            frmBuscarProveedor.Show
            frmBuscarProveedor.txtProveedor.SetFocus
    
        Else
    
            If Index = 0 Then
                txtProveedor(0).Text = .Fields("nombre").Value
                txtProveedor(0).Tag = .Fields("Codigo").Value
                txtProveedor(1).SetFocus
            Else
                txtProveedor(1).Text = .Fields("nombre").Value
                txtProveedor(1).Tag = .Fields("Codigo").Value
                txtLocalidad.SetFocus
        
            End If
    
        End If
    
    End With
    
    sqlProveedores = ""
    
    If rsProveedores.State = 1 Then
        rsProveedores.Close
        Set rsProveedores = Nothing
    End If

If Err Then GrabarLog "BuscarProveedor", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdEjecutar_Click()
    On Error Resume Next

    Dim vSQL As String

    vSQL = ""

    MousePointer = vbHourglass

    If Not Trim(txtProveedor(0).Tag) = "" And Trim(txtProveedor(1).Tag) = "" Then vSQL = vSQL + " and codigo >= '" + Trim(txtProveedor(0).Tag) + "' and codigo <= '" + Trim(txtProveedor(1).Tag) + "'"

    If Not txtLocalidad = "" Then vSQL = vSQL + " and (localidad = '" + txtLocalidad + "')"

    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la impresora !", vbInformation, "Mensaje ..."
    
    
    With Mantenimiento.rsSaldosProveedores
        If .State = 1 Then .Close
        
        If vSQL = "" Then
            .Source = .Source
        Else
            .Source = .Source
        End If
        
        If .State = 0 Then .Open
        .Close
        .Open

        .Sort = "Nombre ASC"

    End With

    With drProveedoresSaldo
        .Show
    End With

    Limpiar

    MousePointer = vbDefault
    
If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    With Me
        .Show
        .Top = 1000
        .Left = 1300
        .Width = 5250
        .Height = 3000
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()
On Error Resume Next

    txtProveedor(0).Text = ""
    txtProveedor(0).Tag = ""
    txtProveedor(1).Text = ""
    txtProveedor(1).Tag = ""
    txtLocalidad.Text = ""
    
If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub txtProveedor_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        Call BuscarProveedor(txtProveedor(Index).Text, Index)
    End If
    
If Err Then GrabarLog "txtProveedor_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub txtLocalidad_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        cmdEjecutar.SetFocus
    End If
    
If Err Then GrabarLog "txtLocalidad_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub


