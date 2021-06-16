VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmBancosCuentaAlta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mantenimiento de Cuentas Bancarias"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7485
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicInferior 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      Picture         =   "frmBancosCuentaAlta.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   7500
      TabIndex        =   8
      Top             =   -30
      Width           =   7500
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   300
         TabIndex        =   18
         Top             =   150
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar"
         Appearance      =   4
         Picture         =   "frmBancosCuentaAlta.frx":50B3
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   6300
         TabIndex        =   19
         Top             =   120
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   4
         Picture         =   "frmBancosCuentaAlta.frx":54BA
      End
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   2535
      Left            =   90
      TabIndex        =   9
      Top             =   630
      Width           =   7335
      _Version        =   851968
      _ExtentX        =   12938
      _ExtentY        =   4471
      _StockProps     =   68
      Color           =   8
      ItemCount       =   1
      Item(0).Caption =   "Ficha"
      Item(0).ControlCount=   16
      Item(0).Control(0)=   "txtAlta(0)"
      Item(0).Control(1)=   "txtAlta(1)"
      Item(0).Control(2)=   "txtAlta(2)"
      Item(0).Control(3)=   "lblAlta(2)"
      Item(0).Control(4)=   "lblAlta(1)"
      Item(0).Control(5)=   "lblAlta(0)"
      Item(0).Control(6)=   "txtAlta(3)"
      Item(0).Control(7)=   "txtAlta(4)"
      Item(0).Control(8)=   "lblAlta(3)"
      Item(0).Control(9)=   "pbCarga(0)"
      Item(0).Control(10)=   "txtAlta(5)"
      Item(0).Control(11)=   "pbCarga(1)"
      Item(0).Control(12)=   "lblAlta(4)"
      Item(0).Control(13)=   "txtAlta(6)"
      Item(0).Control(14)=   "txtAlta(7)"
      Item(0).Control(15)=   "pbCarga(2)"
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   2
         Left            =   2760
         TabIndex        =   2
         Top             =   960
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   255
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   0
         Left            =   2760
         TabIndex        =   0
         Top             =   600
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   10
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   1
         Left            =   4200
         TabIndex        =   1
         Top             =   600
         Width           =   3015
         _Version        =   851968
         _ExtentX        =   5318
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   0
         Left            =   3720
         TabIndex        =   14
         Tag             =   "Banco"
         Top             =   600
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   4
         Left            =   2760
         TabIndex        =   4
         Top             =   1680
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   10
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   5
         Left            =   4200
         TabIndex        =   5
         Top             =   1680
         Width           =   3015
         _Version        =   851968
         _ExtentX        =   5318
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   15
         Tag             =   "CodigoCuenta"
         Top             =   1680
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   6
         Left            =   2760
         TabIndex        =   7
         Top             =   2040
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   10
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   7
         Left            =   4200
         TabIndex        =   6
         Top             =   2040
         Width           =   3015
         _Version        =   851968
         _ExtentX        =   5318
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   2
         Left            =   3720
         TabIndex        =   17
         Tag             =   "TipoCuentaBanco"
         Top             =   2040
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   3
         Left            =   2760
         TabIndex        =   3
         Top             =   1320
         Width           =   2175
         _Version        =   851968
         _ExtentX        =   3836
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   255
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Cuenta de Banco :"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   16
         Top             =   2080
         Width           =   2205
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   13
         Top             =   1360
         Width           =   2200
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion de la Cuenta :"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   12
         Top             =   1000
         Width           =   2200
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Banco :"
         Height          =   195
         Index           =   0
         Left            =   540
         TabIndex        =   11
         Top             =   600
         Width           =   2205
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable Asociada :"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   10
         Top             =   1720
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmBancosCuentaAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vAccion As String
Private Sub Form_Load()
On Error Resume Next
    
    vAccion = "Nuevo"
    Me.KeyPreview = True

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Guardar

        Case 1
            Limpiar
            Unload Me
                        
    End Select
    
If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Limpiar()
On Error Resume Next

    Dim i As Integer
    
    With Me
        For i = 0 To txtAlta.Count - 1
            txtAlta(i).Text = ""
            txtAlta(i).Tag = ""
        Next
    
    End With

If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Guardar()
On Error Resume Next

    If ValidarCampos = Not True Then
        MsgBox "Uno de los campos No es Valido o no esta bien Cargado!!", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    Dim rsBancosCuentas As New ADODB.Recordset, sqlBancosCuentas As String
    
    Select Case vAccion

        Case "Nuevo"
            sqlBancosCuentas = "SELECT * FROM BancosCuentas WHERE 1=2"
        
        Case "Modificar"
            sqlBancosCuentas = "SELECT * FROM BancosCuentas WHERE (idBancosCuentas = " & Trim$(txtAlta(0).Tag) & ")"
        
        Case "Duplicar"
        
    End Select
    
    With rsBancosCuentas
        Call .Open(sqlBancosCuentas, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then If .EOF = True Then .AddNew
            
        .Fields("idBancos").Value = txtAlta(0).Text
        
        .Fields("Cuenta").Value = Trim(txtAlta(3).Text)
        .Fields("Descripcion").Value = Left(txtAlta(2).Text, 255)
        .Fields("CuentaContableAsociada").Value = Trim(txtAlta(4).Text)
        .Fields("idTipoCuentaBanco").Value = Trim(txtAlta(6).Text)
    
        .Update
    End With
    
    Limpiar
    
    sqlBancosCuentas = ""
    
    If rsBancosCuentas.State = 1 Then
        rsBancosCuentas.Close
        Set rsBancosCuentas = Nothing
    End If

If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function ValidarCampos() As Boolean
On Error Resume Next

    Dim i As Integer
    
    ValidarCampos = True
    
    For i = 0 To txtAlta.Count - 4
        If txtAlta(i).Text = "" Then
            ValidarCampos = Not True
            Exit Function
        End If
    Next
    
    If vAccion = "Nuevo" Then
        If Not Trim(TraerDato("BancosCuentas", "(idBancosCuentas = " & Trim(txtAlta(0).Tag) & ")", "idBancosCuentas")) = "" Then
            ValidarCampos = Not True
            Exit Function
        End If
    End If
    
If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub pbCarga_Click(Index As Integer)
    On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
    
        Case 0 To pbCarga.Count - 1
            frmBusqueda.Show

    End Select

If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub ModificarBancosCuentas(vIdBancoCuentas As Long)
    On Error Resume Next
    
    Dim rsBancosCuentas As New ADODB.Recordset, sqlBancosCuentas As String
    
    sqlBancosCuentas = "SELECT * FROM BancosCuentas WHERE (idBancosCuentas = " & vIdBancoCuentas & ")"
    
    With rsBancosCuentas
        Call .Open(sqlBancosCuentas, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
        
            'No Opcionales
            txtAlta(0).Tag = .Fields("idBancosCuentas").Value
            txtAlta(0).Text = .Fields("idBancos").Value
            txtAlta(1).Text = TraerDato("Bancos", "idBancos = '" & .Fields("idBancos").Value & "'", "Descripcion")
            'txtAlta(0).Locked = True
        
            txtAlta(2).Text = .Fields("Descripcion").Value
            txtAlta(3).Text = .Fields("Cuenta").Value
            
            txtAlta(4).Text = .Fields("CuentaContableAsociada").Value
            txtAlta(5).Text = TraerDato("Cuentas", "CodigoCuenta = '" & .Fields("CuentaContableAsociada").Value & "'", "Cuenta")
        
            txtAlta(6).Text = .Fields("idTipoCuentaBanco").Value
            txtAlta(7).Text = TraerDato("TipoCuentaBanco", "idTipoCuentaBanco = '" & .Fields("idTipoCuentaBanco").Value & "'", "TipoCuentaBanco")
        End If

    End With
        
    sqlBancosCuentas = ""
    
    If rsBancosCuentas.State = 1 Then
        rsBancosCuentas.Close
        Set rsBancosCuentas = Nothing
    End If
    
    If Err Then GrabarLog "ModificarBancosCuentas", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtAlta_KeyPress(Index As Integer, Keyascii As Integer)
On Error Resume Next

    If Keyascii = 13 Then
        If Index = Val(txtAlta.Count - 2) Then
            PbAcciones(0).SetFocus
        Else
            txtAlta(Index + 1).Selfocus
        End If
    End If
    
If Err Then GrabarLog "txtAlta_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
