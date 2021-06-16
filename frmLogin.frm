VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmLogin 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Login -  Gestión Comercial - WSF"
   ClientHeight    =   3360
   ClientLeft      =   2835
   ClientTop       =   3495
   ClientWidth     =   4380
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.CheckBox chkGuardarDatos 
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   4095
      _Version        =   851968
      _ExtentX        =   7223
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Guardar datos para la proxima sesión"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdLogin 
      Height          =   375
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Top             =   2895
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Aceptar"
      Appearance      =   2
   End
   Begin VB.ComboBox cboEmpresas 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   480
      Width           =   2400
   End
   Begin VB.ComboBox cboServidor 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Width           =   2400
   End
   Begin VB.ComboBox cboUsuario 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   2400
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   180
      Picture         =   "frmLogin.frx":6852
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   2160
      Width           =   480
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1200
      Width           =   2400
   End
   Begin XtremeSuiteControls.PushButton cmdLogin 
      Height          =   375
      Index           =   1
      Left            =   3120
      TabIndex        =   11
      Top             =   2895
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "&Cancelar"
      Appearance      =   2
   End
   Begin XtremeSuiteControls.PushButton cmdLimpiar 
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2880
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Con&figuracion"
      BackColor       =   12648447
      Appearance      =   2
   End
   Begin VB.Label lblDatos 
      Alignment       =   1  'Right Justify
      Caption         =   "&Contraseña :"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   0
      TabIndex        =   10
      Top             =   1240
      Width           =   1755
   End
   Begin VB.Label lblDatos 
      Alignment       =   1  'Right Justify
      Caption         =   "&Servidor BBDD :"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   8
      Top             =   160
      Width           =   1757
   End
   Begin VB.Label lblSistema 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H008C6063&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   2280
      Width           =   4365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00404040&
      BackStyle       =   1  'Opaque
      Height          =   735
      Left            =   0
      Top             =   2040
      Width           =   4365
   End
   Begin VB.Label lblDatos 
      Alignment       =   1  'Right Justify
      Caption         =   "&Empresa:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   520
      Width           =   1755
   End
   Begin VB.Label lblDatos 
      Alignment       =   1  'Right Justify
      Caption         =   "&Nombre de usuario:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   880
      Width           =   1755
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboEmpresas_GotFocus()
On Error Resume Next

    Call CargarCombo("Empresas", "Alias", cboEmpresas, True, , PathDBConfig)

If Err Then GrabarLog "cboEmpresas_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboServidor_GotFocus()
On Error Resume Next

    Call CargarCombo("Servidor", "Servidor", cboServidor, True, , PathDBConfig)

If Err Then GrabarLog "cboServidor_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cboUsuario_GotFocus()
On Error Resume Next

    Call CargarCombo("Usuarios", "usuario", cboUsuario, True, , PathDBConfig)

If Err Then GrabarLog "cboUsuario_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub chkGuardarDatos_Click()
On Error Resume Next

If Err Then GrabarLog "chkGuardarDatos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdLimpiar_Click()
On Error Resume Next

'    frmConfiguracion.Show
    Unload frmLogin
    
    Exit Sub
    cboServidor.Tag = ""
    cboServidor.Text = ""
    cboServidor.Enabled = True
    
    cboEmpresas.Tag = ""
    cboEmpresas.Text = ""
    cboEmpresas.Enabled = True
    
    cboUsuario.Tag = ""
    cboUsuario.Text = ""
    cboUsuario.Enabled = True
    
    txtPassword.Text = ""
    
    chkGuardarDatos.Value = xtpUnchecked

If Err Then GrabarLog "cmdLimpiar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdLogin_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
    
            Select Case Login(cboServidor.Text, cboEmpresas.Text, cboUsuario.Text, txtPassword.Text)
    
                Case "Correcto"
                    Me.Hide
                    frmPrincipal.Show
        
                    Dim i As Integer
                    If chkGuardarDatos.Value = xtpChecked Then
                        For i = 0 To 2
                            Call LeerConfig("UServidor", cboServidor.Text)
                            Call LeerConfig("UUsuario", cboUsuario.Text)
                            Call LeerConfig("UEmpresa", cboEmpresas.Text)
                        Next
                    Else
                        For i = 0 To 2
                            Call LeerConfig("UServidor", "")
                            Call LeerConfig("UUsuario", "")
                            Call LeerConfig("UEmpresa", "")
                        Next

                    End If
                
                Case "Servidor"
                Case "Empresa"
                Case "Empresa-Usuario"
                    MsgBox "El Usuario no tiene Asociada la Empresa que ha elegido", vbExclamation, "Inicio de sesión"
                    cboEmpresas.SetFocus
                    SendKeys "{Home}+{End}"
        
                Case "Usuario"
                Case "Password"
                    MsgBox "La contraseña no es válida. Vuelva a intentarlo", , "Inicio de sesión"
                    txtPassword.SetFocus
                    SendKeys "{Home}+{End}"
            End Select
        
        Case 1
            Me.Hide
            End
    End Select

If Err Then GrabarLog "cmdLogin_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Initialize()
On Error Resume Next

    If App.PrevInstance = True Then
        'MsgBox " Ya tiene una Instancia en Ejecución ", vbExclamation, "Mensaje ..."
        'End
        'Unload App.PrevInstance
        'App.PrevInstance
    End If
    
If Err Then GrabarLog "Form_Initialize", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next
   
    Dim i As Double
    
    For i = 1 To ScaleWidth
        Picture1.Line (i, Picture1.ScaleHeight)-(i, 0), RGB(0, 0, i)
        Line (i, ScaleHeight)-(i, 0), RGB(i, i, i)
    Next i
   
    lblSistema.Caption = "Gestión Comercial v" & App.Major & "." & App.Minor & "." & App.Revision
   
    If Trim(LeerConfig(2)) = "" Or Trim(LeerConfig(3)) = "" Or Trim(LeerConfig(4)) = "" Then
        cboServidor.Tag = ""
        cboUsuario.Tag = ""
        cboEmpresas.Tag = ""
    Else
        Me.Show
        chkGuardarDatos.Value = xtpChecked
        
        cboServidor.Text = LeerConfig(2)
        'cboServidor.Enabled = False
        
        cboUsuario.Text = LeerConfig(3)
        'cboUsuario.Enabled = False
        
        cboEmpresas.Text = LeerConfig(4)
        'cboEmpresas.Enabled = False
        
        txtPassword.SetFocus
    End If
   
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtPassword_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        cmdLogin(0).SetFocus
    End If
    
If Err Then GrabarLog "txtPassword_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
