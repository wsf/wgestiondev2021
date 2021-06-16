VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5010
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   5145
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5010
   ScaleWidth      =   5145
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraSplash 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   6000
      Left            =   -270
      TabIndex        =   0
      Top             =   -120
      Width           =   5385
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   285
         Left            =   315
         TabIndex        =   9
         Top             =   4725
         Width           =   660
         _Version        =   851968
         _ExtentX        =   1164
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Entrar"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Timer tmrUnload 
         Interval        =   100
         Left            =   4860
         Top             =   2520
      End
      Begin XtremeSuiteControls.PushButton PusLogin 
         Height          =   285
         Left            =   990
         TabIndex        =   10
         Top             =   4725
         Width           =   615
         _Version        =   851968
         _ExtentX        =   1085
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Login"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ProgressBar conteo 
         Height          =   4395
         Left            =   5175
         TabIndex        =   8
         Top             =   225
         Width           =   180
         _Version        =   851968
         _ExtentX        =   317
         _ExtentY        =   7752
         _StockProps     =   93
         BackColor       =   0
         Min             =   1
         Max             =   25
         Scrolling       =   1
         Orientation     =   1
         Appearance      =   6
         UseVisualStyle  =   0   'False
         BarColor        =   65280
      End
      Begin VB.PictureBox Picture1 
         Height          =   4470
         Left            =   315
         Picture         =   "Splash.frx":1CFA
         ScaleHeight     =   4410
         ScaleWidth      =   4995
         TabIndex        =   11
         Top             =   180
         Width           =   5055
      End
      Begin VB.Label lblWGESTION 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   5640
         TabIndex        =   7
         Top             =   630
         Width           =   1800
      End
      Begin VB.Label lblCompania 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   330
         Left            =   1590
         TabIndex        =   4
         Top             =   900
         Width           =   5715
      End
      Begin VB.Label lblTipoVersion 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   330
         Left            =   240
         TabIndex        =   6
         Top             =   180
         Width           =   7020
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1125
         TabIndex        =   2
         Top             =   2970
         Width           =   3015
      End
      Begin VB.Label lblLicence 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   1
         Top             =   3540
         Width           =   1860
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7110
         TabIndex        =   3
         Top             =   3150
         Width           =   105
      End
      Begin VB.Label lblDemo 
         BackStyle       =   0  'Transparent
         Height          =   885
         Left            =   45
         TabIndex        =   5
         Top             =   4455
         Width           =   7320
      End
      Begin VB.Image imgLogo 
         Height          =   6255
         Index           =   0
         Left            =   45
         Stretch         =   -1  'True
         Top             =   -60
         Width           =   6000
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Initialize()

' me.Picture =
On Error Resume Next

Me.Picture1.Picture = LoadPicture(App.Path + "\logo_inst.jpg")

Me.Refresh


'If AppRunning Then
'        vb.Unload Me
'End If



If Err Then

'If AppRunning Then
        vb.Unload Me
'End If

End If
End Sub

Private Function AppRunning() As Boolean
    If vb.App.PrevInstance Then
        'MsgBox vb.App.EXEName & " ya se está ejecutando! ", vbCritical + vbSystemModal
        AppRunning = True
    End If
End Function



Private Sub Form_KeyPress(KeyAscii As Integer)
    VerEntradaAlSistema
End Sub
Private Sub Form_Load()

  '  Call actualizar

    initPersonalizacion

    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTipoVersion.Caption = LeerConfig(8)
    
    'Me.Show
    
    If CBool(LeerConfig(7)) = True Then
        'InitializeSystem
    End If
    
   ' Call controles  ' ejecuta controles
    
    
    
End Sub


Private Sub initPersonalizacion()
vAsientoAutomatico = LeerConfig(28)
End Sub

Private Sub fraSplash_Click()
    VerEntradaAlSistema
End Sub
Private Sub imgLogo_Click(Index As Integer)
    VerEntradaAlSistema
End Sub

Private Sub lblDescripcion_Click()
    VerEntradaAlSistema
End Sub

Private Sub ShoNombreDe_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub

Private Sub PushButton1_Click()
'Call actualizar
'frmPrincipal.Show
frmPrincipal.Show
End Sub

Private Sub PusLogin_Click()
frmLogin.Show
End Sub

Private Sub tmrUnload_Timer()
    Static Timeout As Integer
    Timeout = Timeout + 1
    conteo.Value = Timeout
    If Timeout >= 25 Then
        tmrUnload.Enabled = False
        VerEntradaAlSistema
        Unload Me
    Else
        Me.ZOrder
    End If
End Sub
Private Sub VerEntradaAlSistema()
On Error Resume Next

    Unload Me
    
    Select Case LeerConfig(6)
    
        Case "Auto"
                Call Login(LeerConfig(2), LeerConfig(4), LeerConfig(3), LeerConfig(5))
                'MsgBox ""
                
                mensaje "salio de login"
                
                'Call PushButton1_Click
                frmPrincipal.Show
               ' frmArticulos.Show
                Unload Me
               Exit Sub
     
        Case "Manual"
               frmLogin.Show
        Case Else
          frmPrincipal.Show
    End Select



If Err Then
'mensaje Err.Description
'MsgBox Err.Description
'GrabarLog "VerEntradaAlSistema", Err.Number & " " & Err.Description, Me.Caption
End If
End Sub
