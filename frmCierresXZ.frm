VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmCierresXZ 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.GroupBox gbTipoDeListado 
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6735
      _Version        =   851968
      _ExtentX        =   11880
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "TipoDeListado"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton rbCierres 
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Lectura X"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbCierres 
         Height          =   255
         Index           =   1
         Left            =   1680
         TabIndex        =   12
         Top             =   480
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Cierre Z"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.ComboBox cboImpresora 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   1320
      Width           =   4935
      _Version        =   851968
      _ExtentX        =   8705
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmCierresXZ.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   7005
      TabIndex        =   0
      Top             =   2500
      Width           =   7000
      Begin XtremeSuiteControls.PushButton cmdImprimir 
         Height          =   375
         Left            =   4440
         TabIndex        =   1
         Top             =   90
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Imprimir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCierresXZ.frx":50B3
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   375
         Left            =   5640
         TabIndex        =   2
         Top             =   90
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCierresXZ.frx":564D
      End
      Begin VB.Label lblWGESTION2010 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   0
         Left            =   50
         TabIndex        =   3
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGESTION2010 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   4
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.ComboBox cboModelo 
      Height          =   315
      Left            =   1920
      TabIndex        =   9
      Top             =   1680
      Width           =   4935
      _Version        =   851968
      _ExtentX        =   8705
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin XtremeSuiteControls.ComboBox cboPuerto 
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Top             =   2040
      Width           =   4935
      _Version        =   851968
      _ExtentX        =   8705
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Locked          =   -1  'True
   End
   Begin XtremeSuiteControls.Label lblImpresora 
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   1750
      _Version        =   851968
      _ExtentX        =   3087
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Puerto :"
   End
   Begin XtremeSuiteControls.Label lblImpresora 
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   1750
      _Version        =   851968
      _ExtentX        =   3087
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Modelo :"
   End
   Begin XtremeSuiteControls.Label lblImpresora 
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1750
      _Version        =   851968
      _ExtentX        =   3087
      _ExtentY        =   450
      _StockProps     =   79
      Caption         =   "Impresora :"
   End
End
Attribute VB_Name = "frmCierresXZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboImpresora_GotFocus()
On Error Resume Next

    Call CargarComboNew("Impresoras", "Impresora", cboImpresora, True, , PathDBConfig)

If Err Then GrabarLog "cboImpresora_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboModelo_GotFocus()
On Error Resume Next

    Call CargarComboNew("Impresoras", "Modelo", cboModelo, True, , PathDBConfig)

If Err Then GrabarLog "cboModelo_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboPuerto_GotFocus()
On Error Resume Next

    Call CargarComboNew("Impresoras", "Puerto", cboPuerto, True, , PathDBConfig)

If Err Then GrabarLog "cboModelo_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdCerrar_Click()
On Error Resume Next

    Unload Me

If Err Then GrabarLog "cmdCerrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    cboImpresora.Text = vImpresoras.vNombreImpresora
    cboModelo.Text = vImpresoras.vModelo
    cboPuerto.Text = vImpresoras.vNroPuerto

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdImprimir_Click()
On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To Val(rbCierres.Count) - 1
        
        If rbCierres(i).Value = True Then
        
            Select Case i
        
                Case 0
                    ImprimirCierreX
            
                Case 1
                    ImprimirCierreZ
            
                Case Else
        
            End Select
        
        End If
        
    Next
    
If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ImprimirCierreX()
On Error Resume Next

    Select Case vImpresoras.vNombreImpresora

        Case "Epson"
            With frmPrincipal.FiscalEpson
                .BaudRate = 9600
                .PortNumber = vImpresoras.vNroPuerto

                .CloseJournal "X", "P"
            End With

        Case "Hasar"
            With frmPrincipal.FiscalHasar
                '.Puerto = vImpresoras.vNroPuerto
                '.Modelo = vImpresoras.vModeloInterno
                '.Comenzar
                '.TratarDeCancelarTodo
                .ReporteX
            End With

    End Select

If Err Then GrabarLog "ImprimirCierreX", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ImprimirCierreZ()
On Error Resume Next

    Select Case vImpresoras.vNombreImpresora

        Case "Epson"
            With frmPrincipal.FiscalEpson
                .BaudRate = 9600
                .PortNumber = vImpresoras.vNroPuerto
                .CloseJournal "Z", "P"
            End With

        Case "Hasar"
            With frmPrincipal.FiscalHasar
                '.Puerto = vImpresoras.vNroPuerto
                '.Modelo = vImpresoras.vModeloInterno
                '.Comenzar
                '.TratarDeCancelarTodo
                .ReporteZ
            End With

    End Select

If Err Then GrabarLog "ImprimirCierreZ", Err.Number & " " & Err.Description, Me.Caption
End Sub

