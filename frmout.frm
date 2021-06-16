VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.TaskPanel.v13.0.0.Demo.ocx"
Begin VB.Form frmOut 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3660
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   7620
   ShowInTaskbar   =   0   'False
   Begin XtremeTaskPanel.TaskPanel TaskPanel1 
      Height          =   3705
      Left            =   -120
      TabIndex        =   1
      Top             =   -30
      Width           =   8385
      _Version        =   851968
      _ExtentX        =   14790
      _ExtentY        =   6535
      _StockProps     =   64
      VisualTheme     =   8
      ItemLayout      =   1
      HotTrackStyle   =   3
      Begin VB.Frame Frame1 
         Height          =   2835
         Left            =   240
         TabIndex        =   2
         Top             =   390
         Width           =   7320
         Begin VB.CheckBox chkAcciones 
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   7095
         End
         Begin VB.CheckBox chkAcciones 
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   960
            Width           =   7095
         End
         Begin VB.CheckBox chkAcciones 
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   7
            Top             =   1320
            Width           =   7095
         End
         Begin VB.CheckBox chkAcciones 
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   6
            Top             =   1680
            Width           =   7095
         End
         Begin VB.CheckBox chkAcciones 
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   5
            Top             =   2040
            Width           =   7095
         End
         Begin VB.CheckBox chkAcciones 
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   7095
         End
         Begin VB.CheckBox chkAcciones 
            Height          =   375
            Index           =   6
            Left            =   120
            TabIndex        =   3
            Top             =   2400
            Width           =   7095
         End
      End
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3240
      Width           =   7335
      _Version        =   851968
      _ExtentX        =   12938
      _ExtentY        =   661
      _StockProps     =   93
   End
   Begin VB.Timer Reloj 
      Left            =   7560
      Top             =   0
   End
End
Attribute VB_Name = "frmOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
On Error Resume Next
    
    With Reloj
        .Enabled = True
        .Interval = 150
    End With
    
    With Me
        .Height = 3750
        .Width = 7600
        .Top = 1000
        .Left = .Width - (.Width / 2)
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Salir()
Dim vArchivoZip As String
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ")", PathDBConfig)
    
    If Not vParametro = "NoBackup" Then
    
        MousePointer = vbDefault
    
        If vConfigGral.vServidor = "Local" Or vConfigGral.vServidor = "MySQL" Then
            With chkAcciones(0)
                .Caption = "Vaciando Tablas Temporales ..."
                .Value = 1
                If Not VaciarTemporales = False Then
                    .Caption = .Caption & "Realizada"
                Else
                    .Caption = .Caption & "Realizada.... con errores"
                End If
                barra.Value = 14
            End With
            With chkAcciones(1)
                .Caption = "Copia de BBDD a archivo temporal ..."
                .Value = 1
                CopiarBorrar 0
                .Caption = .Caption & "Realizada"
                barra.Value = 28
            End With
            With chkAcciones(2)
                .Caption = "Comprimiendo BBDD a Archivo Zip para Backup ..."
                .Value = 1
                vArchivoZip = Comprimir
                .Caption = .Caption & "Realizada"
                barra.Value = 42
            End With
            With chkAcciones(3)
                .Caption = "Borrado de BBDD temporal ..."
                .Value = 1
                CopiarBorrar (1)
                .Caption = .Caption & "Realizada"
                barra.Value = 56
            End With
            With chkAcciones(4)
                .Caption = "Copiando Copia de Seguridad en Unidad Portatil ..."
                .Value = 1
                Call CopiarBorrar(2, vArchivoZip)
                .Caption = .Caption & "Realizada"
                barra.Value = 70
            End With
            With chkAcciones(5)
                .Caption = "Subiendo Base de datos a Servidor FTP ..."
                .Value = 1
                'FTP (vArchivoZip)
                .Caption = .Caption & "No Realizada"
                barra.Value = 84
            End With
            
            With chkAcciones(6)
                .Caption = "Cerrando los Formularios.........."
                .Value = 1
                CerrarForms 1
                .Caption = .Caption & "Realizada"
                barra.Value = 100
            End With
        
        
        Else
            End
        End If
    Else
        End
    End If
    
    MousePointer = vbHourglass
    
If Err Then GrabarLog "Salir", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Reloj_Timer()
    Salir
    Reloj.Enabled = False
End Sub

