VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmExcel 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   13845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   13845
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.FlatEdit txtBusqueda 
      Height          =   345
      Left            =   2280
      TabIndex        =   7
      Top             =   6630
      Width           =   11415
      _Version        =   851968
      _ExtentX        =   20135
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
   End
   Begin MSComDlg.CommonDialog cdExcel 
      Left            =   4440
      Top             =   5160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Seleccione la hoja de Excel"
   End
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmExcel.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   13995
      TabIndex        =   0
      Top             =   7080
      Width           =   14000
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   10800
         TabIndex        =   1
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cargar Hoja"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmExcel.frx":50B3
         BorderGap       =   10
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   12240
         TabIndex        =   2
         Top             =   120
         Width           =   1455
         _Version        =   851968
         _ExtentX        =   2558
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmExcel.frx":B915
      End
      Begin VB.Label lblWGestion 
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
      Begin VB.Label lblWGestion 
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
   Begin XtremeSuiteControls.Label lblBuscar 
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6600
      Width           =   2000
      _Version        =   851968
      _ExtentX        =   3528
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Buscar en Articulos:"
      Transparent     =   -1  'True
   End
   Begin VB.OLE oleExcel 
      Height          =   6375
      Left            =   60
      OLEDropAllowed  =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   13695
   End
End
Attribute VB_Name = "frmExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vVieneDesdeExcel As String
Private Sub Form_Load()
On Error Resume Next

    vVieneDesdeExcel = ""


If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index
    
        Case 0
            With cdExcel
                .Filter = "Archivos MS Excel|*.xls"
                .DialogTitle = "Seleccione un archivo"
                
                .ShowOpen
                
                'Si seleccionamos un archivo mostramos la ruta
                If .FileName <> "" Then
                    
                    'MsgBox ""
                    
                    oleExcel.DisplayType = 0
                    oleExcel.Class = "Excel.Sheet.8"
                    oleExcel.SourceDoc = .FileName
                    oleExcel.CreateEmbed (oleExcel.SourceDoc)
                    
                    'oleExcel.SourceItem = .FileName
                    

    
                Else
                    'Si no mostramos un texto de advertencia de que no se seleccionó _
                    ninguno, ya que FileName devuelve una cadena vacía
                    MsgBox "No se seleccionó ningún archivo"

                End If
            
            End With
        
        Case 1
            Unload Me
    
    End Select

If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtBusqueda_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        Select Case vVieneDesdeExcel
            Case "Remito"
                           
                           
                           
            Case "Articulos"
                With frmArticulos
                    .Show
                    .txtBuscar.Text = Trim(txtBusqueda.Text)
                End With
            
            Case Else
                
        
        End Select

    End If

If Err Then GrabarLog "txtBusqueda_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
