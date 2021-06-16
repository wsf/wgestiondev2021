VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmGastosConcepto 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filtrar por fechas..."
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton cmdEjecutar 
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
      _Version        =   851968
      _ExtentX        =   2143
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Ver Listado"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame fraFecha 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CheckBox chkFecha 
         Caption         =   "Anular fechas"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Value           =   1  'Checked
         Width           =   1305
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   285
         Index           =   0
         Left            =   1320
         TabIndex        =   1
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   38023
      End
      Begin MSComCtl2.DTPicker dtpFecha 
         Height          =   285
         Index           =   1
         Left            =   3600
         TabIndex        =   2
         Top             =   180
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Format          =   16842753
         CurrentDate     =   38023
      End
      Begin VB.Label lblFecha 
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblFecha 
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2790
         TabIndex        =   3
         Top             =   180
         Width           =   765
      End
   End
End
Attribute VB_Name = "frmGastosConcepto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkFecha_Click()
On Error Resume Next

    dtpFecha(0).Enabled = Not CBool(chkFecha.Value)
    dtpFecha(1).Enabled = Not CBool(chkFecha.Value)
    
If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdEjecutar_Click()
On Error Resume Next
    
    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox " Prepare la Impresora ", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsGastos_Concepto
        
        If .State = 0 Then .Open
        .Close
        .Open
    
        If chkFecha.Value = 1 Then
            .filter = "(Retiro > 0"
        Else
            .filter = "(Retiro > 0) AND (fecha >= '" & strfechaMySQL(dtpFecha(0).Value) & "' and fecha <= '" & strfechaMySQL(dtpFecha(1).Value) & "')"
        End If
    
    End With

    With drGastos_Concepto
        .Show
    End With

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next

    With Me
        .Height = 2010
        .Width = 5730
    
        .KeyPreview = True
    End With
    
    dtpFecha(0).Value = Date
    dtpFecha(1).Value = Date
    dtpFecha(0).Enabled = False
    dtpFecha(1).Enabled = False
    chkFecha.Value = False

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
