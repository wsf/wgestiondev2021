VERSION 5.00
Begin VB.Form frmClienteInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Información del cliente..."
   ClientHeight    =   1530
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1530
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   Begin VB.Label vcredito 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   7
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Crédito :"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label vsaldo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   5
      Top             =   780
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Saldo :"
      Height          =   255
      Left            =   60
      TabIndex        =   4
      Top             =   780
      Width           =   1335
   End
   Begin VB.Label vupago 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   3
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Último Pago :"
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.Label vuventa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1470
      TabIndex        =   1
      Top             =   180
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Última Venta :"
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   180
      Width           =   1335
   End
End
Attribute VB_Name = "frmClienteInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub foco()
On Error Resume Next
    
    vupago.Caption = gupago
    vuventa.Caption = guventa
    vsaldo.Caption = gsaldo
    vcredito.Caption = gcredito
    Me.Show
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    With Me
        .Left = 10900
        .Height = 1905
        .Width = 3795
    End With

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub

