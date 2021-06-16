VERSION 5.00
Begin VB.Form frmMigraFac 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Modulo de Migración de Facturas y Detalles desde Temporal"
   ClientHeight    =   7095
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7095
   ScaleWidth      =   14820
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraSucesos 
      Caption         =   "Log de Sucesos"
      Height          =   1935
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   9015
      Begin VB.ListBox display 
         Height          =   1425
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   8775
      End
   End
End
Attribute VB_Name = "frmMigraFac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private Sub Form_Load()

End Sub
