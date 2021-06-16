VERSION 5.00
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.TaskPanel.v13.0.0.Demo.ocx"
Begin VB.Form frmTaskPanel 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin XtremeTaskPanel.TaskPanel wndTaskPanel 
      Height          =   6165
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6945
      _Version        =   851968
      _ExtentX        =   12250
      _ExtentY        =   10874
      _StockProps     =   64
      ItemLayout      =   2
      HotTrackStyle   =   1
   End
End
Attribute VB_Name = "frmTaskPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    wndTaskPanel.LockRedraw = True
    wndTaskPanel.Groups.Add 1, Title & "Mantenimiento"
    wndTaskPanel.Groups(1).Items.Add 1, Title & "Nuevo registro", xtpTaskItemTypeText
    wndTaskPanel.Visible = True
    wndTaskPanel.LockRedraw = False
End Sub

Private Sub Form_Resize()
    wndTaskPanel.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

