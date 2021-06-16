VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmLog 
   Caption         =   "Log"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   11130
   Begin VB.CommandButton Command1 
      Caption         =   "Excel"
      Height          =   375
      Left            =   10410
      TabIndex        =   2
      Top             =   6720
      Width           =   1305
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid log2 
      Height          =   6615
      Left            =   0
      TabIndex        =   1
      Top             =   30
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11668
      _Version        =   393216
      BackColor       =   4210752
      ForeColor       =   14737632
      Cols            =   1
      FixedRows       =   0
      FixedCols       =   0
      BackColorBkg    =   4210752
      GridColor       =   4210752
      WordWrap        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin XtremeSuiteControls.ListBox log 
      Height          =   225
      Left            =   0
      TabIndex        =   0
      Top             =   5340
      Width           =   10365
      _Version        =   851968
      _ExtentX        =   18283
      _ExtentY        =   397
      _StockProps     =   77
      BackColor       =   -2147483643
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call grillaToExcel(Me.log2)
End Sub

Private Sub Form_Load()

Me.Width = 11940
Me.Height = 7695

Me.log2.ColWidth(0) = 12000
End Sub
