VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmComentarioR 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Comentario"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   3915
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc bclientes 
      Height          =   330
      Left            =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "bclientes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtComentario 
      Height          =   975
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label lblCantidad 
      AutoSize        =   -1  'True
      Caption         =   "> Repeticiones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   1125
      Width           =   1635
   End
End
Attribute VB_Name = "frmComentarioR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGuardar_Click()
On Error Resume Next
    

    With remito.bcliente
        
        .Recordset("Responsable") = Left(txtComentario.Text, 99)
        .Recordset("CComentario") = Val(txtCantidad.Text)
        .Recordset.Update
    
    End With
    
    Unload frmComentario
    
If Err Then graba_log "cmdGuardar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then SendKeys "{TAB}"

If Err Then graba_log "Form_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF4 Then cmdGuardar_Click
        

If Err Then graba_log "Form_KeyUp", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_Load()
On Error Resume Next

    KeyPreview = True
    With remito.bcliente
        
        If Not (IsNull(.Recordset("Responsable")) = True) Or Not (IsNull(.Recordset("CComentario")) = True) Then
            txtComentario.Text = .Recordset("Responsable").Value + ""
            txtCantidad.Text = .Recordset("CComentario").Value
        End If
        
    End With

If Err Then graba_log "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

