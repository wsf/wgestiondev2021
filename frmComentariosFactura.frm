VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmComentariosFactura 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formulario de Seleccion de Comentarios para impresion"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmComentariosFactura.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   7005
      TabIndex        =   0
      Top             =   2760
      Width           =   7000
      Begin XtremeSuiteControls.PushButton cmdCerrar 
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   90
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmComentariosFactura.frx":50B3
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   375
         Left            =   3150
         TabIndex        =   5
         Top             =   90
         Width           =   2445
         _Version        =   851968
         _ExtentX        =   4313
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Nuevo Comentario"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmComentariosFactura.frx":54B3
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
         TabIndex        =   2
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
         TabIndex        =   3
         Top             =   170
         Width           =   1770
      End
   End
   Begin Grid.KlexGrid KlexComentariosFactura 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Documentos a cobrar"
      Top             =   120
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4471
      EnterKeyBehaviour=   0
      BackColorAlternate=   14737632
      GridLinesFixed  =   2
      AllowUserResizing=   1
      BackColor       =   16777215
      BackColorFixed  =   -2147483626
      Cols            =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridColorFixed  =   8421504
      MouseIcon       =   "frmComentariosFactura.frx":58B3
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmComentariosFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCerrar_Click()
On Error Resume Next

    Unload Me

If Err Then GrabarLog "cmdCerrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    Me.Show
    Call CargarComentariosFactura

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarComentariosFactura()
On Error Resume Next

    Dim rsComentariosFactura As New ADODB.Recordset, sqlComentariosFactura As String
    
    sqlComentariosFactura = "SELECT * FROM ComentariosFactura ORDER BY 1"
    
    With rsComentariosFactura
        .CursorLocation = adUseClient
        
        Call .Open(sqlComentariosFactura, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then .MoveFirst
        
        Call FormatoGrilla(.RecordCount)
        
        Do Until .EOF = True
            KlexComentariosFactura.TextMatrix(.AbsolutePosition, 1) = EsNulo(.Fields("idComentariosFactura").Value)
            KlexComentariosFactura.TextMatrix(.AbsolutePosition, 2) = EsNulo(.Fields("Comentario").Value)
            KlexComentariosFactura.TextMatrix(.AbsolutePosition, 3) = EsNulo(.Fields("Imprimir").Value)
            
            .MoveNext
        Loop
        
    End With
    
    sqlComentariosFactura = ""
    
    If rsComentariosFactura.State = 1 Then
        rsComentariosFactura.Close
        Set rsComentariosFactura = Nothing
    End If

If Err Then GrabarLog "CargarComentariosFactura", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FormatoGrilla(vCantidadRenglones As Integer)
On Error Resume Next

    Dim i As Integer

    With KlexComentariosFactura
        .FixedRows = 1
        .FixedCols = 1
    
        .Cols = 4
        .Rows = vCantidadRenglones + 1
        
        If vCantidadRenglones = 1 Then
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .ColWidth(i) = 0
            Next
        End If
        
        .TextMatrix(0, 0) = ""
        .ColWidth(0) = 400
        
        .TextMatrix(0, 1) = "ID"
        .ColWidth(1) = 750
               
        .TextMatrix(0, 2) = "Comentario"
        .ColWidth(2) = 4500
        
        .TextMatrix(0, 3) = "Imprimir"
        .ColWidth(3) = 750
        
    End With
    
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub KlexComentariosFactura_DblClick()
On Error Resume Next

    With KlexComentariosFactura
    
        If .TextMatrix(.Row, 3) = "N" Then
            Call EjecutarScript("UPDATE ComentariosFactura SET Imprimir = 'S' WHERE (idComentariosFactura = " & Val(.TextMatrix(.Row, 1)) & ")")
            .TextMatrix(.Row, 3) = "S"
        Else
            Call EjecutarScript("UPDATE ComentariosFactura SET Imprimir = 'N' WHERE (idComentariosFactura = " & Val(.TextMatrix(.Row, 1)) & ")")
            .TextMatrix(.Row, 3) = "N"
        End If
    
    End With

If Err Then GrabarLog "KlexComentariosFactura_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub

