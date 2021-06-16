VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTextilArt 
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   6300
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   8235
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   8235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      Caption         =   "Tales 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5460
      Left            =   5940
      TabIndex        =   9
      Top             =   180
      Width           =   2175
      Begin VB.TextBox TxtCantidad3 
         BackColor       =   &H0080FFFF&
         Height          =   330
         Left            =   2115
         TabIndex        =   13
         Top             =   720
         Width           =   1410
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2610
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   675
         Width           =   1770
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Talle3 
         Height          =   5100
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   8996
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000C&
      Caption         =   "Tales 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5460
      Left            =   3105
      TabIndex        =   6
      Top             =   180
      Width           =   2175
      Begin VB.TextBox TxtCantidad2 
         BackColor       =   &H0080FFFF&
         Height          =   330
         Left            =   2115
         TabIndex        =   12
         Top             =   720
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2610
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   675
         Width           =   1770
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Talle2 
         Height          =   5100
         Left            =   90
         TabIndex        =   8
         Top             =   270
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   8996
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox vdescrip 
      Height          =   330
      Left            =   945
      TabIndex        =   5
      Top             =   5850
      Width           =   6315
   End
   Begin VB.Frame FraTales1 
      BackColor       =   &H8000000C&
      Caption         =   "Tales 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5460
      Left            =   135
      TabIndex        =   2
      Top             =   180
      Width           =   2220
      Begin VB.TextBox TxtCantidad 
         BackColor       =   &H0080FFFF&
         Height          =   285
         Left            =   2700
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   675
         Width           =   1680
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Talles 
         Height          =   5100
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   8996
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7380
      TabIndex        =   1
      Top             =   5850
      Width           =   720
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   5805
      Width           =   720
   End
End
Attribute VB_Name = "frmTextilArt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim ConGrilla As Boolean
Dim vc As Double

Private Sub d1_Click()


End Sub

Private Sub CancelButton_Click()
frmRemito.txtDetalle(1).Text = frmRemito.txtDetalle(1).Text + " - " + Me.vdescrip
frmRemito.txtDetalle(0).Text = vc
frmRemito.txtDetalle(2).SetFocus
Unload Me
End Sub

Private Sub Form_Load()

init1

End Sub


Private Sub init1()

Dim i As Integer

Talles.Rows = 23
Talles.Cols = 2


Talles.ColWidth(0) = 500
Talles.ColWidth(1) = 500

For i = 1 To 22
    Talles.TextMatrix(i, 0) = i
Next

Dim ii As Integer

Talle2.Rows = 15
Talle2.Cols = 2

ii = 0


Talle2.ColWidth(0) = 500
Talle2.ColWidth(1) = 500


For i = 26 To 50 Step 2
    ii = ii + 1
    Talle2.TextMatrix(ii, 0) = i
Next


Talle3.ColWidth(0) = 500
Talle3.ColWidth(1) = 500

Talle3.Rows = 7
Talle3.Cols = 2


Talle3.TextMatrix(1, 0) = "XS"
Talle3.TextMatrix(2, 0) = "S"
Talle3.TextMatrix(3, 0) = "M"
Talle3.TextMatrix(4, 0) = "L"
Talle3.TextMatrix(5, 0) = "XL"
Talle3.TextMatrix(6, 0) = "XXL"

End Sub


Private Sub OKButton_Click()
Dim f1, f2, f3 As String

vdescrip.Text = ""

vc = 0

f1 = formar1()
f2 = formar2()
f3 = formar3()

vdescrip.Text = f1 + f2 + f3

End Sub


Function formar1()
Dim i As Integer
Dim v As String

v = ""

For i = 1 To Talles.Row
    If Not Val(Talles.TextMatrix(i, 1)) = 0 Then
        v = v + "T" + Talles.TextMatrix(i, 0) + ":" + Talles.TextMatrix(i, 1) + ", "
        vc = vc + Val(Talles.TextMatrix(i, 1))
    End If
Next

formar1 = v

End Function


Function formar2()
Dim i As Integer
Dim v As String

v = ""

For i = 1 To Talle2.Row
    If Not Val(Talle2.TextMatrix(i, 1)) = 0 Then
        v = v + "T" + Talle2.TextMatrix(i, 0) + ":" + Talle2.TextMatrix(i, 1) + ", "
        vc = vc + Val(Talle2.TextMatrix(i, 1))
    End If
Next

formar2 = v

End Function



Function formar3()
Dim i As Integer
Dim v As String

v = ""

For i = 1 To Talle3.Row
    If Not Val(Talle3.TextMatrix(i, 1)) = 0 Then
        v = v + Talle3.TextMatrix(i, 0) + ":" + Talle3.TextMatrix(i, 1) + ", "
        vc = vc + Val(Talle3.TextMatrix(i, 1))
    End If
Next

formar3 = v

End Function


Private Sub Talles_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim i As Integer
Select Case KeyCode
   Case 13:
      
      If Talles.Row <> Talles.Rows - 1 And Shift = 1 Then
         i = Talles.Row
         Talles.TextMatrix(Talles.Row, 8) = Val(Talles.TextMatrix(Talles.Row, 6)) * Val(Talles.TextMatrix(Talles.Row, 7))
         'Talles.Row = Talles.Row - 1
               
         'SendKeys ("{RIGHT}")
       
         'calcular_ptotales
         Talles.SetFocus
         Talles.TopRow = i
         Talles.Row = i + 1
         Talles.Col = 6
         'Talles.Row = Talles.Row - 1
         Talles.SetFocus
         
         Else
        
       ' para que no salte en el primer foco
         Talles.TopRow = Talles.Row
         Talles.Row = Talles.Row + 1
         Talles.Col = 6
        
         Talles.SetFocus
              
      End If
           

   Case 116:
      'Call BtnCargar_Click
   Case 117:
      'Call BtnNuevo_Click
End Select
If Err Then Exit Sub
End Sub

Private Sub Talles_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim c As String

If Talles.Col = 1 And KeyAscii <> 13 Then
   Call ValidarNumero(Talles, KeyAscii)
   
   c = Chr(KeyAscii)
   
   'id = Talles.TextMatrix(Talles.Row, 1)
         
   TxtCantidad.Width = Talles.ColWidth(6) + 10
   TxtCantidad.Height = Talles.RowHeight(Talles.Row)
   TxtCantidad.Left = Talles.CellLeft + Talles.Left - 30
   TxtCantidad.Top = Talles.CellTop + Talles.Top - 30
   
   TxtCantidad.Visible = True
   TxtCantidad.SetFocus
   TxtCantidad.Text = c
End If
If Err Then Exit Sub
End Sub

Private Sub TxtCantidad_GotFocus()
TxtCantidad.SelStart = Len(TxtCantidad)
TxtCantidad.SelLength = Len(TxtCantidad)
End Sub

Private Sub TxtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13:
      Talles.TextMatrix(Talles.Row, 1) = TxtCantidad
      TxtCantidad.Text = ""
      TxtCantidad.Visible = False
      
     ' Talles.SetFocus
     ' Talles.Row = Talles.Row + 1
      'Talles.SetFocus
      'SendKeys ("{DOWN}")
   Case 27:
      TxtCantidad.Text = ""
      TxtCantidad.Visible = False
      Talles.SetFocus
   Case 38:
      Talles.TextMatrix(Talles.Row, 6) = TxtCantidad
      
      TxtCantidad.Text = ""
      TxtCantidad.Visible = False
      Talles.SetFocus
      
      If Talles.Row <> 0 Then
         Talles.Row = Talles.Row - 1
         Call Talles_KeyDown(13, 0)
      End If
   Case 40:
      Talles.TextMatrix(Talles.Row, 1) = TxtCantidad
      
      TxtCantidad.Text = ""
      TxtCantidad.Visible = False
      Talles.SetFocus
      
      If Talles.Row <> Talles.Rows - 2 Then
         Talles.Row = Talles.Row + 1
         Call Talles_KeyDown(13, 0)
      End If
End Select
End Sub

Private Sub TxtCantidad_KeyPress(KeyAscii As Integer)
Call ValidarNumero(TxtCantidad, KeyAscii)
End Sub

Private Sub TxtCantidad_LostFocus()
TxtCantidad.Text = ""
TxtCantidad.Visible = False
Talles.SetFocus

Call Talles_KeyDown(13, 1)
End Sub




'---------------------------------------------------
'--------------------------------------------------
'--------------------------------------------------



Private Sub Talle2_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim i As Integer
Select Case KeyCode
   Case 13:
      
      If Talle2.Row <> Talle2.Rows - 1 And Shift = 1 Then
         i = Talle2.Row
         Talle2.TextMatrix(Talle2.Row, 8) = Val(Talle2.TextMatrix(Talle2.Row, 6)) * Val(Talle2.TextMatrix(Talle2.Row, 7))
         'Talle2.Row = Talle2.Row - 1
               
         'SendKeys ("{RIGHT}")
       
         'calcular_ptotales
         Talle2.SetFocus
         Talle2.TopRow = i
         Talle2.Row = i + 1
         Talle2.Col = 6
         'Talle2.Row = Talle2.Row - 1
         Talle2.SetFocus
         
         Else
        
       ' para que no salte en el primer foco
         Talle2.TopRow = Talle2.Row
         Talle2.Row = Talle2.Row + 1
         Talle2.Col = 6
        
         Talle2.SetFocus
              
      End If
           

   Case 116:
      'Call BtnCargar_Click
   Case 117:
      'Call BtnNuevo_Click
End Select
If Err Then Exit Sub
End Sub

Private Sub Talle2_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim c As String

If Talle2.Col = 1 And KeyAscii <> 13 Then
   Call ValidarNumero(Talle2, KeyAscii)
   
   c = Chr(KeyAscii)
   
   'id = Talle2.TextMatrix(Talle2.Row, 1)
         
   TxtCantidad2.Width = Talle2.ColWidth(6) + 10
   TxtCantidad2.Height = Talle2.RowHeight(Talle2.Row)
   TxtCantidad2.Left = Talle2.CellLeft + Talle2.Left - 30
   TxtCantidad2.Top = Talle2.CellTop + Talle2.Top - 30
   
   TxtCantidad2.Visible = True
   TxtCantidad2.SetFocus
   TxtCantidad2.Text = c
End If
If Err Then Exit Sub
End Sub

Private Sub TxtCantidad2_GotFocus()
TxtCantidad2.SelStart = Len(TxtCantidad2)
TxtCantidad2.SelLength = Len(TxtCantidad2)
End Sub

Private Sub TxtCantidad2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13:
      Talle2.TextMatrix(Talle2.Row, 1) = TxtCantidad2
      TxtCantidad2.Text = ""
      TxtCantidad2.Visible = False
      
     ' Talle2.SetFocus
     ' Talle2.Row = Talle2.Row + 1
      'Talle2.SetFocus
      'SendKeys ("{DOWN}")
   Case 27:
      TxtCantidad2.Text = ""
      TxtCantidad2.Visible = False
      Talle2.SetFocus
   Case 38:
      Talle2.TextMatrix(Talle2.Row, 6) = TxtCantidad2
      
      TxtCantidad2.Text = ""
      TxtCantidad2.Visible = False
      Talle2.SetFocus
      
      If Talle2.Row <> 0 Then
         Talle2.Row = Talle2.Row - 1
         Call Talle2_KeyDown(13, 0)
      End If
   Case 40:
      Talle2.TextMatrix(Talle2.Row, 6) = TxtCantidad2
      
      TxtCantidad2.Text = ""
      TxtCantidad2.Visible = False
      Talle2.SetFocus
      
      If Talle2.Row <> Talle2.Rows - 2 Then
         Talle2.Row = Talle2.Row + 1
         Call Talle2_KeyDown(13, 0)
      End If
End Select
End Sub

Private Sub TxtCantidad2_KeyPress(KeyAscii As Integer)
Call ValidarNumero(TxtCantidad2, KeyAscii)
End Sub

Private Sub TxtCantidad2_LostFocus()
TxtCantidad2.Text = ""
TxtCantidad2.Visible = False
Talle2.SetFocus

Call Talle2_KeyDown(13, 1)
End Sub

' ----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
' ----------------------------------------------------------------------------


Private Sub Talle3_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Dim i As Integer
Select Case KeyCode
   Case 13:
      
      If Talle3.Row <> Talle3.Rows - 1 And Shift = 1 Then
         i = Talle3.Row
         Talle3.TextMatrix(Talle3.Row, 8) = Val(Talle3.TextMatrix(Talle3.Row, 6)) * Val(Talle3.TextMatrix(Talle3.Row, 7))
         'Talle3.Row = Talle3.Row - 1
               
         'SendKeys ("{RIGHT}")
       
         'calcular_ptotales
         Talle3.SetFocus
         Talle3.TopRow = i
         Talle3.Row = i + 1
         Talle3.Col = 6
         'Talle3.Row = Talle3.Row - 1
         Talle3.SetFocus
         
         Else
        
       ' para que no salte en el primer foco
         Talle3.TopRow = Talle3.Row
         Talle3.Row = Talle3.Row + 1
         Talle3.Col = 6
        
         Talle3.SetFocus
              
      End If
           

   Case 116:
      'Call BtnCargar_Click
   Case 117:
      'Call BtnNuevo_Click
End Select
If Err Then Exit Sub
End Sub

Private Sub Talle3_KeyPress(KeyAscii As Integer)
On Error Resume Next
Dim c As String

If Talle3.Col = 1 And KeyAscii <> 13 Then
   Call ValidarNumero(Talle3, KeyAscii)
   
   c = Chr(KeyAscii)
   
   'id = Talle3.TextMatrix(Talle3.Row, 1)
         
   TxtCantidad3.Width = Talle3.ColWidth(6) + 10
   TxtCantidad3.Height = Talle3.RowHeight(Talle3.Row)
   TxtCantidad3.Left = Talle3.CellLeft + Talle3.Left - 30
   TxtCantidad3.Top = Talle3.CellTop + Talle3.Top - 30
   
   TxtCantidad3.Visible = True
   TxtCantidad3.SetFocus
   TxtCantidad3.Text = c
End If
If Err Then Exit Sub
End Sub

Private Sub TxtCantidad3_GotFocus()
TxtCantidad3.SelStart = Len(TxtCantidad3)
TxtCantidad3.SelLength = Len(TxtCantidad3)
End Sub

Private Sub TxtCantidad3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case 13:
      Talle3.TextMatrix(Talle3.Row, 1) = TxtCantidad3
      TxtCantidad3.Text = ""
      TxtCantidad3.Visible = False
      
     ' Talle3.SetFocus
     ' Talle3.Row = Talle3.Row + 1
      'Talle3.SetFocus
      'SendKeys ("{DOWN}")
   Case 27:
      TxtCantidad3.Text = ""
      TxtCantidad3.Visible = False
      Talle3.SetFocus
   Case 38:
      Talle3.TextMatrix(Talle3.Row, 6) = TxtCantidad3
      
      TxtCantidad3.Text = ""
      TxtCantidad3.Visible = False
      Talle3.SetFocus
      
      If Talle3.Row <> 0 Then
         Talle3.Row = Talle3.Row - 1
         Call Talle3_KeyDown(13, 0)
      End If
   Case 40:
      Talle3.TextMatrix(Talle3.Row, 6) = TxtCantidad3
      
      TxtCantidad3.Text = ""
      TxtCantidad3.Visible = False
      Talle3.SetFocus
      
      If Talle3.Row <> Talle3.Rows - 2 Then
         Talle3.Row = Talle3.Row + 1
         Call Talle3_KeyDown(13, 0)
      End If
End Select
End Sub

Private Sub TxtCantidad3_KeyPress(KeyAscii As Integer)
Call ValidarNumero(TxtCantidad3, KeyAscii)
End Sub

Private Sub TxtCantidad3_LostFocus()
TxtCantidad3.Text = ""
TxtCantidad3.Visible = False
Talle3.SetFocus

Call Talle3_KeyDown(13, 1)
End Sub


