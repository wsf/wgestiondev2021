VERSION 5.00
Begin VB.Form frmEnvasesAdeudados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de envases adeudados..."
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkAgrupar 
      Caption         =   "Agrupar por Artículo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   2910
      TabIndex        =   15
      Top             =   1980
      Width           =   2595
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Modificar Envases Iniciales"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   2340
      Width           =   1485
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Por Repartidor"
      Height          =   375
      Left            =   2850
      TabIndex        =   7
      Top             =   2460
      Width           =   1245
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Por clientes"
      Height          =   375
      Left            =   1620
      TabIndex        =   3
      Top             =   2460
      Width           =   1245
   End
   Begin VB.ComboBox vorden 
      Height          =   315
      ItemData        =   "Envases_Adeudados.frx":0000
      Left            =   2880
      List            =   "Envases_Adeudados.frx":000D
      TabIndex        =   13
      Text            =   "últimoDeNombre"
      Top             =   1590
      Width           =   2625
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Artículo"
      Height          =   255
      Left            =   1260
      TabIndex        =   11
      Top             =   2100
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Repartidor"
      Height          =   255
      Left            =   1260
      TabIndex        =   10
      Top             =   1830
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Cliente"
      Height          =   255
      Left            =   1260
      TabIndex        =   9
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Todos los Clientes"
      Height          =   375
      Left            =   30
      TabIndex        =   4
      Top             =   2460
      Width           =   1605
   End
   Begin VB.Frame fraClienteEmpleado 
      Height          =   1155
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtEmpleado 
         Height          =   285
         Left            =   1050
         TabIndex        =   5
         Top             =   720
         Width           =   4275
      End
      Begin VB.TextBox txtCliente 
         Height          =   285
         Left            =   1050
         TabIndex        =   1
         Top             =   300
         Width           =   4245
      End
      Begin VB.Label Label2 
         Caption         =   "> Repartidor: "
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   780
         Width           =   945
      End
      Begin VB.Label Label1 
         Caption         =   "> Cliente :"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   945
      End
   End
   Begin VB.Label lblEnvacesAdeudados 
      BackColor       =   &H00404040&
      Caption         =   "Ordenado por: "
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
      Height          =   255
      Index           =   1
      Left            =   2880
      TabIndex        =   12
      Top             =   1290
      Width           =   2595
   End
   Begin VB.Label lblEnvacesAdeudados 
      BackColor       =   &H00404040&
      Caption         =   "Agrupado por:"
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
      Height          =   225
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   1290
      Width           =   2595
   End
End
Attribute VB_Name = "frmEnvasesAdeudados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()

If chkAgrupar.Value = 1 Then

MsgBox "No se pueden listar agrupados por artículos y por clientes", vbCritical, "Error..."
Exit Sub


    If Not Mantenimiento.rsDevol.State = 1 Then
        Mantenimiento.rsDevol_articulo.Open
        Mantenimiento.rsDevol_articulo.Close
        Mantenimiento.rsDevol_articulo.Open
    Else
        Mantenimiento.rsDevol_articulo.Close
        Mantenimiento.rsDevol_articulo.Open
    End If


    If txtCliente.Text = "" Then
        Mantenimiento.rsDevol_articulo.filter = "Mora <> 0 and envase = true "
    Else
        Mantenimiento.rsDevol_articulo.filter = "clientes.codigo ='" + Trim(txtCliente.Tag) + "'and Mora <> 0 and envase = true"
    End If

    '"cod_Repartidor ='" + Trim(vcod_repartidor) + "' and últimodefecha >= '" + Str(fdesde) + "' and últimodefecha <= '" + Str(fhasta) + "'"
    Mantenimiento.rsDevol_articulo.Sort = Trim(vorden)
    drEnvases_Adeudados_Articulos.Sections("TituloEmpresa").Controls("vcliente").Caption = txtCliente.Text
    drEnvases_Adeudados_Articulos.Show
Else

    If Not Mantenimiento.rsDevol.State = 1 Then
        Mantenimiento.rsDevol.Open
        Mantenimiento.rsDevol.Close
        Mantenimiento.rsDevol.Open
    Else
        Mantenimiento.rsDevol.Close
        Mantenimiento.rsDevol.Open
    End If


    If txtCliente.Text = "" Then
        Mantenimiento.rsDevol.filter = "Mora <> 0 and envase = true "
    Else
        Mantenimiento.rsDevol.filter = "clientes.codigo ='" + Trim(txtCliente.Tag) + "'and Mora <> 0 and envase = true"
    End If

    '"cod_Repartidor ='" + Trim(vcod_repartidor) + "' and últimodefecha >= '" + Str(fdesde) + "' and últimodefecha <= '" + Str(fhasta) + "'"
    Mantenimiento.rsDevol.Sort = Trim(vorden)
    drEnvases_Adeudados.Sections("TituloEmpresa").Controls("vcliente").Caption = txtCliente.Text
    drEnvases_Adeudados.Show

End If


End Sub

Private Sub Command2_Click()
    MsgBox "Prepare la impresora", vbInformation, "Mensaje..."
    
    If Not Mantenimiento.rsDevol.State = 1 Then
        Mantenimiento.rsDevol.Open
        Mantenimiento.rsDevol.Close
        Mantenimiento.rsDevol.Open
    Else
        Mantenimiento.rsDevol.Close
        Mantenimiento.rsDevol.Open
    End If
    
    
    Mantenimiento.rsDevol.filter = "Mora <> 0  and envase = true"
    Mantenimiento.rsDevol.Sort = Trim(vorden)
    drEnvases_Adeudados.Sections("TituloEmpresa").Controls("vcliente").Caption = txtCliente.Text
    drEnvases_Adeudados.Show
End Sub

Private Sub Command3_Click()
     
If chkAgrupar.Value = 1 Then
    
        
    If Not Mantenimiento.rsDevol_articulo.State = 1 Then
        Mantenimiento.rsDevol_articulo.Open
        Mantenimiento.rsDevol_articulo.Close
        Mantenimiento.rsDevol_articulo.Open
    Else
        Mantenimiento.rsDevol_articulo.Close
        Mantenimiento.rsDevol_articulo.Open
    End If
     
    Mantenimiento.rsDevol_articulo.filter = "Repartidor ='" + Trim(txtEmpleado.Tag) + "' and Mora <> 0 and envase = true"
    Mantenimiento.rsDevol_articulo.Sort = Trim(vorden)
 
    drEnvases_Adeudados_Articulos.Sections("TituloEmpresa").Controls("vcliente").Caption = txtEmpleado.Text
    drEnvases_Adeudados_Articulos.Show


Else

    If Not Mantenimiento.rsDevol.State = 1 Then
        Mantenimiento.rsDevol.Open
        Mantenimiento.rsDevol.Close
        Mantenimiento.rsDevol.Open
    Else
        Mantenimiento.rsDevol.Close
        Mantenimiento.rsDevol.Open
    End If
     
    Mantenimiento.rsDevol.filter = "últimoDeRepartidor ='" + Trim(txtEmpleado.Tag) + "' and Mora <> 0 and envase = true"
    Mantenimiento.rsDevol.Sort = Trim(vorden)
 
    drEnvases_Adeudados.Sections("TituloEmpresa").Controls("vcliente").Caption = txtEmpleado.Text
    drEnvases_Adeudados.Show
End If
End Sub

Private Sub Command4_Click()
    
    frmEnvasesIniciales.Show

End Sub
Private Sub Form_Load()
On Error Resume Next

    With Me
        .Width = 5805
        .Height = 3495
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtCliente_Keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
    
        Dim rsClientes As New ADODB.Recordset
        Dim sqlClientes As String
    
        sqlClientes = "SELECT * FROM Clientes WHERE (codigo = '" + txtCliente.Text + "') OR (nombre LIKE '%" + txtCliente.Text + "%')"
    
        With rsClientes
            If .State = 1 Then .Close
    
            .CursorLocation = adUseClient
        
            Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
        
            If Not .EOF = True Then
                txtCliente.Text = .Fields("nombre").Value
                txtCliente.Text = .Fields("codigo").Value
                Command1.SetFocus
            Else
                txtCliente.Text = ""
                txtCliente.Tag = ""
                MsgBox "El Cliente no fue encontrado.", vbInformation, "Mensaje ..."
            End If
        
        End With
    
        sqlClientes = ""
    
        If rsClientes.State = 1 Then
            rsClientes.Close
            Set rsClientes = Nothing
        End If
    
    End If
    
If Err Then GrabarLog "txtCliente_Keypress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtEmpleado_KeyPress(KeyAscii As Integer)
    On Error Resume Next

    If KeyAscii = 13 Then
        
        Dim rsempleados As New ADODB.Recordset
        Dim sqlEmpleados As String
        
        sqlEmpleados = "SELECT * FROM empleados WHERE (codigo = '" + txtEmpleado.Text + "') OR (nombre LIKE '%" + txtEmpleado.Text + "%')"
    
        With rsempleados
            If .State = 1 Then .Close
    
            .CursorLocation = adUseClient
        
            Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockReadOnly)
        
            If Not .EOF = True Then
                txtEmpleado.Text = .Fields("nombre").Value
                txtEmpleado.Text = .Fields("codigo").Value
                Command1.SetFocus
            Else
                txtEmpleado.Text = ""
                txtEmpleado.Tag = ""
                MsgBox "El Empleado no fue encontrado.", vbInformation, "Mensaje ..."
            End If
        
        End With
    
        sqlEmpleados = ""
    
        If rsempleados.State = 1 Then
            rsempleados.Close
            Set rsempleados = Nothing
        End If
    End If

If Err Then GrabarLog "txtEmpleado_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub


