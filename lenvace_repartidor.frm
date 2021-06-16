VERSION 5.00
Begin VB.Form frmListadoEnvaceRepartidor 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de envaces por repartidor y cliente ..."
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   240
   ClientWidth     =   5640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   825
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   5535
      Begin VB.TextBox txtCliente 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   300
         Width           =   4065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "> Cliente :"
         Height          =   195
         Left            =   30
         TabIndex        =   3
         Top             =   320
         Width           =   945
      End
   End
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "Ejecutar Listado"
      Height          =   375
      Left            =   3900
      TabIndex        =   0
      Top             =   900
      Width           =   1605
   End
End
Attribute VB_Name = "frmListadoEnvaceRepartidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEjecutar_Click()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox " Prepare la Impresora ", vbInformation, "Mensaje ..."

    With Mantenimiento.rsEnvaces_Repartidor

        If Not .State = 1 Then .Open
        .Close
        .Open
        
        If Not Trim(txtCliente.Text) = "" Then
            .filter = "Cod_Cliente = '" & Trim(txtCliente.Tag) & "'"
        End If

        .Sort = "Cod_empleado ASC"

    End With

    drEnvace_Repartidor.Show

If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next
        
    With Me
        .Width = 5805
        .Height = 1800
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
        Dim rsClientes As New ADODB.Recordset, sqlClientes As String
        
        sqlClientes = "SELECT * FROM Clientes WHERE (codigo = '" + txtCliente.Text + "') OR (nombre LIKE '%" + txtCliente.Text + "%')"

        With rsClientes
            .CursorLocation = adUseClient
            
            Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            
            If Not .EOF = True Then
                txtCliente.Text = .Fields("Nombre").Value
                txtCliente.Tag = .Fields("Codigo").Value
                cmdEjecutar.SetFocus
            Else
                MsgBox "El Cliente no fue encontrado.", vbInformation, "Mensaje ..."
            End If
        
        
        End With

    
    End If
    
    sqlClientes = ""
    
    If rsClientes.State = 1 Then
        rsClientes.Close
        Set rsClientes = Nothing
    End If
    

If Err Then GrabarLog "txtCliente_Keypress", Err.Number & " " & Err.Description, Me.Name
End Sub

