VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAddCargacamion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Nuevo dato a la planilla de Carga de Camión..."
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   210
   ClientWidth     =   9060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9060
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkRecarga 
      Caption         =   "Recarga"
      Height          =   255
      Left            =   3120
      TabIndex        =   16
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtArticulo 
      Height          =   315
      Left            =   1440
      TabIndex        =   14
      Top             =   960
      Width           =   3705
   End
   Begin VB.TextBox txtRepartidor 
      Height          =   315
      Left            =   1440
      TabIndex        =   12
      Top             =   1680
      Width           =   3705
   End
   Begin VB.Frame fraRubro 
      Caption         =   "Rubro :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   30
      TabIndex        =   10
      Top             =   2520
      Width           =   3375
      Begin VB.Label lblRubro 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   60
         TabIndex        =   11
         Top             =   240
         Width           =   3165
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   525
      Left            =   4380
      Picture         =   "frmAddCargacamion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Imprimir"
      Top             =   2580
      UseMaskColor    =   -1  'True
      Width           =   915
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   525
      Left            =   3510
      Picture         =   "frmAddCargacamion.frx":0102
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ejecutar búsqueda"
      Top             =   2580
      UseMaskColor    =   -1  'True
      Width           =   885
   End
   Begin VB.TextBox txtComentario 
      Height          =   315
      Left            =   1440
      TabIndex        =   2
      Top             =   2040
      Width           =   3705
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Top             =   600
      Width           =   3705
   End
   Begin MSComCtl2.DTPicker vfecha 
      Height          =   315
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      Format          =   16711681
      CurrentDate     =   38573
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "> Artículo :"
      Height          =   195
      Left            =   0
      TabIndex        =   15
      Top             =   1000
      Width           =   1350
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "> Repartidor :"
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   1720
      Width           =   1350
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "> Comentario :"
      Height          =   195
      Left            =   30
      TabIndex        =   8
      Top             =   2080
      Width           =   1350
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "> Cantidad :"
      Height          =   195
      Left            =   0
      TabIndex        =   7
      Top             =   1360
      Width           =   1350
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "> Cód. Artículo :"
      Height          =   195
      Left            =   30
      TabIndex        =   6
      Top             =   640
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "> Fecha:"
      Height          =   195
      Left            =   30
      TabIndex        =   5
      Top             =   280
      Width           =   1350
   End
End
Attribute VB_Name = "frmAddCargacamion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()
    On Error Resume Next
    
    'frmCargaCamion.DgCargaCamion.Refresh
    Unload Me

    If Err Then GrabarLog "cmdCancelar_Click", Err.Number & "-" & Err.Description, Me.Name
End Sub
'Private Sub cmdAceptar_Click()
'    On Error Resume Next''

'    With frmCargaCamion.bcarga_camion

'        If Not chkRecarga.Value = 1 Then
'            .RecordSource = "SELECT * FROM Carga_Camion WHERE (Fecha = '" & strfechaMySQL(vfecha.Value) & "') AND (repartidor = '" + (txtRepartidor.Tag) + "') and (codigo = '" + txtCodigo.Text + "') order by fecha"
'            .Refresh

'            If .Recordset.RecordCount <> 0 Then
'                MsgBox "El articulo ya ha sido cargado en este reparto.", vbInformation, "Mensaje ..."
'                Exit Sub
'            Else

'                If Trim(txtArticulo) = "" Then Exit Sub
'                .Recordset.AddNew

'                .Recordset("fecha").Value = vfecha.Value
'                .Recordset("codigo").Value = txtCodigo.Text
'                .Recordset("cantidad").Value = Val(txtCantidad.Text)
'                .Recordset("articulo").Value = txtArticulo.Text
'                .Recordset("comentario").Value = txtComentario.Text
'                .Recordset("repartidor").Value = (txtRepartidor.Tag)
            
'            End If

'        Else
        
'            .RecordSource = "SELECT * FROM Carga_Camion WHERE (fecha = '" & strfechaMySQL(vfecha.Value) + "') AND (repartidor = '" + (txtRepartidor.Tag) + "') AND (codigo = '" + txtCodigo + "') ORDER BY fecha ASC"
'            .Refresh
'            If Not .Recordset.EOF = True Then .Recordset("recarga").Value = Val(txtCantidad)
'
'        End If

'        .Recordset.Update
'        Limpiar'

'    End With
    
'    If Err Then
'        GrabarLog "cmdAceptar_Click", Err.Number & "-" & Err.Description, Me.Name
'        MsgBox "Revise las operaciones", vbInformation, "Error"
'    End If
'End Sub
Private Sub Form_Load()
On Error Resume Next

    With Me
        .width = 5490
        .height = 3585
        .Top = 300
        .Left = 300
    End With
    
    vfecha = Date - 1

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub Limpiar()
On Error Resume Next

    txtCodigo.Text = ""
    txtArticulo.Text = ""
    txtCantidad.Text = ""
    txtComentario.Text = ""
    txtCodigo.SetFocus

If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtarticulo_keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        Dim rsArticulos As New ADODB.Recordset, sqlArticulos As String
        
        sqlArticulos = "SELECT * FROM Articulos WHERE (descrip like '%" + (txtArticulo.Text) + "%')"
        
        With rsArticulos
            .CursorLocation = adUseClient
            
            Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If .EOF Then
                frmBuscarArticulo.Show
                frmBuscarArticulo.o = 10
                frmBuscarArticulo.txtArticulo = txtArticulo.Text
                frmBuscarArticulo.txtArticulo.SetFocus
            Else
                txtCodigo.Text = .Fields("codigo").Value
                txtArticulo.Text = .Fields("descrip").Value
                txtCantidad.SetFocus
            End If
        
        End With
        
        sqlArticulos = ""
        
        If rsArticulos.State = 1 Then
            rsArticulos.Close
            Set rsArticulos = Nothing
        End If
    
    End If

If Err Then GrabarLog "txtarticulo_keypress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        txtRepartidor.SetFocus
    End If

If Err Then GrabarLog "txtCantidad_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtCodigo_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        Dim rsArticulos As New ADODB.Recordset, sqlArticulos As String
        
        sqlArticulos = "SELECT * FROM Articulos WHERE (codigo = '" + txtCodigo.Text + "')"
        
        With rsArticulos
            .CursorLocation = adUseClient
            
            Call .Open(sqlArticulos, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If .EOF Then
                frmBuscarArticulo.Show
                frmBuscarArticulo.o = 10
                frmBuscarArticulo.txtArticulo = txtArticulo.Text
                frmBuscarArticulo.txtArticulo.SetFocus
            Else
                txtCodigo.Text = .Fields("codigo").Value
                txtArticulo.Text = .Fields("descrip").Value
                txtCantidad.SetFocus
            End If
        
        End With
        
        sqlArticulos = ""
        
        If rsArticulos.State = 1 Then
            rsArticulos.Close
            Set rsArticulos = Nothing
        End If
    
    End If

If Err Then GrabarLog "txtCodigo_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtComentario_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        Me.cmdAceptar.SetFocus
    End If

If Err Then GrabarLog "txtComentario_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtRepartidor_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        
        Dim rsEmpleados As New ADODB.Recordset, sqlEmpleados As String
        
        sqlEmpleados = "SELECT * FROM Empleados WHERE (codigo LIKE '%" + Trim(txtRepartidor.Text) + "%') or (nombre LIKE '%" + Trim(txtRepartidor.Text) + "%')"
        
        With rsEmpleados
            .CursorLocation = adUseClient
            
            Call .Open(sqlEmpleados, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            If Not .EOF = True Then
                txtRepartidor.Text = .Fields("nombre").Value
                txtRepartidor.Tag = .Fields("codigo").Value
            Else
                txtRepartidor.Text = ""
                txtRepartidor.Tag = ""
            End If

            txtComentario.SetFocus
        End With
        
        sqlEmpleados = ""
        
        If rsEmpleados.State = 1 Then
            rsEmpleados.Close
            Set rsEmpleados = Nothing
        End If
        
    End If

If Err Then GrabarLog "txtRepartidor_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
