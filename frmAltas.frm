VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmAltas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   -30
      Picture         =   "frmAltas.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   7125
      TabIndex        =   11
      Top             =   3120
      Width           =   7125
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   4560
         TabIndex        =   5
         Top             =   105
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAltas.frx":50B3
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   5640
         TabIndex        =   6
         Top             =   105
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmAltas.frx":54BA
      End
      Begin VB.Label lblWGestion 
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
         TabIndex        =   13
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGestion 
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
         TabIndex        =   12
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   3255
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   6825
      _Version        =   851968
      _ExtentX        =   12039
      _ExtentY        =   5741
      _StockProps     =   68
      Appearance      =   11
      Color           =   128
      ItemCount       =   1
      Item(0).Caption =   "Ingreso de valores de alta"
      Item(0).ControlCount=   10
      Item(0).Control(0)=   "txtAlta(0)"
      Item(0).Control(1)=   "txtAlta(1)"
      Item(0).Control(2)=   "txtAlta(2)"
      Item(0).Control(3)=   "lblAlta(2)"
      Item(0).Control(4)=   "lblAlta(1)"
      Item(0).Control(5)=   "lblAlta(0)"
      Item(0).Control(6)=   "txtAlta(3)"
      Item(0).Control(7)=   "txtAlta(4)"
      Item(0).Control(8)=   "lblAlta(3)"
      Item(0).Control(9)=   "lblAlta(4)"
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   0
         Left            =   2760
         TabIndex        =   0
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   1
         Top             =   960
         Visible         =   0   'False
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   2
         Left            =   2760
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   3
         Left            =   2760
         TabIndex        =   3
         Top             =   1680
         Visible         =   0   'False
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   4
         Left            =   2760
         TabIndex        =   4
         Top             =   2040
         Visible         =   0   'False
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   15
         Top             =   2080
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   3
         Left            =   510
         TabIndex        =   14
         Top             =   1710
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   10
         Top             =   645
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   9
         Top             =   1000
         Visible         =   0   'False
         Width           =   2200
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   8
         Top             =   1360
         Visible         =   0   'False
         Width           =   2200
      End
   End
End
Attribute VB_Name = "frmAltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const vHeight = 250
Public vcampos As Integer
Public vaccion As String
Private Sub Form_Load()
On Error Resume Next
    
    Me.Show
    ArmarFormulario

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ArmarFormulario()
On Error Resume Next
    
    Dim i As Integer
    
    TabAlta.Height = txtAlta(vcampos).Height + txtAlta(vcampos).Top + vHeight
    PicInferior.Top = TabAlta.Height + vHeight
    Height = PicInferior.Top + PicInferior.Height + (vHeight * 2)
    
    For i = Val(vcampos + 1) To txtAlta.Count - 1
        txtAlta(i).Visible = Not True
        lblAlta(i).Visible = Not True
    Next
    
    Select Case vVieneBusqueda
    
        Case "CodigoPostal"
            Caption = "Nueva Localidad"
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo Postal:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Localidad:"
            
            txtAlta(2).Visible = True
            lblAlta(2).Visible = True
            lblAlta(2).Caption = "Provincia:"
            
        Case "Vendedor"
            Caption = "Nuevo Vendedor"
        
    
        Case "Reparto"
            Caption = "Nuevo Reparto"
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Descripcion:"
        
        Case "TipoMovimientosBanco"
            Caption = "Nuevo Tipo Movimiento"
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "> Codigo:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "> Descripcion:"
            
            
            txtAlta(2).Visible = True
            lblAlta(2).Visible = True
            lblAlta(2).Caption = "> (P)ro (C)li (B)an"
            txtAlta(2).Text = ""
            
            
            txtAlta(3).Visible = True
            lblAlta(3).Visible = True
            lblAlta(3).Caption = "> (I)ng.(E)gr."
            txtAlta(3).MaxLength = 1
            txtAlta(3).Text = ""
            
            
            txtAlta(4).Visible = True
            lblAlta(4).Visible = True
            lblAlta(4).Caption = "> (S)i (N)o"
            txtAlta(4).MaxLength = 1
            txtAlta(4).Text = "S"
            
            
        
        Case "TipoIva"
            Caption = "Nuevo Tipo de Iva"
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Tipo de Iva:"

        Case "TipoCliente"
            Caption = "Nuevo Tipo de Cliente"
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Tipo de Cliente:"

        Case "Actividad"
            Caption = "Nueva Actividad"
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Actividad:"

        Case "Lista"
            Caption = "Nueva Lista de Precio"
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Lista de Precio:"

        Case "EstadoCliente"
            Caption = "Nuevo Estado del Cliente"
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Estado del Cliente:"
                        
            txtAlta(2).Visible = True
            lblAlta(2).Visible = True
            lblAlta(2).Caption = "Permite Facturar"
            txtAlta(2).MaxLength = 1
            txtAlta(2).Text = "S"
            
            txtAlta(3).Visible = True
            lblAlta(3).Visible = True
            lblAlta(3).Caption = "Aparece en Listados:"
            txtAlta(3).MaxLength = 1
            txtAlta(3).Text = "S"

        Case "Rubro"
            Caption = "Nueva Ficha de Rubro"
            
            txtAlta(0).Text = Val(GenerarDato("SELECT MAX(idRubros) AS UltimoCodigo FROM Rubros", "UltimoCodigo")) + 1
            txtAlta(0).Text = FormatoUltimoCodigo(3, txtAlta(0).Text)
            
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo:"
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Descripcion de Rubro:"
            
        Case "SubRubro"
            Caption = "Nueva Ficha de Sub-Rubro"
            
            txtAlta(0).Text = Val(GenerarDato("SELECT MAX(idSubRubros) AS UltimoCodigo FROM SubRubros", "UltimoCodigo")) + 1
            txtAlta(0).Text = FormatoUltimoCodigo(3, txtAlta(0).Text)
            
            txtAlta(0).Visible = True
            lblAlta(0).Visible = True
            lblAlta(0).Caption = "Codigo:"
            
            
            
            txtAlta(1).Visible = True
            lblAlta(1).Visible = True
            lblAlta(1).Caption = "Descripcion de Rubro:"
    
            
    End Select
    
    txtAlta(0).SetFocus
    
If Err Then GrabarLog "ArmarFormulario", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index
    
        Case 0
            Grabar
        Case 1
            Unload Me
    End Select
    
    frmBusqueda.CargarRegistros

If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Grabar()
    On Error Resume Next

    If Not ValidarCampos() = True Then
        Exit Sub
    End If
    
    Dim rsAlta As New ADODB.Recordset, sqlAlta As String
    
    Select Case Me.vaccion

        Case "Nuevo"
            'sqlAlta = "SELECT * FROM " & vGrabarTabla & " WHERE 1=2"
            sqlAlta = "SELECT * FROM " & vGrabarTabla
        
        Case "Modificar"
            'sqlAlta = "SELECT * FROM " & vGrabarTabla & " WHERE (Codigo = '" & Trim(txtAlta(0).Text) & "')"
        
        Case "Duplicar"
            
    End Select
        
    With rsAlta
    
        Call .Open(sqlAlta, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .State = 0 Then
        
            Select Case Me.vaccion
            
                Case "Nuevo"
                    .AddNew
                
                Case "Modificar"
                    'No hago nada
                    
                Case "Duplicar"
                    .AddNew
                    '.Fields("Codigo").Value = "" 'Tendria que traer el ultimo codigo
                    '.Fields("codigo_num").Value = Val(txtAlta(0).Text)

            End Select
            
            
           If vVieneBusqueda = "TipoMovimientosBanco" Then vVieneBusqueda = "TMB" ' Alfredo: hago esto porque el case que sigue no me toma la comparacion "TipoMovimientosBanco"
            
            Select Case vVieneBusqueda
    
                Case "CodigoPostal"
                    .Fields(0).Value = txtAlta(0).Text
                    .Fields(1).Value = Left(txtAlta(1).Text, 255)
                    .Fields(2).Value = Left(txtAlta(2).Text, 255)
                
                Case "Vendedor"
                    .Fields("Codigo").Value = txtAlta(0).Text
                    .Fields("Nombre").Value = Left(txtAlta(1).Text, 255)

                Case "Reparto"
                    .Fields("NReparto").Value = txtAlta(0).Text
                    .Fields("Descript").Value = Left(txtAlta(1).Text, 255)
                    
                Case "TMB"
                
                'Dim vvalues, vcampos As String
                'vvalues = txtAlta(0).Text + "," + txtAlta(1).Text + "," + txtAlta(2).Text + txtAlta(3).Text + "," + txtAlta(4).Text + "," + txtAlta(5).Text
                'vcampos = "Codigo,TipoMovimiento,ProveedorClienteBanco,IngresoEgreso,ListadoIva"
                
                
                'Call InsertarEnTabla("tipomovimientos", vcampos, vvalues)
                
                    .Fields("codigo").Value = Left(txtAlta(0).Text, 2)
                    .Fields("TipoMovimiento").Value = Left(txtAlta(1).Text, 255)
                    .Fields("ProveedorClienteBanco").Value = Left(txtAlta(2).Text, 1)
                    .Fields("IngresoEgreso").Value = Left(txtAlta(3).Text, 1)
                    .Fields("ListadoIva").Value = Left(txtAlta(4).Text, 1)
                
                Case "TipoIva", "TipoCliente", "Actividad", ""
                    .Fields(0).Value = txtAlta(0).Text
                    .Fields(1).Value = Left(txtAlta(1).Text, 255)
            
                Case "Lista"
                    .Fields(0).Value = txtAlta(0).Text
                    .Fields(1).Value = Left(txtAlta(1).Text, 255)
                
                Case "Rubro"
                    .Fields(0).Value = txtAlta(0).Text
                    .Fields(1).Value = Left(txtAlta(1).Text, 255)
                
                Case "SubRubro"
                    .Fields(0).Value = txtAlta(0).Text
                    .Fields(2).Value = Left(txtAlta(1).Text, 255)
                
                Case "EstadoCliente"
                    .Fields(0).Value = Left(txtAlta(0).Text, 3)
                    .Fields(1).Value = Left(txtAlta(1).Text, 255)
                    .Fields(2).Value = Left(txtAlta(2).Text, 1)
                    .Fields(3).Value = Left(txtAlta(3).Text, 1)
            End Select
            
            .Update
        
        End If
        
    End With

    sqlAlta = ""
    
    If rsAlta.State = 1 Then
        rsAlta.Close
        Set rsAlta = Nothing
    End If
    
    Call frmBusqueda.Form_Unload(0) ' Alfredo: cierro la busqueda porque no va a mostar lo ultimo que puse
    

    If Err Then
        GrabarLog "Grabar", Err.Number & " " & Err.Description, Me.Name
    Else
        Unload Me
    End If

End Sub
Private Function ValidarCampos() As Boolean
    On Error Resume Next

    Dim i As Integer
    
    ValidarCampos = True
    
    For i = 0 To txtAlta.Count - 1
        If Trim(txtAlta(1).Text) = "" Then
            MsgBox "Campos importantes vacios!", vbExclamation, "Mensaje ..."
           ' ValidarCampos = Not True
            ValidarCampos = True
            Exit Function
        End If
    Next
    
    If Me.vaccion = "Nuevo" Then
       
        
        'Hacer un Select
        
        
        
        
        'If Not Trim(TraerDato(vGrabarTabla, "Codigo = '" & Trim(txtAlta(0).Text) & "'", "Codigo")) = "" Then
        '    MsgBox "Existe un registro con ese codigo!", vbExclamation, "Mensaje ..."
        '    ValidarCampos = Not True
        '    Exit Function
        'End If
    End If
    
    If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub txtAlta_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next
    
    If KeyAscii = 13 Then
        If txtAlta(Index + 1).Visible = True Then
            txtAlta(Index + 1).SetFocus
        Else
            PbAcciones(0).SetFocus
        End If
    End If
    
If Err Then GrabarLog "txtAlta_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
