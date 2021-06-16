VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBuscarProveedor 
   BackColor       =   &H80000016&
   Caption         =   "Busca Proveedor"
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   8775
   Begin VB.TextBox txtProveedor 
      Height          =   315
      Left            =   2400
      TabIndex        =   0
      Top             =   3060
      Width           =   6255
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   345
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "<<"
      Height          =   345
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Volver al módulo anterior"
      Top             =   2400
      Visible         =   0   'False
      Width           =   825
   End
   Begin MSDataGridLib.DataGrid dgProveedores 
      Bindings        =   "buscaproveedor.frx":0000
      Height          =   2955
      Left            =   60
      TabIndex        =   3
      Top             =   60
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   5212
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   -2147483628
      BorderStyle     =   0
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nombre"
         Caption         =   "Nombre"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Direccion"
         Caption         =   "Direccion"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Telefono"
         Caption         =   "Telefono"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Ibrutos"
         Caption         =   "Ibrutos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Iva"
         Caption         =   "Iva"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Cuit"
         Caption         =   "Cuit"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Credito"
         Caption         =   "Credito"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Responsable"
         Caption         =   "Responsable"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2970.142
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
         BeginProperty Column08 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblProveedor 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba código o descripción :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   3120
      Width           =   2355
   End
End
Attribute VB_Name = "frmBuscarProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Buscar As String
Public o As Integer
Dim rsProveedores As ADODB.Recordset
Private Sub cmdSalir_Click()
On Error Resume Next

    Unload Me

If Err Then GrabarLog "cmdSalir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdSeleccionar_Click()
    On Error Resume Next
        
    With rsProveedores
        If .EOF = True Then cmdSalir_Click

        Select Case o
            Case 0
                'dgProveedores.Col = 0
                
                'frmProveedores.bProveedor.Recordset.MoveFirst
    
                'If Buscar = "" Then Buscar = .Fields("Codigo").Value
                'Unload Me
                
            Case 1
                frmCompras.txtProveedor(0) = .Fields("Codigo").Value
                frmCompras.txtProveedor_KeyPress 0, 13
                Unload Me

            Case 5
                frmCheques.txtNombre = .Fields("Codigo").Value
                frmCheques.txtNombre_KeyPress 13
                Unload Me

            Case 6
                frmCheques.cvnombre = .Fields("Codigo").Value
                frmCheques.cvnombre_KeyPress 13
                frmCheques.cvncheque.SetFocus
                Unload Me

            Case 3
                'frmBuscarCompra.txtProveedor.Text = .Fields("Codigo").value
                'frmBuscarCompra.txtProveedor_KeyPress 13
                Unload Me

            Case 4
                frmCtaCteP.txtProveedor.Text = .Fields("Codigo").Value
                frmCtaCteP.txtProveedor_KeyPress 13
                Unload Me

            Case 8
               ' frmSaldosProveedores.txtProveedor(0).Tag = .Fields("Codigo").Value
               ' frmSaldosProveedores.txtProveedor(0).Text = .Fields("Nombre").Value
                Unload Me

            Case 9
                'frmSaldosProveedores.txtProveedor(1).Tag = .Fields("Codigo").Value
                'frmSaldosProveedores.txtProveedor(1).Text = .Fields("Nombre").Value
                Unload Me

            Case 10
                Unload Me
        End Select

    End With
    
    
    
If Err Then GrabarLog "cmdSeleccionar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next

    With Me
        .Show
        .Top = 1400
        .Left = 2500
        .Width = 8925
        .Height = 4000
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgProveedores_DblClick()
On Error Resume Next

    cmdSeleccionar_Click
    
If Err Then GrabarLog "dgProveedores_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub txtProveedor_Change()
On Error Resume Next

    Set rsProveedores = New ADODB.Recordset
    Dim sqlProveedores As String
    
    If Trim(txtProveedor.Text) = "" Then
        sqlProveedores = "SELECT * FROM proveedores"
    Else
        sqlProveedores = "SELECT * FROM proveedores WHERE (nombre LIKE '%" + txtProveedor + "%') OR (codigo LIKE '%" + txtProveedor + "%')"
    End If
        
    With rsProveedores
        If .State = 1 Then .Close
        .CursorLocation = adUseClient

        Call .Open(sqlProveedores, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        Set dgProveedores.DataSource = rsProveedores
    
    End With

If Err Then GrabarLog "txtProveedor_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtProveedor_KeyPress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 27 Then cmdSalir_Click
    If KeyAscii = 13 Then cmdSeleccionar_Click
    
If Err Then GrabarLog "varticulo_keypress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If rsProveedores.State = 1 Then
        rsProveedores.Close
        Set rsProveedores = Nothing
    End If
    
    If Err Then GrabarLog "dgClientes_DblClick", Err.Number & " " & Err.Number, Me.Caption
End Sub
