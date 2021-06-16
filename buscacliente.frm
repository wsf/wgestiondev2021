VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmBuscarCliente 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Busca Clientes"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   10770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   345
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4080
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "<<"
      Height          =   345
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Volver al módulo anterior"
      Top             =   4080
      Visible         =   0   'False
      Width           =   795
   End
   Begin MSAdodcLib.Adodc bclientes 
      Height          =   405
      Left            =   0
      Top             =   3480
      Visible         =   0   'False
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   714
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
   Begin VB.TextBox txtClientes 
      Height          =   285
      Left            =   2460
      TabIndex        =   0
      Top             =   3000
      Width           =   8265
   End
   Begin MSDataGridLib.DataGrid dgClientes 
      Bindings        =   "buscacliente.frx":0000
      Height          =   2955
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10635
      _ExtentX        =   18759
      _ExtentY        =   5212
      _Version        =   393216
      AllowUpdate     =   -1  'True
      BackColor       =   16777215
      BorderStyle     =   0
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   4
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
         EndProperty
         BeginProperty Column01 
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
   Begin VB.Label lblNombre 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba código o descripción :"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   3060
      Width           =   2130
   End
End
Attribute VB_Name = "frmBuscarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Buscar As String
Public o As Integer
Dim ConnClientes As ADODB.Connection
Dim rsClientes As ADODB.Recordset
Private Sub cmdSalir_Click()
On Error Resume Next
    
    Unload Me
    
If Err Then GrabarLog "cmdSalir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Seleccionar()
On Error Resume Next

    
    With rsClientes
        If .EOF = True Then cmdSalir_Click
   
        Select Case o

            Case 0
            
            Case 1
                frmRemito.txtCliente(0).Text = .Fields("Codigo").Value
                frmRemito.txtCliente_KeyPress 0, 13

                Unload Me
            
            Case 2
                'frmControlFacturacion.txtCliente.Text = .Fields("Codigo").value
                'frmControlFacturacion.txtCliente_Keypress 13
                'ZOrder 0
                'Unload Me
        
            Case 3
                frmBuscarFactura.txtCliente.Text = .Fields(1).Value
                frmBuscarFactura.txtCliente_KeyPress 13
                Unload Me
    
            Case 4
               ' frmCtaCteC.txtCliente = .Fields(1).Value
                'frmCtaCteC.txtCliente_Keypress 13
                Unload Me
    
            Case 5
                frmCheques.txtNombre = .Fields(1).Value
                frmCheques.txtNombre_KeyPress 13
                Unload Me
    
            Case 6
                frmCheques.cvnombre = .Fields(1).Value
                frmCheques.cvnombre_KeyPress 13
                Unload Me
    
            Case 7
                'frmCreditos.vnombre = .Fields(1).Value
                'frmCreditos.vnombre_KeyPress 13
                Unload Me
    
            Case 8
                'frmSaldosClientes.vcdesde = .Fields(1).value
                'frmSaldosClientes.vcdesde_KeyPress 13
                Unload Me
    
            
            Case 9
                'frmSaldosClientes.vchasta = .Fields(1).value
                'frmSaldosClientes.vchasta_KeyPress 13
                Unload Me
            
            Case 10
                'frmComentario.txtCliente.Text = .Fields(1).value
                'frmComentario.txtCliente_Keypress 13
                Unload Me
    
            Case 11
                Unload Me
        End Select

    End With

If Err Then GrabarLog "Seleccionar", Err.Number & " " & Err.Number, Me.Caption
End Sub
Private Sub cmdSeleccionar_Click()
On Error Resume Next

    Seleccionar

If Err Then GrabarLog "cmdSeleccionar_Click", Err.Number & " " & Err.Number, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    With Me
        .Top = 1400
        .Left = 2500
        .Width = 10890
        .Height = 3990
    End With
 Call CentrarFormulario(Me)
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Number, Me.Caption
End Sub

Private Sub dgClientes_Click()
    On Error Resume Next
    
    Buscar = Trim(dgClientes.Text)

    If Err Then GrabarLog "dgClientes_Click", Err.Number & " " & Err.Number, Me.Caption
End Sub
Private Sub dgClientes_DblClick()
    On Error Resume Next
    
    Seleccionar
    
    If Err Then GrabarLog "dgClientes_DblClick", Err.Number & " " & Err.Number, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    If rsClientes.State = 1 Then
        rsClientes.Close
        Set rsClientes = Nothing
    End If
    
    If Err Then GrabarLog "dgClientes_DblClick", Err.Number & " " & Err.Number, Me.Caption
End Sub
Private Sub txtClientes_Change()
On Error Resume Next

    Set rsClientes = New ADODB.Recordset
    Dim sqlClientes As String
    
    If Trim(txtClientes.Text) = "" Then
        sqlClientes = "SELECT * FROM clientes"
    Else
        sqlClientes = "SELECT * FROM clientes WHERE (codigo LIKE '%" + Trim(txtClientes.Text) + "%') OR (nombre LIKE '%" + Trim(txtClientes.Text) + "%')"
    End If
        
    With rsClientes
        .CursorLocation = adUseClient
        Call .Open(sqlClientes, ConnDDBB, adOpenStatic, adLockReadOnly)
        Set dgClientes.DataSource = rsClientes
    End With
        
    sqlClientes = ""
    
    If Err Then GrabarLog "txtClientes_Change", Err.Number & " " & Err.Number, Me.Caption
End Sub
Public Sub txtClientes_keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 27 Then
        Unload Me
    End If

    If KeyAscii = 13 Then
        Seleccionar
    End If
    
If Err Then GrabarLog "txtClientes_keypress", Err.Number & " " & Err.Number, Me.Caption
End Sub

