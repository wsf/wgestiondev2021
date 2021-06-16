VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBuscarContacto 
   Caption         =   "Busca Contáctos"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   6390
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   6390
   Begin MSDataGridLib.DataGrid dgBuscarContactos 
      Bindings        =   "buscacontacto.frx":0000
      Height          =   2835
      Left            =   45
      TabIndex        =   2
      Top             =   90
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   5001
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
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
         DataField       =   "Correo"
         Caption         =   "Correo"
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
         DataField       =   "Ocupación"
         Caption         =   "Ocupación"
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
            ColumnWidth     =   959.811
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2910.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1454.74
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
   Begin VB.TextBox txtContacto 
      Height          =   285
      Left            =   2490
      TabIndex        =   0
      Top             =   3000
      Width           =   3795
   End
   Begin VB.CommandButton cmdSeleccionar 
      Caption         =   "<<"
      Height          =   345
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Volver al módulo anterior"
      Top             =   1950
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   345
      Left            =   2610
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2070
      Visible         =   0   'False
      Width           =   795
   End
   Begin MSAdodcLib.Adodc barticulo 
      Height          =   405
      Left            =   240
      Top             =   3480
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
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
      Caption         =   "barticulo"
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
   Begin VB.Label lblBuscar 
      BackColor       =   &H00000080&
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba código o descripción :"
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Top             =   3030
      Width           =   2385
   End
End
Attribute VB_Name = "frmBuscarContacto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Buscar As String

Public o As Integer

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSeleccionar_Click()

    If barticulo.Recordset.EOF Then cmdSalir_Click
    ' buscamo el código de artículo
    dgBuscarContactos.Col = 0
    
    Select Case o

        Case 1
           frmRemito.txtCliente(0).Text = barticulo.Recordset(0).Value
           frmRemito.txtCliente_KeyPress 0, 13
            Unload Me

        Case 3
            frmBuscarFactura.txtCliente.Text = barticulo.Recordset(0).Value
            frmBuscarFactura.txtCliente_KeyPress 13
            Unload Me

        Case 4
            'frmCtaCteC.txtCliente = barticulo.Recordset(0).Value
            'frmCtaCteC.txtCliente_Keypress 13
            Unload Me

        Case 5
            frmCheques.txtNombre = barticulo.Recordset(0).Value
            frmCheques.txtNombre_KeyPress 13
            Unload Me

        Case 6
            frmCheques.cvnombre = barticulo.Recordset(0).Value
            frmCheques.cvnombre_KeyPress 13
            Unload Me

        Case 7
            'frmCreditos.vNombre = barticulo.Recordset(0).Value
            'frmCreditos.vnombre_KeyPress 13
            Unload Me

        Case 11
            'EstCliente.varticulo = barticulo.Recordset(0)
            'EstCliente.varticulo_KeyPress 13
            'EstCliente.WindowState = 2
            'Unload Me

        Case 8
            frmSaldosClientes.vcdesde = barticulo.Recordset(0).Value
            frmSaldosClientes.vcdesde_KeyPress 13
            'frmSaldosClientes.WindowState = 2
            Unload Me

        Case 9
            frmSaldosClientes.vchasta = barticulo.Recordset(0).Value
            frmSaldosClientes.vchasta_KeyPress 13
            'frmSaldosClientes.WindowState = 2
            Unload Me

        Case 0
            frmGuia.bguia.Recordset.MoveFirst

            If Buscar = "" Then Buscar = barticulo.Recordset(0).Value
            frmGuia.bguia.Recordset.Find ("codigo = '" + Buscar + "'")
            frmGuia.mostrar
            Unload Me
    End Select

End Sub
Private Sub Form_Load()
On Error Resume Next

    With barticulo
        .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM Guia"
        .Refresh
    End With
    
    With Me
        .Top = 1400
        .Left = 2500
        .width = 6510
        .height = 3780
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgBuscarContactos_Click()
    On Error Resume Next
    
    Buscar = dgBuscarContactos.Text

    If Err Then GrabarLog "dgBuscarContactos_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgBuscarContactos_DblClick()
On Error Resume Next

    cmdSeleccionar_Click

If Err Then GrabarLog "dgBuscarContactos_DblClick", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtContacto_Change()
On Error Resume Next

    With barticulo
        If Trim(txtContacto.Text) = "" Then
            .RecordSource = "SELECT * FROM guia"
        Else
            .RecordSource = "SELECT * FROM guia WHERE (nombre LIKE '%" & Trim(txtContacto.Text) & "%') OR (codigo LIKE '%" & Trim(txtContacto.Text) + "%')"
        End If

        .Refresh
    End With
    
If Err Then GrabarLog "txtContacto_Change", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub txtContacto_keypress(KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 27 Then
        Unload Me
    End If

    If KeyAscii = 13 Then
        cmdSeleccionar_Click
    End If

If Err Then GrabarLog "txtContacto_keypress", Err.Number & " " & Err.Description, Me.Name
End Sub

