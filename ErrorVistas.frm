VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmErrorVistas 
   Caption         =   "Errores en las Vistas"
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8355
   Icon            =   "ErrorVistas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   8355
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   5
      Top             =   8580
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   714
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Imprimir"
      Enabled         =   0   'False
      Height          =   585
      Left            =   7350
      Picture         =   "ErrorVistas.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7980
      UseMaskColor    =   -1  'True
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   30
      TabIndex        =   1
      Top             =   60
      Width           =   8295
      Begin VB.Label lblCantidad 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   18
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   6780
         TabIndex        =   3
         Top             =   180
         Width           =   1290
      End
      Begin VB.Label lblErrores 
         Caption         =   "> Cantidad de Erorres:"
         Height          =   285
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   330
         Width           =   1725
      End
   End
   Begin MSDataGridLib.DataGrid dgErrorVista 
      Bindings        =   "ErrorVistas.frx":6954
      Height          =   6705
      Left            =   30
      TabIndex        =   0
      Top             =   990
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   11827
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
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
         DataField       =   ""
         Caption         =   ""
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label lblErrores 
      Caption         =   "Haga doble clic sobre el cliente con error para ir automaticamente a la CtaCte para corregirlo."
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   6
      Top             =   8280
      Width           =   6735
   End
End
Attribute VB_Name = "frmErrorVistas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsErrorVista As ADODB.Recordset
Dim connErrorVista As ADODB.Connection
Private Sub cmdGenerar_Click()
On Error Resume Next

    Call Imprimir
    
If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub dgErrorVista_DblClick()
On Error Resume Next

    With frmCtaCteC
        .txtCliente.Text = rsErrorVista.Fields("codigo").Value
        '.txtCliente_Keypress (13)
    End With

If Err Then GrabarLog "dgErrorVista_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next

    If KeyCode = vbKeyF5 Then
        Call Imprimir
    End If
    
If Err Then GrabarLog "Form_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Imprimir()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox " Prepare la Impresora!!!!", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsVistas
        If .State = 1 Then .Close
        
        .Source = "SELECT * FROM ErrorVista2"
        
        If .State = 0 Then .Open
        .Close
        .Open
        
    End With
    
    With drArreglos
        .Sections("SecEncabezado").Controls("lblTitulo").Caption = "Listado de Clientes con diferencias en las Vistas de la Ctas. Ctes."
        .Sections("SecEncabezado").Controls("lblCodigo").Caption = "Codigo"
        .Sections("SecEncabezado").Controls("lblNombre").Caption = "Nombre"
        .Sections("SecEncabezado").Controls("lblFichaCliente").Caption = "Ficha Cliente"
        .Sections("SecEncabezado").Controls("lblFichaFactura").Caption = "Ficha Factura"
        .Sections("SecEncabezado").Controls("lblDiferencia").Caption = "Diferencia"
    
        .Sections("Detalle").Controls("lblDiferencia").Caption = "Diferencia"
    
        .Sections("Detalle").Controls("txtCodigo").DataField = Mantenimiento.rsVistas.Fields("Codigo").Name
        .Sections("Detalle").Controls("txtNombre").DataField = Mantenimiento.rsVistas.Fields("Nombre").Name
        .Sections("Detalle").Controls("txtFichaCliente").DataField = Mantenimiento.rsVistas.Fields("FichaCliente").Name
        .Sections("Detalle").Controls("txtFichaFactura").DataField = Mantenimiento.rsVistas.Fields("FichaFactura").Name
        .Sections("Detalle").Controls("txtDiferencia").DataField = Mantenimiento.rsVistas.Fields("Diferencia").Name
        
        .Show
    End With
    

If Err Then GrabarLog "Imprimir", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next

    CargarErrores

    KeyPreview = True
    cmdGenerar.Enabled = True

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarErrores()
On Error Resume Next

    Set connErrorVista = New ADODB.Connection
    Set rsErrorVista = New ADODB.Recordset
    Dim sqlErrorVista As String
    
    
    With connErrorVista
        .ConnectionString = pathDBMySQL
        .Open
    End With

    sqlErrorVista = "SELECT * FROM errorvista2"
    
    With rsErrorVista
        .CursorLocation = adUseClient
        Call .Open(sqlErrorVista, connErrorVista, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            Set dgErrorVista.DataSource = rsErrorVista
        End If
        
        lblCantidad.Caption = .RecordCount
    End With

If Err Then GrabarLog "CargarErrores", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    rsErrorVista.Close
    Set rsErrorVista = Nothing
    
    connErrorVista.Close
    Set connErrorVista = Nothing

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
