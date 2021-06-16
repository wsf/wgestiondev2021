VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{63BEADB1-20E1-478A-9B40-DDDAFBF3624F}#1.0#0"; "bsGradientLabel.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmSistemaDeRecargos 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sistema de Recargos de Clientes"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   9015
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid dgRecargos 
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4048
      _Version        =   393216
      HeadLines       =   1.25
      RowHeight       =   15
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
   Begin XtremeSuiteControls.GroupBox GBRecarga 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   8880
      _Version        =   851968
      _ExtentX        =   15663
      _ExtentY        =   3201
      _StockProps     =   79
      Caption         =   "Recarga"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   465
         Index           =   0
         Left            =   7680
         TabIndex        =   7
         Top             =   1275
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   820
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Picture         =   "frmMigrarCuentas.frx":0000
      End
      Begin XtremeSuiteControls.GroupBox GBClientes 
         Height          =   495
         Left            =   4680
         TabIndex        =   12
         Top             =   720
         Width           =   4095
         _Version        =   851968
         _ExtentX        =   7223
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RBClientes 
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Tag             =   "Todos"
            Top             =   180
            Width           =   855
            _Version        =   851968
            _ExtentX        =   1508
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Todos"
            Appearance      =   6
            Value           =   -1  'True
         End
      End
      Begin XtremeSuiteControls.GroupBox GBAplicacable 
         Height          =   495
         Left            =   4680
         TabIndex        =   9
         Top             =   210
         Width           =   4095
         _Version        =   851968
         _ExtentX        =   7223
         _ExtentY        =   873
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin XtremeSuiteControls.RadioButton RBAplicable 
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Tag             =   "UF"
            Top             =   180
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Ultima Factura"
            Appearance      =   6
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton RBAplicable 
            Height          =   255
            Index           =   1
            Left            =   1560
            TabIndex        =   11
            Tag             =   "ND"
            Top             =   180
            Width           =   1335
            _Version        =   851968
            _ExtentX        =   2355
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Nuevo Debito"
            Appearance      =   6
         End
      End
      Begin VB.TextBox txtPorcenjateRecargo 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   840
         Width           =   1800
      End
      Begin VB.TextBox txtCantidadDeDias 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1800
      End
      Begin VB.Label lblRecargo 
         Alignment       =   1  'Right Justify
         Caption         =   "Clientes :"
         Height          =   195
         Index           =   3
         Left            =   3480
         TabIndex        =   8
         Top             =   870
         Width           =   1140
      End
      Begin VB.Label lblRecargo 
         Alignment       =   1  'Right Justify
         Caption         =   "Aplicable A:"
         Height          =   195
         Index           =   2
         Left            =   3480
         TabIndex        =   6
         Top             =   390
         Width           =   1140
      End
      Begin VB.Label lblRecargo 
         Alignment       =   1  'Right Justify
         Caption         =   "% de Recargo :"
         Height          =   195
         Index           =   1
         Left            =   30
         TabIndex        =   5
         Top             =   870
         Width           =   1500
      End
      Begin VB.Label lblRecargo 
         Alignment       =   1  'Right Justify
         Caption         =   "Cantidad de Dias :"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   390
         Width           =   1500
      End
   End
   Begin Project1.bsGradientLabel bsTitulo 
      Height          =   345
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   9000
      _ExtentX        =   15875
      _ExtentY        =   609
      Caption         =   "Sistema de Recarga Por Cta. Cte a Clientes"
      BeginProperty Fount {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColour   =   0
      Colour1         =   14737632
      Colour2         =   12632256
      CaptionAlignment=   1
      BorderStyle     =   6
   End
End
Attribute VB_Name = "frmSistemaDeRecargos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsRecargos As ADODB.Recordset
Private Sub Form_Load()
On Error Resume Next

    CargarGrilla

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla()
On Error Resume Next

    'Me.Show
    With dgRecargos
        .HeadLines = 1.5
        
        .Columns(0).Width = 0
        
        .Columns(1).Width = 2000
        .Columns(1).Caption = "Cantidad De Dias"
        .Columns(1).Alignment = dbgRight
        
        .Columns(2).Width = 2000
        .Columns(2).Alignment = dbgRight
        .Columns(2).DataFormat.Format = " %######0.00"
        
        .Columns(3).Width = 2000
        .Columns(3).Alignment = dbgRight
        
        .Columns(4).Width = 2000
        .Columns(4).Alignment = dbgRight
    
    End With

If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarGrilla()
On Error Resume Next

    Set rsRecargos = New ADODB.Recordset
    Dim sqlRecargos  As String
    
    sqlRecargos = "SELECT * FROM Recargos"
    
    With rsRecargos
        Call .Open(sqlRecargos, ConnDDBB, adOpenStatic, adLockReadOnly)
    
        If .State = 1 Then
            Set dgRecargos.DataSource = rsRecargos
            FormatoGrilla
        Else
            Set dgRecargos.DataSource = Nothing
        End If
    
    End With

    sqlRecargos = ""

If Err Then GrabarLog "CargarGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
    
        Select Case Index
        
            Case 0
                AgregarRecargo
        
        End Select

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub AgregarRecargo()
On Error Resume Next

    If Val(txtCantidadDeDias.Text) = 0 Then
        MsgBox "Ingrese la cantidad de Dias...", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    If Val(txtPorcenjateRecargo.Text) = 0 Then
        MsgBox "Ingrese el Porcentaje de Recarga...", vbExclamation, "Mensaje ..."
        Exit Sub
    End If

    Dim rsRecargo As New ADODB.Recordset, sqlRecargo As String
    Dim i As Integer

    sqlRecargo = "SELECT * FROM Recargos WHERE 1=2"
    
    With rsRecargo
        .CursorLocation = adUseClient
        
        Call .Open(sqlRecargo, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then
            .AddNew
            
            .Fields("CantidadDias").Value = Val(txtCantidadDeDias.Text)
            .Fields("Porcentaje").Value = Val(txtPorcenjateRecargo.Text)
            
            For i = 0 To RBAplicable.Count - 1
                If RBAplicable(0).Value = True Then
                    .Fields("AplicableA").Value = RBAplicable(0).Tag
                    Exit For
                Else
                    .Fields("AplicableA").Value = RBAplicable(1).Tag
                    Exit For
                End If
            
            Next
            
            .Fields("TipoClientes").Value = RBClientes.Tag
            
            .Update
        End If
        
    End With
    
    Limpiar
    
    sqlRecargo = ""
    
    If rsRecargo.State = 1 Then
        rsRecargo.Close
        Set rsRecargo = Nothing
    End If
    
If Err Then GrabarLog "AgregarRecargo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Limpiar()
On Error Resume Next

    RBAplicable(0).Value = True
    RBClientes.Value = True
    txtCantidadDeDias.Text = ""
    txtPorcenjateRecargo.Text = ""

If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
