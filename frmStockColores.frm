VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmStockColores 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stock Rango - Colores"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4020
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   4020
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.ColorPicker CPAlerta 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   4560
      Width           =   1755
      _Version        =   851968
      _ExtentX        =   3087
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Elija el color"
      Appearance      =   5
      SelectedColor   =   65280
      ShowAutomaticColor=   0   'False
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   4200
      Width           =   1750
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   1750
   End
   Begin XtremeSuiteControls.GroupBox GBAcciones 
      Height          =   975
      Left            =   0
      TabIndex        =   7
      Top             =   4920
      Width           =   4020
      _Version        =   851968
      _ExtentX        =   7091
      _ExtentY        =   1720
      _StockProps     =   79
      Caption         =   "Acciones"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   310
         Width           =   1200
         _Version        =   851968
         _ExtentX        =   2117
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "&Limpiar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmStockColores.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   495
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   315
         Width           =   1200
         _Version        =   851968
         _ExtentX        =   2117
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Guardar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmStockColores.frx":041E
      End
      Begin XtremeSuiteControls.PushButton PBAcciones 
         Height          =   495
         Index           =   2
         Left            =   2760
         TabIndex        =   10
         Top             =   315
         Width           =   1200
         _Version        =   851968
         _ExtentX        =   2117
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Salir"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmStockColores.frx":0833
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Detalle 
      Height          =   3675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   6482
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   12632256
      ForeColorFixed  =   4210752
      BackColorSel    =   255
      ScrollBars      =   2
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblAlta 
      Alignment       =   1  'Right Justify
      Caption         =   "> Color de Alerta :"
      Height          =   195
      Index           =   2
      Left            =   0
      TabIndex        =   5
      Top             =   4560
      Width           =   1995
   End
   Begin VB.Label lblAlta 
      Alignment       =   1  'Right Justify
      Caption         =   "> Cantidad Maxima :"
      Height          =   195
      Index           =   1
      Left            =   0
      TabIndex        =   4
      Top             =   4200
      Width           =   1995
   End
   Begin VB.Label lblAlta 
      Alignment       =   1  'Right Justify
      Caption         =   "> Cantidad Minima :"
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   1995
   End
End
Attribute VB_Name = "frmStockColores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsStockColores As ADODB.Recordset
Private Sub Form_Load()
On Error Resume Next
    
    Me.Show
    CargarStockColores
    ConfigurarGrilla

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CargarStockColores()
On Error Resume Next

    Set rsStockColores = New ADODB.Recordset
    Dim sqlStockColores As String
    
    sqlStockColores = "SELECT * FROM StockColores"

    With rsStockColores
        If .State = 1 Then .Close
        
        .CursorLocation = adUseClient
        
        Call .Open(sqlStockColores, ConnDDBB, adOpenStatic, adLockReadOnly)
        
    End With
    
If Err Then GrabarLog "CargarStockColores", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ConfigurarGrilla()
On Error Resume Next

    Dim i As Integer
   
    With Detalle
        .Cols = 4
        
        .FixedRows = 1
        .FixedCols = 1
            
            
        .ColWidth(0) = 250
        .ColWidth(1) = 0
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1250
        
        .Row = .Rows - 1
    
        Set .DataSource = rsStockColores
        
        .TextMatrix(0, 2) = "C. Min"
        .TextMatrix(0, 3) = "C. Max"
        .TextMatrix(0, 4) = "Color"


        AsignarColor
    End With

    
If Err Then GrabarLog "ConfigurarGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub AsignarColor()
On Error Resume Next

    Dim i, j, k As Integer, vCantidadStock As Long, vColor As Long
    
    With Detalle
        
        'Recorro las Filas
        i = 0
        For i = 1 To .Rows - 1
            .Row = i
            
            vColor = .TextMatrix(.Row, 4)
            
            If vColor = 0 Then vColor = &HFFFFFF
            
            
            'Recorro las Columnas por fila
            j = 0
            For j = 2 To .Cols - 1
                .Col = j
                .CellBackColor = vColor
            Next

        Next
    
    End With
    
If Err Then GrabarLog "AsignarColor", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Limpiar
            
        Case 1
            Guardar
            Limpiar
            CargarStockColores
            ConfigurarGrilla
        
        Case 2
            Unload Me
    
    End Select

If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Limpiar()
On Error Resume Next

    txtCantidad(0).Text = ""
    txtCantidad(1).Text = ""
    CPAlerta.SelectedColor = vbBlack
    
If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Guardar()
On Error Resume Next
    
    Dim rsGuardar As New ADODB.Recordset, sqlGuardar As String
    
    sqlGuardar = "SELECT * FROM StockColores"
    
    With rsGuardar
        Call .Open(sqlGuardar, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        .AddNew
        
        .Fields("CantidadMin").Value = Val(txtCantidad(0).Text)
        .Fields("CantidadMax").Value = Val(txtCantidad(1).Text)
        .Fields("Color").Value = "&H" & Hex(CPAlerta.SelectedColor)
    
        .Update
        
    End With
    
    sqlGuardar = ""

    If rsGuardar.State = 1 Then
        rsGuardar.Close
        Set rsGuardar = Nothing
    End If
    
If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtCantidad_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If Index = 0 Then txtCantidad(Index + 1).SetFocus
        If Index = 1 Then PbAcciones(1).SetFocus
    End If
    
If Err Then GrabarLog "txtCantidad_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
