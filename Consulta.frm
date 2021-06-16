VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmConsulta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consulta de Precio de Artículo"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   11940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   11940
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar BarraEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   6555
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   9690
      TabIndex        =   8
      Top             =   5130
      Width           =   2145
      Begin XtremeSuiteControls.PushButton PBRango 
         Height          =   315
         Left            =   30
         TabIndex        =   12
         Top             =   740
         Width           =   2050
         _Version        =   851968
         _ExtentX        =   3616
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Ver Rangos"
         UseVisualStyle  =   -1  'True
         Picture         =   "Consulta.frx":0000
      End
      Begin XtremeSuiteControls.PushButton cmdEjecutar 
         Height          =   615
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   2050
         _Version        =   851968
         _ExtentX        =   3616
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "&Ejecutar"
         UseVisualStyle  =   -1  'True
         Picture         =   "Consulta.frx":0462
      End
   End
   Begin VB.Frame fraBusqueda 
      Caption         =   "Datos de la consultas :"
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   60
      TabIndex        =   2
      Top             =   5130
      Width           =   9495
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   1890
         TabIndex        =   0
         Top             =   300
         Width           =   5115
      End
      Begin VB.TextBox txtDescripcion 
         Height          =   315
         Left            =   1890
         TabIndex        =   1
         Top             =   660
         Width           =   5115
      End
      Begin VB.ComboBox ordenado 
         Height          =   315
         ItemData        =   "Consulta.frx":0A4E
         Left            =   7590
         List            =   "Consulta.frx":0A61
         TabIndex        =   4
         Text            =   "Codigo"
         Top             =   570
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Frame3"
         Height          =   975
         Left            =   10230
         TabIndex        =   3
         Top             =   150
         Width           =   15
      End
      Begin VB.Label lblBusqueda 
         Alignment       =   1  'Right Justify
         Caption         =   "Ingresar Codigo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   30
         TabIndex        =   7
         Top             =   330
         Width           =   1800
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Ingresar Descricion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   30
         TabIndex        =   6
         Top             =   690
         Width           =   1800
      End
      Begin VB.Label Label3 
         Caption         =   "Listado ordenado por :"
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   7590
         TabIndex        =   5
         Top             =   300
         Width           =   1635
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Detalle 
      Height          =   4995
      Left            =   60
      TabIndex        =   9
      Top             =   120
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   8811
      _Version        =   393216
      BackColor       =   16777215
      Cols            =   26
      FixedRows       =   0
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
      _Band(0).Cols   =   26
   End
   Begin XtremeSuiteControls.ProgressBar Barra 
      Height          =   255
      Left            =   60
      TabIndex        =   11
      Top             =   6250
      Width           =   11775
      _Version        =   851968
      _ExtentX        =   20770
      _ExtentY        =   450
      _StockProps     =   93
      Scrolling       =   1
      Appearance      =   4
      UseVisualStyle  =   0   'False
   End
End
Attribute VB_Name = "frmConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBusqueda As ADODB.Recordset
Dim vRango() As String
Dim vCantidadRangos As Integer
Private Sub cmdEjecutar_Click()
    On Error Resume Next
    
    Set rsBusqueda = New ADODB.Recordset
    Dim sqlBusqueda As String, sql As String
    
    MousePointer = vbHourglass
    
    cmdEjecutar.Enabled = Not True
    
    CargarRangos

    sql = ""
    
    If Not Trim(txtCodigo.Text) = "" Then sql = sql + " AND (Codigo = '" & Trim(txtCodigo.Text) & "')"
    If Not Trim(txtDescripcion.Text) = "" Then sql = sql + " AND (Descrip LIKE '%" & Trim(txtDescripcion.Text) & "%')"
    
    sqlBusqueda = "SELECT codigo, Descrip, Stock, PCosto, PVenta1, (PVenta1-PCosto) as Ganancia FROM Articulos WHERE 1=1" & sql & " ORDER BY " & Trim(ordenado.Text)

    With rsBusqueda
        If .State = 1 Then .Close
        
        .CursorLocation = adUseClient
        
        Call .Open(sqlBusqueda, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If .EOF = True Then Set Detalle.DataSource = Nothing
    
        ConfigurarGrilla
    End With
    
    txtCodigo.SetFocus
    
    cmdEjecutar.Enabled = True
    MousePointer = vbDefault
    
    If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ConfigurarGrilla()
On Error Resume Next

    Dim i As Integer
   
    With Detalle
        .Cols = 7
        
        .FixedRows = 1
        .FixedCols = 1
            
        .ColWidth(0) = 250
        .ColWidth(1) = 1000
        .ColWidth(2) = 5000
        .ColWidth(3) = 1250
        .ColWidth(4) = 1250
        .ColWidth(5) = 1250
        .ColWidth(6) = 1250
    
        .Row = Detalle.Rows - 1
    
        Set .DataSource = rsBusqueda
        
        .TextMatrix(0, 1) = "Código"
        .TextMatrix(0, 2) = "Detalle"
        .TextMatrix(0, 3) = "Stock"
        .TextMatrix(0, 4) = "P. Costo"
        .TextMatrix(0, 5) = "P. Venta"
        .TextMatrix(0, 6) = "Ganancia"


        
        AsignarColor
    End With

    
If Err Then GrabarLog "ConfigurarGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub AsignarColor()
On Error Resume Next

    Dim i, j, k As Integer, vCantidadStock As Long, vColor As Long
    
    With Detalle
        
        barra.Value = 0
        barra.Max = .Rows - 1
        
        'Recorro las Filas
        i = 0
        For i = 1 To .Rows - 1
            DoEvents
            .Row = i

            vCantidadStock = Val(.TextMatrix(.Row, 3))
            
            For k = 0 To vCantidadRangos
                If vCantidadStock >= vRango(0, k) And vCantidadStock <= vRango(1, k) Then
                    vColor = vRango(2, k)
                    Exit For
                End If
            Next
            
            If vColor = 0 Then vColor = &HFFFFFF
            
            
            'Recorro las Columnas por fila
            j = 0
            For j = 1 To .Cols - 1
                .Col = j
                .CellBackColor = vColor
            Next

            .TextMatrix(.Row, 4) = Val(Format(.TextMatrix(.Row, 4), "######0.00"))
            .TextMatrix(.Row, 5) = Val(Format(.TextMatrix(.Row, 5), "######0.00"))
            .TextMatrix(.Row, 6) = Val(Format(.TextMatrix(.Row, 6), "######0.00"))
            
            barra.Value = .Row
        Next
    
    End With
    
If Err Then GrabarLog "AsignarColor", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next

    'If Not vIdUsuarioNivel = 1 Then ControlarPermisos
    Me.Show

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarRangos() 'Cargo los Rangos de Cantidad y el color asignado en un Array
On Error Resume Next
    
    Dim rsStockColores As New ADODB.Recordset, sqlStockColores As String
    Dim i As Integer
    
    sqlStockColores = "SELECT * FROM StockColores"
    
    With rsStockColores
        Call .Open(sqlStockColores, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        vCantidadRangos = .RecordCount - 1
        
        ReDim vRango(2, vCantidadRangos)
    
            
        If Not .EOF = True Then .MoveFirst

        i = 0
        
        Do Until .EOF = True
            vRango(0, i) = .Fields("CantidadMin").Value
            vRango(1, i) = .Fields("CantidadMax").Value
            vRango(2, i) = .Fields("Color").Value
            
            BarraEstado.Panels.Add
            BarraEstado.Panels(i + 1).Width = 1850
            BarraEstado.Panels(i + 1).Text = "   Min: " & vRango(0, i) & " - Max: " & vRango(1, i) & ""
            
            Call CrearControl(i, vRango(2, i))
            
            .MoveNext
            i = i + 1
        Loop
        
    
    End With
    
    sqlStockColores = ""

    If rsStockColores.State = 1 Then
        rsStockColores.Close
        Set rsStockColores = Nothing
    End If
    
If Err Then GrabarLog "CargarRangos", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub PBRango_Click()
On Error Resume Next

    frmStockColores.Show
    
If Err Then GrabarLog "PBRango_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtCodigo_KeyPress(Keyascii As Integer)
On Error Resume Next
    
        If Keyascii = 13 Then txtDescripcion.SetFocus
    
If Err Then GrabarLog "txtCodigo_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub txtDescripcion_KeyPress(Keyascii As Integer)
On Error Resume Next
    
    If Keyascii = 13 Then cmdEjecutar.SetFocus
    
If Err Then GrabarLog "txtDescripcion_KeyPress", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CrearControl(vIndex, vColor)
    On Error Resume Next
    
    Dim i As Integer
    
    Dim PicColor
    
    Set PicColor = Controls.Add("VB.PictureBox", "Pic" & vIndex)

    With PicColor
        .Visible = True
        .BackColor = vColor
        .Width = 200
        .Height = BarraEstado.Height
        .Top = BarraEstado.Top + 25
        .ZOrder (0)
        .Left = BarraEstado.Panels(vIndex + 1).Left
    End With

End Sub



