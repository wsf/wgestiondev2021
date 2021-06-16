VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "Copia de KlexGrid.ocx"
Begin VB.Form frmConsultas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Control de inconsistencia de datos"
   ClientHeight    =   8505
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11040
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11040
   WindowState     =   2  'Maximized
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   525
      Left            =   -30
      TabIndex        =   4
      Top             =   -30
      Width           =   17475
      _Version        =   851968
      _ExtentX        =   30824
      _ExtentY        =   926
      _StockProps     =   79
      BackColor       =   -2147483644
      Appearance      =   1
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PusStatus 
         Height          =   285
         Left            =   10620
         TabIndex        =   13
         Top             =   180
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Status"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   315
         Left            =   60
         TabIndex        =   6
         Top             =   150
         Width           =   1545
         _Version        =   851968
         _ExtentX        =   2725
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Agrear datos"
         BackColor       =   -2147483626
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConsultas.frx":0000
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   315
         Left            =   3690
         TabIndex        =   5
         Top             =   150
         Visible         =   0   'False
         Width           =   2235
         _Version        =   851968
         _ExtentX        =   3942
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Actualizar Modificaciones"
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConsultas.frx":0B4A
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   315
         Left            =   1650
         TabIndex        =   7
         Top             =   150
         Visible         =   0   'False
         Width           =   1545
         _Version        =   851968
         _ExtentX        =   2725
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Borrar"
         BackColor       =   -2147483626
         UseVisualStyle  =   -1  'True
         Picture         =   "frmConsultas.frx":155C
      End
      Begin VB.Label lblINSERTPara 
         Caption         =   "<INSERT> para modificar campo editado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6630
         TabIndex        =   8
         Top             =   180
         Width           =   3765
      End
   End
   Begin XtremeSuiteControls.GroupBox GroCondicionesDe 
      Height          =   555
      Left            =   -60
      TabIndex        =   2
      Top             =   420
      Width           =   17505
      _Version        =   851968
      _ExtentX        =   30877
      _ExtentY        =   979
      _StockProps     =   79
      BackColor       =   -2147483645
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
      Begin XtremeSuiteControls.FlatEdit vbuscando 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   180
         Width           =   6855
         _Version        =   851968
         _ExtentX        =   12091
         _ExtentY        =   556
         _StockProps     =   77
      End
      Begin XtremeSuiteControls.FlatEdit vbuscar2 
         Height          =   315
         Left            =   9540
         TabIndex        =   9
         Top             =   135
         Width           =   3300
         _Version        =   851968
         _ExtentX        =   5821
         _ExtentY        =   556
         _StockProps     =   77
      End
      Begin XtremeSuiteControls.FlatEdit vnro 
         Height          =   315
         Left            =   13455
         TabIndex        =   11
         Top             =   135
         Width           =   4020
         _Version        =   851968
         _ExtentX        =   7091
         _ExtentY        =   556
         _StockProps     =   77
      End
      Begin XtremeSuiteControls.Label lblNro 
         Height          =   285
         Left            =   13005
         TabIndex        =   12
         Top             =   180
         Width           =   465
         _Version        =   851968
         _ExtentX        =   820
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Nro:"
         ForeColor       =   -2147483634
         BackColor       =   -2147483636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin XtremeSuiteControls.Label lblDirec 
         Height          =   285
         Left            =   9000
         TabIndex        =   10
         Top             =   180
         Width           =   465
         _Version        =   851968
         _ExtentX        =   820
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Direc.:"
         ForeColor       =   -2147483634
         BackColor       =   -2147483636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar campo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   180
         TabIndex        =   3
         Top             =   225
         Width           =   1260
      End
   End
   Begin Grid.KlexGrid grid 
      Height          =   7965
      Left            =   0
      TabIndex        =   1
      Top             =   990
      Width           =   17475
      _ExtentX        =   30824
      _ExtentY        =   14049
      EnterKeyBehaviour=   2
      BackColorAlternate=   14737632
      GridLinesFixed  =   2
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      BackColor       =   16777215
      BackColorBkg    =   16777215
      BackColorFixed  =   -2147483638
      BackColorSel    =   49152
      BorderStyle     =   0
      Cols            =   5
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      ForeColorSel    =   255
      GridColor       =   14737632
      GridColorFixed  =   8421504
      MergeCells      =   1
      MouseIcon       =   "frmConsultas.frx":1F6E
      Rows            =   10
      SelectionMode   =   1
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public bcontrol As ADODB.Recordset
Public vcontrol As String
Public vform As Form
Public vtabla As String
Public vcampoMostrar As String, vcampoID As String, bandera As String
Public vsql, vgsql As String
Dim vmodificar, vcol, vrow As Integer
Public vcampo2 As String
Public enComuna As Boolean
Dim vidArticulos As Long
Dim vulinea As Integer
Dim i As Integer


Private Sub pintar(i As Integer, g As KlexGrid)
On Error Resume Next
Dim j, k, kk As Integer

k = g.Row
kk = g.Col

g.Row = i

For j = 1 To g.Cols - 1
    g.Col = j
    g.CellBackColor = vbGreen
Next

g.Row = k
g.Col = kk
If Err Then Exit Sub
End Sub

Private Sub despintar(i As Integer, g As KlexGrid)
On Error Resume Next

Dim j, k, kk As Integer
k = g.Row
kk = g.Col
If i = 0 Then Exit Sub
g.Row = i

For j = 1 To g.Cols - 1
    g.Col = j
    g.CellBackColor = vbWhite
Next

g.Row = k
g.Col = kk

If Err Then Exit Sub
End Sub

Public Sub Buscar(ByVal vsql As String)
On Error Resume Next

'If Not Len(vsql) > 3 Then Exit Sub

Dim bcontrol As New ADODB.Recordset


With bcontrol
        If .State = 1 Then .Close

        .CursorLocation = adUseServer
        
       If Me.enComuna Then
            
           ' Dim ConnComunaDB3 As New ADODB

            ConnComunaDB.ConnectionString = LeerXml("ComunaCnn")
            'ConnComunaDB3.ConnectionString = pathDBMySQL
       
           Call .Open(vsql, ConnComunaDB, adOpenDynamic, adLockPessimistic)
       
       Else
             Call .Open(vsql, ConnDDBB, adOpenDynamic, adLockPessimistic)
  
       End If
  
  End With


    Set Me.grid.Recordset = bcontrol

If Not Me.vbuscar2.Text = "" Then
    vbuscar2.SetFocus
Else
    vbuscando.SetFocus
End If

If Err Then Exit Sub
  End Sub
   
  
Private Sub finit()
On Error Resume Next
vsql = "select * from " + vtabla + " order by " + vcampoMostrar
vsql = ""
Call Buscar(vsql)
If Err Then Exit Sub
End Sub

Private Sub Form_LinkError(LinkErr As Integer)
'MsgBox "xx"
End Sub

Private Sub Form_Load()
finit
'Call vbuscando.SetFocus
End Sub

Private Sub grid_BeforeEdit(Cancel As Boolean)
  MsgBox "No puede editar"
    Me.vbuscando.SetFocus
    Exit Sub
End Sub

Private Sub grid_Click()
On Error Resume Next
'Call vbuscando_Change

'pintarFilafd

'vidArticulos = grid.TextMatrix(grid.Row, grid.Cols - 1)

'Call pintar(grid.Row, grid)

'Call despintar(vulinea, grid)

'grid.CellBackColor = vbRed

grid.SetFocus
vulinea = grid.Row
grid.RowSel = grid.Row
grid.SetFocus
 'grid.RowPosition(vulinea) = vulinea

'Call grid_Click
'Me.vbuscando.SetFocus


If Err Then Exit Sub
End Sub

Private Sub grid_DblClick()

'grid.SetFocus
'grid.RowSel = vulinea

'Call grid_GotFocus



If vulinea = 1 And bandera = "" Then Exit Sub


bandera = ""

If Me.grid.Row = 0 Then
    Call frmAlert.DisplayAlert("No se hay elemento para seleccionar", 1000)
    Me.WindowState = 2
    Exit Sub
End If
vcampoMostrar = Me.grid.TextMatrix(Me.grid.Row, fncolTopos(vcampoMostrar)) ' tomo el valor del campo a mostrar

If Not vcampo2 = "" Then
    vcampoMostrar = Trim(Me.grid.TextMatrix(Me.grid.Row, fncolTopos(vcampo2))) + ", " + Trim(vcampoMostrar)
End If
vcampoID = Me.grid.TextMatrix(Me.grid.Row, fncolTopos(vcampoID)) ' tomo el valor del id



Dim vc As Integer
vc = Val(vcontrol)
'vc = Val(vcampoID)
If vc > 0 Then

    vform.Controls(vc).Tag = vcampoID
    vform.Controls(vc) = vcampoMostrar

Else
    vform.Controls(vcontrol).Tag = vcampoID
    vform.Controls(vcontrol) = vcampoMostrar
End If

vform.Show
vform.WindowState = vmaximizar

'vform.vform.Controls(vcontrol).GotFocus

Unload Me
End Sub


Function fncolTopos(vcampo As String) As Integer    ' devuelve la columna del campo con nombre vcampo
Dim i As Integer


For i = 0 To grid.Cols - 1

    If Trim(grid.TextMatrix(0, i)) = Trim(vcampo) Then fncolTopos = i

Next

End Function



Private Sub grid_KeyDown(KeyCode As Integer, Shift As Integer)
  MsgBox "No puede editar"
    Me.vbuscando.SetFocus
    Exit Sub
End Sub

Private Sub grid_KeyPress(KeyAscii As Integer)
On Error Resume Next

  MsgBox "No puede editar"
    Me.vbuscando.SetFocus
    Exit Sub

If Err Then
    MsgBox "No puede editar"
    Me.vbuscando.SetFocus
    Exit Sub
End If
End Sub

Private Sub grid_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
On Error Resume Next
'Dim w1, w2 As Integer

'Exit Sub

'vsql = "update " + Me.vtabla + " set " + grid.TextMatrix(0, Col) + " = '" + grid.TextMatrix(Row, Col) + "' where " + grid.TextMatrix(0, 1) + "=" + grid.TextMatrix(Row, 1)
'vmodificar = 1
'vcol = Col
'vrow = Row

If Err Then Exit Sub
End Sub

Private Sub grid_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
'On Error Resume Next

'Exit Sub

'If Not KeyCode = vbKeyInsert Then
'   Exit Sub
'End If

'If vmodificar = 1 Then
'vsql = "update " + Me.vtabla + " set " + grid.TextMatrix(0, vcol) + " = '" + grid.TextMatrix(vrow, vcol) + "' where " + grid.TextMatrix(0, 1) + "=" + grid.TextMatrix(vrow, 1)
'
'    Call EjecutarScript(vsql, pathDBMySQL)
'End If
'
'vmodificar = 0
'
'MsgBox "Los dato fueron modificados", vbInformation

'MsgBox "No puede editar"
vbuscando.SetFocus
If Err Then
    
    If Me.vbuscar2.Tag = "foco" Then
        vbuscar2.SetFocus
    Else
        vbuscando.SetFocus
    End If
    
    Exit Sub
End If
End Sub

Private Sub PushButton2_Click()
On Error Resume Next
Dim vsql As String

    
Select Case vtabla

     Case "proveedores"
        frmProveedoresAlta.Show
        frmProveedoresAlta.txtAlta(1).Text = Me.vbuscando
        frmProveedoresAlta.viente = "frmConsultas"


Case "(select * from proveedores where tipoproveedor='Eventuales') as p"

        frmProveedoresAlta.Show
        frmProveedoresAlta.txtAlta(1).Text = Me.vbuscando
        frmProveedoresAlta.viente = "frmConsultas"

Case "(select * from articulos order by Descrip) as t"

        frmArticulosAlta.Show
        frmArticulosAlta.txtAlta(1).Text = Me.vbuscando
        frmArticulosAlta.vViene = "frmConsultas"



End Select
    
    
               
             '   vsql = "INSERT INTO " + vtabla + " () VALUES()"
             '   Call EjecutarScript(vsql, pathDBMySQL)
                
             '   Call vbuscando_Change


If Err Then Exit Sub
End Sub

Private Sub PushButton3_Click()
On Error Resume Next
Dim vsql As String

If MsgBox("Está seguro de borrar la fila", vbYesNo, "Borrar") = vbNo Then Exit Sub

vsql = "delete from " + vtabla + " where " + Me.grid.TextMatrix(0, 1) + "=" + Me.grid.TextMatrix(grid.RowSel, 1)
Call EjecutarScript(vsql, pathDBMySQL)

Call vbuscando_Change

If Err Then Exit Sub
End Sub


Private Sub PusStatus_Click()
MsgBox "ConnComunaDB.ConnectionString" + Chr(13) + ConnComunaDB.ConnectionString
End Sub

Public Sub vbuscando_Change()
On Error Resume Next
Dim vwhere As String
Dim vbuscando2 As String


vbuscando2 = vbuscando



If Me.enComuna Then

    If Val(vbuscando) > 0 Then
    
        vwhere = " where " + vcampo2 + " like '" + vbuscando + "%' or " + vcampoMostrar + " like '%" + vbuscando2 + "%'"
        vsql = "select * from " + vtabla + " " + vwhere + " order by " + vcampoMostrar
 
    Else
    
        vwhere = " where " + vcampoMostrar + " like '%" + vbuscando + "%' or " + vcampo2 + " like '%" + vbuscando + "%' or " + vcampoMostrar + " like '%" + vbuscando2 + "%'"
        vsql = "select * from " + vtabla + " " + vwhere + " order by " + vcampoMostrar
    End If
Else
    
    
    
   ' If Val(vbuscando) > 0 And vcampoID = "idCuentas" Then
    
    If Val(vbuscando) > 0 And vcampoID = "CodigoCuenta" Then
         
             vwhere = " where " + vcampoID + " like '" + vbuscando + "%' or replace(REPLACE(" + vcampoID + ", '.', ''),'0','')  like '%" + vbuscando2 + "%'"
            vsql = "select * from " + vtabla + " " + vwhere + " order by " + " CAST(" + vcampoID + " AS UNSIGNED)"

    
    
    
    Else
    
    
        If Not Val(vbuscando) > 0 Then
       
            vwhere = " where " + vcampoID + " like '%" + vbuscando + "%' or " + vcampoMostrar + " like '%" + vbuscando + "%' or replace(REPLACE(" + vcampoID + ", '.', ''),'0','')  like '%" + vbuscando2 + "%'"
            vsql = "select * from " + vtabla + " " + vwhere + " order by " + vcampoMostrar

        Else
            vwhere = " where " + vcampoID + " = '" + vbuscando + "'" ' or " + vcampoMostrar + " like '%" + vbuscando + "%' or replace(REPLACE(" + vcampoID + ", '.', ''),'0','')  like '" + vbuscando2 + "%'"
            vsql = "select * from " + vtabla + " " + vwhere + " order by " + vcampoMostrar
        End If
        
    End If
    
    
    
    
    
   
End If

'vsql = "select * from " + vtabla + " " + vwhere + " order by " + vcampoMostrar


Call Buscar(vsql + " limit 30")
'pintarFila
bandera = ""

vbuscando.SetFocus
If Err Then Exit Sub
End Sub

Private Sub vbuscando_GotFocus()
vbuscar2.Tag = ""
vbuscar2.BackColor = vbWhite
vbuscando.BackColor = vbYellow
End Sub

Private Sub vbuscando_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = 13 Then
    grid.Row = 1
    bandera = "enter"
    Call grid_DblClick
End If

If Err Then Exit Sub
End Sub



Private Sub pintarFila()
On Error Resume Next
Dim i As Integer


For i = 0 To grid.Cols - 1
    grid.Col = i
    grid.CellBackColor = vbGreen
    'If Trim(grid.TextMatrix(0, i)) = Trim(vcampo) Then fncolTopos = i

Next


If Err Then Exit Sub
End Sub

Private Sub vbuscando_LostFocus()
vbuscando = UCase(vbuscando)
'grid.SetFocus
'Call grid_Click
End Sub

Private Sub vbuscar2_Change()
On Error Resume Next

Dim vwhere As String
Dim vbuscando2 As String


' vbuscando2 = vbuscando

If Me.enComuna Then
  
        vwhere = " where direccion like '%" + vbuscar2 + "%'"
        vsql = "select * from " + vtabla + " " + vwhere + " order by direccion"
End If

'vsql = "select * from " + vtabla + " " + vwhere + " order by " + vcampoMostrar

Call Buscar(vsql + " limit 30")
'pintarFila
'bandera = ""

vbuscar2.SetFocus
If Err Then Exit Sub

End Sub

Private Sub vbuscar2_GotFocus()
vbuscar2.Tag = "foco"
vbuscar2.BackColor = vbYellow
vbuscando.BackColor = vbWhite
End Sub

Private Sub vbuscar2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grid.Row = 1
    bandera = "enter"
    Call grid_DblClick
End If
End Sub

Private Sub vnro_Change()
On Error Resume Next

Dim vwhere As String
Dim vbuscando2 As String

If Me.enComuna Then
  
        vwhere = " where id_contribuyentes like '%" + vnro + "%'"
        vsql = "select * from " + vtabla + " " + vwhere + " order by 1"
End If

Call Buscar(vsql + " limit 30")


vnro.SetFocus

If Err Then Exit Sub

End Sub

Private Sub vnro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grid.Row = 1
    bandera = "enter"
    Call grid_DblClick
End If

End Sub
