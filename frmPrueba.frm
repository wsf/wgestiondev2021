VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmPrueba 
   Caption         =   "presupuesto"
   ClientHeight    =   4470
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   6675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   4470
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   300
      Left            =   5340
      TabIndex        =   0
      Top             =   3960
      Width           =   1080
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      DragIcon        =   "frmPrueba.frx":0000
      Height          =   3840
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   6360
      _ExtentX        =   11218
      _ExtentY        =   6773
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   0
      Rows            =   4
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   8421376
      ForeColorFixed  =   16777215
      GridColor       =   8421504
      GridColorFixed  =   0
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      MergeCells      =   4
      AllowUserResizing=   1
      FormatString    =   "idcuentas|idpresupuesto|periodo|importe"
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
      _Band(0).GridLineWidthBand=   1
      _Band(0).TextStyleBand=   0
   End
End
Attribute VB_Name = "frmPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MARGIN_SIZE = 60      ' en twips
' variables para enlace de datos
Private datPrimaryRS As ADODB.Recordset

' variables para arrastrar columnas
Private m_bDragOK As Boolean
Private m_iDragCol As Integer
Private xdn As Integer, ydn As Integer

Private Sub Form_Load()

    Dim sConnect As String
    Dim sSQL As String
    Dim dfwConn As ADODB.Connection
    Dim i As Integer

    ' establecer cadenas
    sConnect = "Provider=MSDASQL.1;Extended Properties='DATABASE=wgestioncomuna;DSN=wgestioncomuna;OPTION=3;PWD=dalas.2009;PORT=3306;SERVER=localhost;UID=root;'"
    sSQL = "select idcuentas,idpresupuesto,periodo,importe from presupuesto"

    ' abrir conexión
    Set dfwConn = New Connection
    dfwConn.Open sConnect

    ' crear un conjunto de registros con la colección proporcionada
    Set datPrimaryRS = New Recordset
    datPrimaryRS.CursorLocation = adUseClient
    datPrimaryRS.Open sSQL, dfwConn, adOpenForwardOnly, adLockReadOnly

    Set MSHFlexGrid1.DataSource = datPrimaryRS

    With MSHFlexGrid1

        .Redraw = False
        ' establecer anchos de columna de cuadrícula
        .ColWidth(0) = -1
        .ColWidth(1) = -1
        .ColWidth(2) = -1
        .ColWidth(3) = -1

        ' establecer combinación y orden de columna de cuadrícula
        For i = 0 To .Cols - 1
            .MergeCol(i) = True
        Next i

        .Sort = flexSortGenericAscending

        ' establecer tipo de cuadrícula
        .AllowBigSelection = True
        .FillStyle = flexFillRepeat

        ' encabezado en negrita
        .Row = 0
        .Col = 0
        .RowSel = .FixedRows - 1
        .ColSel = .Cols - 1
        .CellFontBold = True

        ' atenuar otra columna
        For i = .FixedCols To .Cols() - 1 Step 2
            .Col = i
            .Row = .FixedRows
            .RowSel = .Rows - 1
            .CellBackColor = &HC0C0C0   ' gris claro
        Next i

        .AllowBigSelection = False
        .FillStyle = flexFillSingle
        .Redraw = True

    End With

End Sub

Private Sub MSHFlexGrid1_DragDrop(Source As Control, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' el código de los eventos DragDrop, MouseDown, MouseMove y MouseUp permite arrastrar columnas
'-------------------------------------------------------------------------------------------

    If m_iDragCol = -1 Then Exit Sub    ' no se estaba arrastrando
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub

    With MSHFlexGrid1
        .Redraw = False
        .ColPosition(m_iDragCol) = .MouseCol

        .FillStyle = flexFillRepeat
        .Col = 0
        .Row = .FixedRows
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .CellBackColor = &HFFFFFF
        Dim iLoop As Integer
        For iLoop = .FixedCols To .Cols() - 1 Step 2
            .Col = iLoop
            .Row = .FixedRows
            .RowSel = .Rows - 1
            .CellBackColor = &HC0C0C0
        Next iLoop
        .FillStyle = flexFillSingle

        DoSort
        .Redraw = True
    End With

End Sub

Private Sub MSHFlexGrid1_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' el código de los eventos DragDrop, MouseDown, MouseMove y MouseUp permite arrastrar columnas
'-------------------------------------------------------------------------------------------

    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub

    xdn = X
    ydn = Y
    m_iDragCol = -1     ' borrar indicador de arrastre
    m_bDragOK = True

End Sub

Private Sub MSHFlexGrid1_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' el código de los eventos DragDrop, MouseDown, MouseMove y MouseUp permite arrastrar columnas
'-------------------------------------------------------------------------------------------

    ' probar si se debe iniciar el arrastre
    If Not m_bDragOK Then Exit Sub
    If Button <> 1 Then Exit Sub                        ' botón incorrecto
    If m_iDragCol <> -1 Then Exit Sub                   ' ya se está arrastrando
    If Abs(xdn - X) + Abs(ydn - Y) < 50 Then Exit Sub   ' no se ha movido suficiente
    If MSHFlexGrid1.MouseRow <> 0 Then Exit Sub         ' hay que arrastrar el encabezado

    ' si se llega aquí, iniciar el arrastre
    m_iDragCol = MSHFlexGrid1.MouseCol
    MSHFlexGrid1.Drag vbBeginDrag

End Sub

Private Sub MSHFlexGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
'-------------------------------------------------------------------------------------------
' el código de los eventos DragDrop, MouseDown, MouseMove y MouseUp permite arrastrar columnas
'-------------------------------------------------------------------------------------------

    m_bDragOK = False

End Sub

Sub DoSort()

    With MSHFlexGrid1
        .Redraw = False
        .Col = 0
        .Row = 1
        .RowSel = .Rows - 1
        .Sort = flexSortGenericAscending
        .Redraw = True
    End With

End Sub

Private Sub Form_Resize()

    Dim sngButtonTop As Single
    Dim sngScaleWidth As Single
    Dim sngScaleHeight As Single

    On Error GoTo Form_Resize_Error
    With Me
        sngScaleWidth = .ScaleWidth
        sngScaleHeight = .ScaleHeight

        ' mueve el botón Cerrar a la esquina superior derecha
        With .cmdClose
                sngButtonTop = sngScaleHeight - (.Height + MARGIN_SIZE)
                .Move sngScaleWidth - (.Width + MARGIN_SIZE), sngButtonTop
        End With

        .MSHFlexGrid1.Move MARGIN_SIZE, _
            MARGIN_SIZE, _
            sngScaleWidth - (2 * MARGIN_SIZE), _
            sngButtonTop - (2 * MARGIN_SIZE)

    End With
    Exit Sub

Form_Resize_Error:
    ' evita errores en valores negativos
    Resume Next

End Sub
Private Sub cmdClose_Click()

    Unload Me

End Sub


