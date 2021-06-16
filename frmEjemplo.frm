VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Begin VB.Form frmEjemplo 
   Caption         =   "Form1"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   8130
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "reporte"
      Height          =   435
      Left            =   3120
      TabIndex        =   15
      Top             =   120
      Width           =   1905
   End
   Begin XtremeSuiteControls.PushButton PushButton4 
      Height          =   345
      Left            =   6990
      TabIndex        =   14
      Top             =   7320
      Width           =   1935
      _Version        =   851968
      _ExtentX        =   3413
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "PushButton4"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   345
      Left            =   4980
      TabIndex        =   13
      Top             =   7320
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "PushButton3"
      UseVisualStyle  =   -1  'True
   End
   Begin MSComDlg.CommonDialog file 
      Left            =   450
      Top             =   990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   315
      Left            =   2910
      TabIndex        =   12
      Top             =   7350
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "PushButton2"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btn_idSeleccionado 
      Height          =   315
      Left            =   8700
      TabIndex        =   10
      Top             =   1560
      Width           =   2205
      _Version        =   851968
      _ExtentX        =   3889
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Btn_id seleccionado"
      UseVisualStyle  =   -1  'True
   End
   Begin Grid.KlexGrid grd_Grilla 
      Height          =   2895
      Left            =   2970
      TabIndex        =   6
      Top             =   3810
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   5106
      GridLinesFixed  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmEjemplo.frx":0000
   End
   Begin XtremeSuiteControls.PushButton btn_Guardar 
      Height          =   315
      Left            =   3150
      TabIndex        =   5
      Top             =   3360
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Btn_ guardar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btn_Boton 
      Height          =   315
      Left            =   7440
      TabIndex        =   4
      Top             =   2670
      Width           =   945
      _Version        =   851968
      _ExtentX        =   1667
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Btn_ boton"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.FlatEdit txt_textBox1 
      Height          =   375
      Left            =   3180
      TabIndex        =   2
      Top             =   2070
      Width           =   5295
      _Version        =   851968
      _ExtentX        =   9340
      _ExtentY        =   661
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "Txt_text box"
   End
   Begin XtremeSuiteControls.ComboBox cmb_Combo 
      Height          =   315
      Left            =   3180
      TabIndex        =   0
      Top             =   1560
      Width           =   5355
      _Version        =   851968
      _ExtentX        =   9446
      _ExtentY        =   556
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "Cmb_ combo"
   End
   Begin XtremeSuiteControls.FlatEdit txt_Texbox2 
      Height          =   345
      Left            =   3180
      TabIndex        =   3
      Top             =   2640
      Width           =   4065
      _Version        =   851968
      _ExtentX        =   7170
      _ExtentY        =   609
      _StockProps     =   77
      BackColor       =   -2147483643
      Text            =   "Txt_ texbox2"
   End
   Begin XtremeSuiteControls.PushButton btn_Borrar 
      Height          =   315
      Left            =   2940
      TabIndex        =   7
      Top             =   6960
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Btn_ borrar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btn_listar 
      Height          =   315
      Left            =   4980
      TabIndex        =   8
      Top             =   6960
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Btn_listar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton btn_modificar 
      Height          =   315
      Left            =   6990
      TabIndex        =   9
      Top             =   6960
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Btn_modificar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   315
      Left            =   9060
      TabIndex        =   11
      Top             =   6960
      Width           =   1965
      _Version        =   851968
      _ExtentX        =   3466
      _ExtentY        =   556
      _StockProps     =   79
      Caption         =   "Btn_mostrar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.Label Label1 
      Height          =   495
      Left            =   3180
      TabIndex        =   1
      Top             =   840
      Width           =   5745
      _Version        =   851968
      _ExtentX        =   10134
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Label1"
   End
End
Attribute VB_Name = "frmEjemplo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_Borrar_Click()

Dim vsql As String

Call t_borrarFila(Me.grd_Grilla.TextMatrix(Me.grd_Grilla.Row, 1), "clientes")

vsql = "Select idClientes,Nombre, Direccion from clientes order by idClientes desc"
Call LlenarGrilla2(Me.grd_Grilla, vsql, "1")

End Sub

Private Sub btn_Boton_Click()

Call fbuscarGrilla("clientes", "Nombre", "idClientes", Me.txt_Texbox2.Name, Me) ' ema:

End Sub

Private Sub btn_Guardar_Click()
'fValidacion
fGuardar
fMostrarEnGrilla
End Sub

Private Sub btn_idSeleccionado_Click()
MsgBox cmb_Combo.ItemData(Me.cmb_Combo.ListIndex)
End Sub

Private Sub Command1_Click()
Call prueba.PrintReport(False, rptRangeAllPages)
'prueba.Show
End Sub

Private Sub Form_Load()
finit

End Sub
Private Sub fMostrarEnGrilla()
Dim vsql As String

vsql = "Select idClientes,Nombre, Direccion from clientes order by idClientes desc"

Call LlenarGrilla2(Me.grd_Grilla, vsql, "1")

End Sub


Private Sub finit()
' llenar de datos el combo
Call CargarComboNew2("clientes", "idClientes", "Nombre", Me.cmb_Combo, False)

End Sub


Private Sub fGuardar()

Dim vcampos As String
Dim vvalores As String
Dim vsql As String

vcampos = "Nombre,Codigo,Direccion"
vvalores = "'" + Me.txt_Texbox2 + "'" + "," + "'2'" + "," + "'" + Me.txt_textBox1 + "'"


vsql = "insert into clientes (" + vcampos + ") values (" + vvalores + ")"

Call EjecutarScript(vsql, pathDBMySQL)

End Sub

Private Sub PushButton1_Click()
fMostrarEnGrilla
End Sub

Private Sub PushButton2_Click()
On Error Resume Next
    
    file.ShowOpen

    strDir = App.Path
    strFile = file.FileName
    
    If Not strFile = "" Then
    

        strWkShtName = TraerPropiedadExcel(strFile, "")
        strWkShtName = InputBox("Ingresar nombre de la Hoja de Excel :", "Mensaje ...", strWkShtName)
        
        Dim xlwbook As Excel.Workbook
        Dim xl As New Excel.Application
        Dim xlSheet As Excel.Worksheet

                    
        Set xlwbook = xl.Workbooks.Open(strFile)
        Set xlSheet = xlwbook.Sheets.Item(strWkShtName)

        ExcelToGrid xlSheet, xlwbook


    End If
    
    If Err Then GrabarLog "cmdCargarLista_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub


Function ExcelToGrid(xls As Worksheet, wxls As Excel.Workbook) As Worksheet

    Dim nxlsheet As New Excel.Worksheet
    
    'Set xlwbook = xl.Workbooks.Open(strFile)
    Set nxlsheet = wxls.Sheets.Add()
                          
    Dim fila, i, j As Long
    fila = xls.Range("B65000").End(xlUp).Row
    
    Dim rsExcelExport As New ADODB.Recordset
       
    Dim cmdExcelExport As New ADODB.Command
    
    cmdExcelExport.ActiveConnection = ConnDDBB
    Dim sqlExcelExport, sqlBorrar As String
    
    sqlExcelExport = "Select * from ExcelExport"
    sqlBorrar = "Delete from ExcelExport"
    
    If rsExcelExport.State = 0 Then
        rsExcelExport.Open sqlBorrar, ConnDDBB, adOpenKeyset, adLockOptimistic
        rsExcelExport.Open sqlExcelExport, ConnDDBB, adOpenKeyset, adLockOptimistic
    End If
    
    If Not rsExcelExport.BOF Then
        rsExcelExport.MoveFirst
    End If
    
    BarraExcel.Max = fila
    BarraExcel.Min = 0
    BarraExcel.Value = 0
              
    vdisplay.AddItem ("Cantidad de artículos: " + Str(fila))
              
    For i = 1 To fila
        On Error Resume Next
        If (IsNumeric(Trim(xls.Range("A" & i).Value)) Or IsNumeric(Trim(xls.Range("B" & i).Value)) Or IsNumeric(Trim(xls.Range("C" & i).Value)) Or IsNumeric(Trim(xls.Range("D" & i).Value)) Or IsNumeric(Trim(xls.Range("E" & i).Value)) Or IsNumeric(Trim(xls.Range("F" & i).Value)) Or IsNumeric(Trim(xls.Range("G" & i).Value)) Or IsNumeric(Trim(xls.Range("H" & i).Value)) Or IsNumeric(Trim(xls.Range("I" & i).Value)) Or IsNumeric(Trim(xls.Range("J" & i).Value)) Or IsNumeric(Trim(xls.Range("K" & i).Value)) Or IsNumeric(Trim(xls.Range("L" & i).Value)) Or IsNumeric(Trim(xls.Range("M" & i).Value)) Or IsNumeric(Trim(xls.Range("N" & i).Value)) Or IsNumeric(Trim(xls.Range("O" & i).Value)) Or IsNumeric(Trim(xls.Range("P" & i).Value)) Or IsNumeric(Trim(xls.Range("Q" & i).Value)) Or IsNumeric(Trim(xls.Range("R" & i).Value)) Or IsNumeric(Trim(xls.Range("S" & i).Value)) Or IsNumeric(Trim(xls.Range("T" & i).Value)) Or IsNumeric(Trim(xls.Range("U" & i).Value)) Or IsNumeric(Trim(xls.Range("V" & i).Value))) Then
                                 
            rsExcelExport.AddNew
            
            For j = 1 To 35
                rsExcelExport(j) = xls.Cells(i, j)
            Next j
            
            rsExcelExport.Update
            DoEvents
        End If
       ' BarraExcel.Value = Me.BarraExcel.Value + 1
     Next i
      
      With rsExcel
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            
            Call .Open("SELECT * FROM ExcelExport", ConnDDBB, adOpenStatic, adLockPessimistic)

            If Not .State = 1 Then
                MsgBox Err.Description
                Exit Function
            Else
                Set dgExcel.DataSource = rsExcel
            End If
      End With
    
    
End Function

Private Sub PushButton3_Click()
Dim oXL As Excel.Application
      Dim oWB As Excel.Workbook
      Dim oSheet As Excel.Worksheet
      Dim oRng As Excel.Range
      

      'On Error GoTo Err_Handler
      
   ' Start Excel and get Application object.
      Set oXL = CreateObject("Excel.Application")
      oXL.Visible = True
      
   ' Get a new workbook.
      Set oWB = oXL.Workbooks.Add
      Set oSheet = oWB.ActiveSheet
      
   ' Add table headers going cell by cell.
      oSheet.Cells(1, 1).Value = "First Name"
      oSheet.Cells(1, 2).Value = "Last Name"
      oSheet.Cells(1, 3).Value = "Full Name"
      oSheet.Cells(1, 4).Value = "Salary"
      

   ' Format A1:D1 as bold, vertical alignment = center.
      With oSheet.Range("A1", "D1")
         .Font.Bold = True
         .VerticalAlignment = xlVAlignCenter
      End With
      
   ' Create an array to set multiple values at once.
      Dim saNames(5, 2) As String
      saNames(0, 0) = "John"
      saNames(0, 1) = "Smith"
      saNames(1, 0) = "Tom"
      saNames(1, 1) = "Brown"
      saNames(2, 0) = "Sue"
      saNames(2, 1) = "Thomas"
      saNames(3, 0) = "Jane"

      saNames(3, 1) = "Jones"
      saNames(4, 0) = "Adam"
      saNames(4, 1) = "Johnson"
      
    ' Fill A2:B6 with an array of values (First and Last Names).
      oSheet.Range("A2", "B6").Value = saNames
      
    ' Fill C2:C6 with a relative formula (=A2 & " " & B2).
      Set oRng = oSheet.Range("C2", "C6")
      oRng.Formula = "=A2 & "" "" & B2"
      
    ' Fill D2:D6 with a formula(=RAND()*100000) and apply format.
      Set oRng = oSheet.Range("D2", "D6")
      oRng.Formula = "=RAND()*100000"
      oRng.NumberFormat = "$0.00"
      
    ' AutoFit columns A:D.
      Set oRng = oSheet.Range("A1", "D1")
      oRng.EntireColumn.AutoFit
      
    ' Manipulate a variable number of columns for Quarterly Sales Data.
      Call DisplayQuarterlySales(oSheet)
      
    ' Make sure Excel is visible and give the user control
    ' of Microsoft Excel's lifetime.
      oXL.Visible = True
      oXL.UserControl = True
      
    ' Make sure you release object references.
      Set oRng = Nothing
      Set oSheet = Nothing
      Set oWB = Nothing
      Set oXL = Nothing
      
   Exit Sub
Err_Handler:
      MsgBox Err.Description, vbCritical, "Error: " & Err.Number
   End Sub
   
   Private Sub DisplayQuarterlySales(oWS As Excel.Worksheet)
      Dim oResizeRange As Excel.Range
      Dim oChart As Excel.Chart
      Dim iNumQtrs As Integer
      Dim sMsg As String
      Dim iRet As Integer
      
    ' Determine how many quarters to display data for.
      For iNumQtrs = 4 To 2 Step -1
         sMsg = "Enter sales data for" & Str(iNumQtrs) & " quarter(s)?"
         iRet = MsgBox(sMsg, vbYesNo Or vbQuestion _
            Or vbMsgBoxSetForeground, "Quarterly Sales")
         If iRet = vbYes Then Exit For
      Next iNumQtrs
      

      sMsg = "Displaying data for" & Str(iNumQtrs) & " quarter(s)."
      MsgBox sMsg, vbMsgBoxSetForeground, "Quarterly Sales"
      
    ' Starting at E1, fill headers for the number of columns selected.
      Set oResizeRange = oWS.Range("E1", "E1").Resize(ColumnSize:=iNumQtrs)

      oResizeRange.Formula = "=""Q"" & COLUMN()-4 & CHAR(10) & ""Sales"""
      
    ' Change the Orientation and WrapText properties for the headers.
      oResizeRange.Orientation = 38
      oResizeRange.WrapText = True
      
    ' Fill the interior color of the headers.
      oResizeRange.Interior.ColorIndex = 36
      
    ' Fill the columns with a formula and apply a number format.
      Set oResizeRange = oWS.Range("E2", "E6").Resize(ColumnSize:=iNumQtrs)
      oResizeRange.Formula = "=RAND()*100"
      oResizeRange.NumberFormat = "$0.00"
      
    ' Apply borders to the Sales data and headers.
      Set oResizeRange = oWS.Range("E1", "E6").Resize(ColumnSize:=iNumQtrs)
      oResizeRange.Borders.Weight = xlThin
      
    ' Add a Totals formula for the sales data and apply a border.
      Set oResizeRange = oWS.Range("E8", "E8").Resize(ColumnSize:=iNumQtrs)
      oResizeRange.Formula = "=SUM(E2:E6)"
      With oResizeRange.Borders(xlEdgeBottom)
         .LineStyle = xlDouble
         .Weight = xlThick
      End With
      
    ' Add a Chart for the selected data
      Set oResizeRange = oWS.Range("E2:E6").Resize(ColumnSize:=iNumQtrs)
      Set oChart = oWS.Parent.Charts.Add
      With oChart
         .ChartWizard oResizeRange, xl3DColumn, , xlColumns
         .SeriesCollection(1).XValues = oWS.Range("A2", "A6")
            For iRet = 1 To iNumQtrs
               .SeriesCollection(iRet).Name = "=""Q" & Str(iRet) & """"
            Next iRet
         .Location xlLocationAsObject, oWS.Name
      End With
      
    ' Move the chart so as not to cover your data.
      With oWS.Shapes("Chart 1")
         .Top = oWS.Rows(10).Top
         .Left = oWS.Columns(2).Left

      End With
      
    ' Free any references.
      Set oChart = Nothing
      Set oResizeRange = Nothing
End Sub

Private Sub PushButton4_Click()
   Dim xlSheet As Object
   Dim xlApp As Object
   Set xlSheet = CreateObject("Excel.Sheet")
   MsgBox xlSheet.Application.Name
   Set xlApp = GetObject("Excel.Application")
   MsgBox xlApp.Name
   Set xlSheet = Nothing
   MsgBox xlApp.Name
            
End Sub
