VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmCotizacion 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Formulario de Cotizacion de Dolar"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmCotizacion.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   7125
      TabIndex        =   5
      Top             =   2640
      Width           =   7125
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   3600
         TabIndex        =   6
         Top             =   105
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCotizacion.frx":50B3
      End
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   1
         Left            =   4680
         TabIndex        =   7
         Top             =   105
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Cerrar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmCotizacion.frx":54BA
      End
      Begin VB.Label lblWGestion 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   240
         Index           =   0
         Left            =   50
         TabIndex        =   8
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGestion 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   170
         Width           =   1770
      End
   End
   Begin XtremeSuiteControls.TabControl TabCotizacion 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _Version        =   851968
      _ExtentX        =   10398
      _ExtentY        =   4260
      _StockProps     =   68
      Color           =   8
      ItemCount       =   1
      Item(0).Caption =   "Cotizacion del dia"
      Item(0).ControlCount=   9
      Item(0).Control(0)=   "txtCotizacion(0)"
      Item(0).Control(1)=   "txtCotizacion(1)"
      Item(0).Control(2)=   "txtCotizacion(2)"
      Item(0).Control(3)=   "txtCotizacion(3)"
      Item(0).Control(4)=   "lblAlta(3)"
      Item(0).Control(5)=   "lblAlta(0)"
      Item(0).Control(6)=   "lblAlta(1)"
      Item(0).Control(7)=   "lblAlta(2)"
      Item(0).Control(8)=   "Label1"
      Begin XtremeSuiteControls.FlatEdit txtCotizacion 
         Height          =   315
         Index           =   0
         Left            =   2600
         TabIndex        =   1
         Top             =   600
         Width           =   2750
         _Version        =   851968
         _ExtentX        =   4851
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCotizacion 
         Height          =   315
         Index           =   1
         Left            =   2595
         TabIndex        =   2
         Top             =   960
         Width           =   2745
         _Version        =   851968
         _ExtentX        =   4851
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin XtremeSuiteControls.FlatEdit txtCotizacion 
         Height          =   315
         Index           =   2
         Left            =   2595
         TabIndex        =   3
         Top             =   1320
         Width           =   2745
         _Version        =   851968
         _ExtentX        =   4851
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtCotizacion 
         Height          =   315
         Index           =   3
         Left            =   2595
         TabIndex        =   4
         Top             =   1680
         Width           =   2745
         _Version        =   851968
         _ExtentX        =   4851
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
         Locked          =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   185
         Left            =   240
         TabIndex        =   14
         Top             =   2120
         Width           =   5175
         _Version        =   851968
         _ExtentX        =   9128
         _ExtentY        =   326
         _StockProps     =   79
         Caption         =   "* Cotizacion no oficial extraida desde www.midolar.com.ar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Transparent     =   -1  'True
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Hora Ult. Cotizacion:"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   13
         Top             =   1360
         Width           =   2205
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizacion Dolar Venta:"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Top             =   1000
         Width           =   2205
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cotizacion Dolar Compra:"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   11
         Top             =   640
         Width           =   2205
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Ult. Cotizacion:"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   10
         Top             =   1720
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vAccion As String
Private Declare Function URLDownloadToFile Lib "urlmon" _
   Alias "URLDownloadToFileA" _
  (ByVal pCaller As Long, _
   ByVal szURL As String, _
   ByVal szFileName As String, _
   ByVal dwReserved As Long, _
   ByVal lpfnCB As Long) As Long
Dim ERROR_SUCCESS
Public Function ObtenerPedazoDePaginaWeb(Pagina As String, Cantidad As Long) As String
Dim hOpen As Long, hFile As Long, sIP As String, ret As Long
Dim Longitud As Integer, ax As Integer, valido As Boolean, Scaracter As String
Dim sURL  As String

    sURL = Pagina
    sIP = Space(Cantidad)
    hOpen = InternetOpen("MERCADO", 1, vbNullString, vbNullString, 0)
    hFile = InternetOpenUrl(hOpen, sURL, vbNullString, ByVal 0&, &H80000000, ByVal 0&)
    InternetReadFile hFile, sIP, Cantidad, ret
    InternetCloseHandle hFile
    InternetCloseHandle hOpen
    ObtenerPedazoDePaginaWeb = sIP

End Function
Private Sub Form_Load()
On Error Resume Next

    vAccion = "Nuevo"
    
    Dim sSourceUrl As String

    'sSourceUrl = ""

    Call DownloadFile("http://www.midolar.com.ar/dolar.xml", App.Path & "\dolar.xml")
    
    ActualizarCotizacion

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next
    
    Select Case Index
    
        Case 0
            Grabar
            
            
        Case 1
            Unload Me
    End Select
    
    'frmBusqueda.CargarRegistros

If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub ActualizarCotizacion()
On Error Resume Next

    Dim j As Integer
    With Me
        For j = 0 To .txtCotizacion.Count - 1
            .txtCotizacion(j).Text = EsNulo(CotizacionDolar(Trim(j)))
        Next
    
    End With

If Err Then GrabarLog "ActualizarCotizacion", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Function DownloadFile(ByVal sURL As String, ByVal sLocalFile As String) As Boolean
  DownloadFile = URLDownloadToFile(0, sURL, _
    sLocalFile, 0, 0) = ERROR_SUCCESS
End Function
Private Sub Grabar()
    On Error Resume Next

    If Not ValidarCampos() = True Then
        Exit Sub
    End If
    
    Dim rsCotizacion As New ADODB.Recordset, sqlCotizacion As String
    
    Select Case Me.vAccion

        Case "Nuevo"
            sqlCotizacion = "SELECT * FROM Cotizaciones WHERE 1=2"
        
        Case "sqlCotizacion"
            sqlCotizacion = "SELECT * FROM Cotizaciones WHERE (idCotizacion = '" & Trim(txtCotizacion(0).Tag) & "')"
        
        Case "Duplicar"
            
    End Select
        
    With rsCotizacion
        Call .Open(sqlCotizacion, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .State = 0 Then
        
            Select Case Me.vAccion
            
                Case "Nuevo"
                    .AddNew
                
                Case "Modificar"
                    'No hago nada
                    
                Case "Duplicar"
                    .AddNew
                    '.Fields("Codigo").Value = "" 'Tendria que traer el ultimo codigo
                    '.Fields("codigo_num").Value = Val(txtAlta(0).Text)

            End Select

            .Fields("DolarCompra").Value = PonerPunto(txtCotizacion(0).Text)
            .Fields("DolarVenta").Value = PonerPunto(Left(txtCotizacion(1).Text, 255))
            .Fields("Hora").Value = txtCotizacion(2).Text
            .Fields("Fecha").Value = strfechaMySQL(txtCotizacion(3).Text)
            
            .Update
        
        End If
        
    End With

    sqlCotizacion = ""
    
    If rsCotizacion.State = 1 Then
        rsCotizacion.Close
        Set rsCotizacion = Nothing
    End If
    
    If Err Then
        GrabarLog "Grabar", Err.Number & " " & Err.Description, Me.Name
    Else
        Unload Me
    End If

End Sub
Private Function ValidarCampos() As Boolean
    On Error Resume Next

    Dim i As Integer
    
    ValidarCampos = True
    
    For i = 0 To txtCotizacion.Count - 1
        If Trim(txtCotizacion(1).Text) = "" Then
            MsgBox "Campos obligatorios vacios!", vbExclamation, "Mensaje ..."
            ValidarCampos = Not True
            Exit Function
        End If
    Next
    
    If Me.vAccion = "Nuevo" Then
        If Not TraerDato("Cotizaciones", "Fecha = '" & strfechaMySQL(Trim(txtCotizacion(3).Text)) & "' AND Hora= '" & Trim(txtCotizacion(2).Text) & "'", "idCotizaciones") = "" Then
            MsgBox "Existe un registro con esa Hora/Fecha!", vbExclamation, "Mensaje ..."
            ValidarCampos = Not True
            Exit Function
        End If
    End If
    
    If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub txtCotizacion_KeyPress(Index As Integer, Keyascii As Integer)
On Error Resume Next
    
    If Keyascii = 13 Then
        If txtCotizacion(Index + 1).Visible = True Then
            txtCotizacion(Index + 1).SetFocus
        Else
            PbAcciones(0).SetFocus
        End If
    End If
    
If Err Then GrabarLog "txtCotizacion_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
