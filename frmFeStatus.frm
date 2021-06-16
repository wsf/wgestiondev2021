VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{757B5B41-998B-41F8-95D8-B90E12A1D40B}#222.0#0"; "WSAFIPFEOCX.ocx"
Begin VB.Form frmFeStatus 
   Caption         =   "FE Status"
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   9570
   StartUpPosition =   3  'Windows Default
   Begin WSAFIPFEOCX.WSAFIPFEx vfe2 
      Left            =   6975
      Top             =   1440
      _ExtentX        =   2725
      _ExtentY        =   1376
   End
   Begin VB.Frame Frame1 
      Height          =   915
      Left            =   90
      TabIndex        =   11
      Top             =   930
      Width           =   5535
      Begin VB.TextBox vcuit 
         Height          =   345
         Left            =   3150
         TabIndex        =   13
         Top             =   150
         Width           =   2295
      End
      Begin VB.TextBox vempresa2 
         Height          =   345
         Left            =   3150
         TabIndex        =   12
         Top             =   510
         Width           =   2295
      End
      Begin VB.Label lblCuit 
         Caption         =   "Cuit"
         Height          =   225
         Left            =   180
         TabIndex        =   15
         Top             =   180
         Width           =   1905
      End
      Begin VB.Label lblEmpresa 
         Caption         =   "Empresa:"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   510
         Width           =   2025
      End
   End
   Begin VB.TextBox vnroempresa 
      Height          =   285
      Left            =   3660
      TabIndex        =   9
      Text            =   "0"
      Top             =   510
      Width           =   945
   End
   Begin VB.TextBox vnro2 
      Height          =   285
      Left            =   8400
      TabIndex        =   5
      Text            =   "vnro"
      Top             =   120
      Width           =   1125
   End
   Begin VB.TextBox vsucursal 
      Height          =   285
      Left            =   6090
      TabIndex        =   4
      Text            =   "1001"
      Top             =   120
      Width           =   945
   End
   Begin VB.TextBox vtipo 
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Text            =   "Vtipo"
      Top             =   120
      Width           =   945
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   405
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   795
      _Version        =   851968
      _ExtentX        =   1402
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Consultar"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.ListBox logStatus 
      Height          =   3990
      Left            =   0
      TabIndex        =   0
      Top             =   2610
      Width           =   9525
      _Version        =   851968
      _ExtentX        =   16801
      _ExtentY        =   7038
      _StockProps     =   77
      ForeColor       =   15591427
      BackColor       =   4210752
      BackColor       =   4210752
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin WSAFIPFEOCX.WSAFIPFEx vfe 
      Left            =   30
      Top             =   5070
      _ExtentX        =   2196
      _ExtentY        =   1879
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   405
      Left            =   930
      TabIndex        =   2
      Top             =   30
      Width           =   585
      _Version        =   851968
      _ExtentX        =   1032
      _ExtentY        =   714
      _StockProps     =   79
      Caption         =   "Volver"
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton PusConsultarTodas 
      Height          =   525
      Left            =   135
      TabIndex        =   16
      Top             =   2070
      Width           =   3825
      _Version        =   851968
      _ExtentX        =   6747
      _ExtentY        =   926
      _StockProps     =   79
      Caption         =   "Consultar todas las Empresas"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Label Label4 
      Caption         =   "Nro Empresa:"
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   540
      Width           =   1125
   End
   Begin VB.Label Label3 
      Caption         =   "Nro. Compr."
      Height          =   255
      Left            =   7200
      TabIndex        =   8
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label Label2 
      Caption         =   "Sucursal: "
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   150
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo Doc AFIP:"
      Height          =   255
      Left            =   2220
      TabIndex        =   6
      Top             =   150
      Width           =   1125
   End
End
Attribute VB_Name = "frmFeStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vnroFactA, vnroFactB As Long

Private Sub PusConsultarTodas_Click()

Dim j, i As Integer

j = InputBox("Ingrese la cantidad de empresas que tiene registrada", "WSF")

If j = 0 Then j = 17


Me.logStatus.Clear

For i = 1 To j
    Call getStatusAfipEmpresas(Str(i))
Next

End Sub

Private Sub PushButton1_Click()
Call getStatusAfip
End Sub


Public Sub getStatusAfip(Optional ByRef vnfa, Optional ByRef vnfb)
On Error Resume Next

Dim vcantiIVA, vultimoNroComprobante, vPtoVta2 As Integer
Dim bResultado As Boolean
Dim cIdentificador, vsql, vultimoMensajeError As String
Dim v As Variant
Dim vDocAfip As Variant
Dim vmen As Variant

Screen.MousePointer = vbHourglass

' Documentación en: https://sites.google.com/site/facturaelectronicax/documentacion-wsfev1/wsfev1/wsfev1-metodos
v = Test


MsgBox "Comenenzando ..."

Dim vvcuit As String
Dim vvcertificado As String
Dim vvlicencia, vmodofiscal As String

vPtoVta2 = traerDatos2("select * from configuracion", "SucursalDocVenta", PathDBConfig)

vmodofiscal = Trim(LeerXml("modoFiscal"))

If (Val(vempresa2.Text) > 0) And (Not vcuit = "") Then
            vvcuit = vcuit.Text
            vvcertificado = "Empresa" + Trim(Str(Val(vempresa2.Text))) + ".pfx"
            vvlicencia = "Licencia" + Trim(Str(Val(vempresa2.Text))) + ".lic"
            vPtoVta2 = Me.vsucursal.Text
            vmodofiscal = 1
Else
            vvcuit = Trim(LeerXml("vcuit"))
            vvcertificado = Trim(LeerXml("vcertificado"))
            vvlicencia = Trim(LeerXml("LicenciaWSAFIP"))
End If

MsgBox "a) vcuit, vvcert, vvlic " + vcuit + " " + vvcertificado + " " + vvlicencia

  
'If Trim(LeerXml("modoFiscal")) = "1" Then
'    bResultado = vfe2.iniciar(1, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), App.Path + "\" + Trim(LeerXml("LicenciaWSAFIP")))
' Else
'     bResultado = vfe2.iniciar(0, Trim(LeerXml("vcuit")), App.Path + "\" + Trim(LeerXml("vcertificado")), "")
'End If


vvcuit = "30707384316"

MsgBox "vvcuit : " + vvcuit

If vmodofiscal = "1" Then
    bResultado = vfe2.iniciar(1, vvcuit, App.Path + "\" + vvcertificado, App.Path + "\" + vvlicencia)
    MsgBox vvcuit
 Else
    bResultado = vfe2.iniciar(0, vvcuit, App.Path + "\" + vvcertificado, "")
End If

MsgBox "b) bresultado 1 " + Str(bResultado)



'vsql = "select SucursalDocVenta as c from configuracion limit 1"
'If vPtoVta2 = 0 Then vPtoVta2 = 2
bResultado = vfe2.f1ObtenerTicketAcceso()


MsgBox "c) bresultado 2 : " + Str(bResultado)

vultimoMensajeError = Trim(vfe2.UltimoMensajeError)


MsgBox "d) Ulltimo mensaje error: " + vultimoMensajeError + Chr(13) + _
"bResultado : " + Str(bResultado)



Debug.Print vultimoMensajeError
vnroFactA = vfe2.f1CompUltimoAutorizado(vPtoVta2, 1)
vnroFactB = vfe2.f1CompUltimoAutorizado(vPtoVta2, 6)



MsgBox "e) ------------ ahora salen los nros de facturas --------------"
MsgBox vnroFactA
MsgBox vnroFactB
MsgBox "e) ---------------------------------------"

vfe2.ArchivoXMLRecibido = App.Path + "\Log\recibido.xml"
vfe2.ArchivoXMLEnviado = App.Path + "\Log\enviado.xml"


If Val(vtipo) > 0 Then


    MsgBox "f) " + vPtoVta2 + "  " + Me.vtipo + " " + Me.vnro2
    
    vDocAfip = vfe2.F1CompConsultar(vPtoVta2, Val(Me.vtipo), Val(Me.vnro2))
    vmen = vfe2.F1DetalleCAEA
    vmen = vmen + "  ---- " + vfe2.f1RespuestaCAEA
    vmen = vmen + vfe2.F1RespuestaDetalleCae
    
    
    vfe2.ArchivoXMLRecibido = App.Path + "\Log\recibido.xml"
    vfe2.ArchivoXMLEnviado = App.Path + "\Log\enviado.xml"
    
    MsgBox "12" + vmen
     
     
    
End If


MsgBox "2"


vnfa = vnroFactA
vnfb = vnroFactB

Me.logStatus.Clear

Me.logStatus.AddItem "Ultimo mensaje WS AFIP: " + vultimoMensajeError + Chr(13)
Me.logStatus.AddItem "Ultima Factura A: " + Str(vnroFactA)
Me.logStatus.AddItem "Ultima Factura B: " + Str(vnroFactB)

If Not vDocAfip = "" Then
    Me.logStatus.AddItem "__________________________________________"
    Me.logStatus.AddItem ""
    Me.logStatus.AddItem "Documento AFIP :  " + Trim(vDocAfip)
    Me.logStatus.AddItem "Documento AFIP :  " + vmen
    Me.logStatus.AddItem "__________________________________________"
End If

Screen.MousePointer = vbDefault

If Err < 0 Then
    MsgBox " Err.Descripcion: " + Err.Description
    Exit Sub
End If
End Sub






Public Sub getStatusAfipEmpresas(ByVal vnroempresa As String, Optional ByRef vnfa, Optional ByRef vnfb)
On Error Resume Next

Dim vcantiIVA, vultimoNroComprobante, vPtoVta2 As Integer
Dim bResultado As Boolean
Dim cIdentificador, vsql, vultimoMensajeError As String
Dim v As Variant
Dim vDocAfip As Variant
Dim vmen As Variant

Screen.MousePointer = vbHourglass

' Documentación en: https://sites.google.com/site/facturaelectronicax/documentacion-wsfev1/wsfev1/wsfev1-metodos
v = Test

Dim vvcuit As String
Dim vvcertificado As String
Dim vvlicencia, vmodofiscal As String

vPtoVta2 = 1001

vmodofiscal = Trim(LeerXml("modoFiscal"))



            vvcuit = vcuit.Text
            vvcertificado = "Empresa" + Trim(Str(Val(vempresa2.Text))) + ".pfx"
            vvlicencia = "Licencia" + Trim(Str(Val(vempresa2.Text))) + ".lic"
            vPtoVta2 = Me.vsucursal.Text
            vmodofiscal = 1
vnroempresa = Trim(vnroempresa)

        bResultado = vfe2.iniciar(1, Trim((getCuitFE(vnroempresa))), App.Path + "\" + Trim(getCertificadoFE(vnroempresa)), App.Path + "\" + Trim(getLicenciaFE(vnroempresa)))


Debug.Print "Cuit: " + Trim(Str(getCuitFE(vnroempresa)))
Debug.Print "Certificado: " + Trim((getCertificadoFE(vnroempresa)))
Debug.Print "Licencia : " + Trim(getLicenciaFE(vnroempresa))




bResultado = vfe2.f1ObtenerTicketAcceso()

vultimoMensajeError = Trim(vfe2.UltimoMensajeError)

Debug.Print vultimoMensajeError

vnroFactA = vfe2.f1CompUltimoAutorizado(vPtoVta2, 1)
vnroFactB = vfe2.f1CompUltimoAutorizado(vPtoVta2, 6)


vtipo = 1

If Val(vtipo) > 0 Then
    vDocAfip = vfe2.F1CompConsultar(vPtoVta2, Val(Me.vtipo), Val(Me.vnro2))
    
    vmen = vfe2.F1DetalleCAEA
    vmen = vmen + "  " + vfe2.f1RespuestaCAEA
    vmen = vmen + vfe2.F1RespuestaDetalleCae
    vmen = vmen + Chr(13)
    
    
    
    vfe2.ArchivoXMLRecibido = App.Path + "\Log\recibido-" + Str(vnroempresa) + ".xml"
    vfe2.ArchivoXMLEnviado = App.Path + "\Log\enviado-" + Str(vnroempresa) + ".xml"
     
     
    
End If

vnfa = vnroFactA
vnfb = vnroFactB


Me.logStatus.AddItem "Ultimo mensaje WS AFIP: " + vultimoMensajeError + Chr(13)
Me.logStatus.AddItem "Ultima Factura A: " + Str(vnroFactA)
Me.logStatus.AddItem "Ultima Factura B: " + Str(vnroFactB)
    
    Me.logStatus.AddItem "____________  " + Str(vnroempresa) + "  ______________________________"
    Me.logStatus.AddItem ""
    Me.logStatus.AddItem "Documento AFIP :  " + Trim(vDocAfip)
    Me.logStatus.AddItem "Documento AFIP :  " + vmen
    Me.logStatus.AddItem "__________________________________________"

    Me.logStatus.AddItem ""

Screen.MousePointer = vbDefault

If Err < 0 Then
    'getNroCompAfip = 0
    Exit Sub
End If
End Sub


Private Sub vnrocuit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

End Sub
