VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmListadoSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Saldo Clientes"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   7890
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   615
      Left            =   90
      TabIndex        =   40
      Top             =   5100
      Width           =   7755
      _Version        =   851968
      _ExtentX        =   13679
      _ExtentY        =   1085
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkmd 
         Height          =   285
         Left            =   5730
         TabIndex        =   44
         Top             =   225
         Width           =   1905
         _Version        =   851968
         _ExtentX        =   3360
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Mostrar Dirección"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadioButton1 
         Height          =   375
         Left            =   4470
         TabIndex        =   43
         Top             =   180
         Width           =   945
         _Version        =   851968
         _ExtentX        =   1667
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Normal"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton chkCorto 
         Height          =   345
         Left            =   2370
         TabIndex        =   42
         Top             =   180
         Width           =   1545
         _Version        =   851968
         _ExtentX        =   2725
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Corto"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox chkLrapido 
         Height          =   375
         Left            =   150
         TabIndex        =   41
         Top             =   180
         Width           =   1605
         _Version        =   851968
         _ExtentX        =   2831
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Listado rápido"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   585
      Left            =   60
      TabIndex        =   29
      Top             =   6150
      Width           =   7815
      _Version        =   851968
      _ExtentX        =   13785
      _ExtentY        =   1032
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit vlocalidad 
         Height          =   285
         Left            =   1080
         TabIndex        =   31
         Top             =   180
         Width           =   6660
         _Version        =   851968
         _ExtentX        =   11747
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblLocalidad 
         Caption         =   "Localidad:"
         Height          =   255
         Left            =   150
         TabIndex        =   30
         Top             =   210
         Width           =   1005
      End
   End
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   495
      Left            =   6960
      TabIndex        =   24
      Top             =   6840
      Width           =   915
      _Version        =   851968
      _ExtentX        =   1614
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Prueba"
      UseVisualStyle  =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   30
      TabIndex        =   21
      Top             =   2670
      Width           =   7815
      Begin XtremeSuiteControls.RadioButton RadTodos 
         Height          =   420
         Left            =   4275
         TabIndex        =   45
         Top             =   1800
         Width           =   825
         _Version        =   851968
         _ExtentX        =   1455
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Todos"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2010
         Picture         =   "frmListadoSaldosClientes.frx":0000
         ScaleHeight     =   375
         ScaleWidth      =   435
         TabIndex        =   28
         Top             =   180
         Width           =   435
      End
      Begin MSComCtl2.DTPicker vfecha 
         Height          =   345
         Left            =   2460
         TabIndex        =   22
         Top             =   1800
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         CheckBox        =   -1  'True
         Format          =   203948033
         CurrentDate     =   40821
      End
      Begin XtremeSuiteControls.ComboBox vtipoProveedor 
         Height          =   315
         Left            =   2460
         TabIndex        =   26
         Top             =   240
         Width           =   5205
         _Version        =   851968
         _ExtentX        =   9181
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Text            =   "Proveedor"
      End
      Begin XtremeSuiteControls.ComboBox vtipoProveedor2 
         Height          =   315
         Left            =   2460
         TabIndex        =   34
         Top             =   600
         Width           =   5205
         _Version        =   851968
         _ExtentX        =   9181
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox vtipoProveedor3 
         Height          =   315
         Left            =   2460
         TabIndex        =   36
         Top             =   960
         Width           =   5205
         _Version        =   851968
         _ExtentX        =   9181
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ComboBox vtipoProveedor4 
         Height          =   315
         Left            =   2460
         TabIndex        =   38
         Top             =   1320
         Width           =   5205
         _Version        =   851968
         _ExtentX        =   9181
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.RadioButton RadSoloDoc 
         Height          =   420
         Left            =   5310
         TabIndex        =   46
         Top             =   1800
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1773
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Solo Doc."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RadSoloFact 
         Height          =   420
         Left            =   6435
         TabIndex        =   47
         Top             =   1800
         Width           =   1230
         _Version        =   851968
         _ExtentX        =   2170
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Solo Fact."
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo de persona 4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   39
         Top             =   1350
         Width           =   1710
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo de persona 3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   37
         Top             =   990
         Width           =   1770
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de persona 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   35
         Top             =   630
         Width           =   1800
      End
      Begin VB.Label lblTipoDe 
         Caption         =   "Tipo de persona 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   120
         TabIndex        =   27
         Top             =   270
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Saldo a la fecha:"
         Height          =   225
         Left            =   405
         TabIndex        =   23
         Top             =   1800
         Width           =   1305
      End
   End
   Begin XtremeSuiteControls.GroupBox GBTipoSaldos 
      Height          =   465
      Left            =   60
      TabIndex        =   16
      Top             =   5700
      Width           =   7815
      _Version        =   851968
      _ExtentX        =   13785
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "Mostrar Por Saldo"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.RadioButton RBTipoSaldos 
         Height          =   210
         Index           =   0
         Left            =   2610
         TabIndex        =   17
         Top             =   210
         Width           =   855
         _Version        =   851968
         _ExtentX        =   1508
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Todos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBTipoSaldos 
         Height          =   210
         Index           =   1
         Left            =   3930
         TabIndex        =   18
         Top             =   210
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Deudores"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.RadioButton RBTipoSaldos 
         Height          =   210
         Index           =   2
         Left            =   5250
         TabIndex        =   19
         Top             =   210
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Saldo"
         Appearance      =   6
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBTipoSaldos 
         Height          =   210
         Index           =   3
         Left            =   6390
         TabIndex        =   20
         Top             =   210
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Saldo Acredor"
         Appearance      =   6
      End
   End
   Begin XtremeSuiteControls.GroupBox GBOtros 
      Height          =   735
      Left            =   30
      TabIndex        =   12
      Top             =   1920
      Width           =   7815
      _Version        =   851968
      _ExtentX        =   13785
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Otros Datos"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkConSaldos 
         Height          =   210
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   2775
         _Version        =   851968
         _ExtentX        =   4895
         _ExtentY        =   370
         _StockProps     =   79
         Caption         =   "Imprimir sólo Personas con saldos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.ColorPicker CPSaldo 
         Height          =   345
         Index           =   1
         Left            =   5760
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Saldo Total"
         Appearance      =   3
         SelectedColor   =   49152
      End
      Begin XtremeSuiteControls.ColorPicker CPSaldo 
         Height          =   345
         Index           =   0
         Left            =   3840
         TabIndex        =   15
         Top             =   240
         Visible         =   0   'False
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Saldo "
         Appearance      =   3
         SelectedColor   =   255
         DefaultColor    =   255
      End
   End
   Begin XtremeSuiteControls.GroupBox GBOrden 
      Height          =   735
      Left            =   30
      TabIndex        =   6
      Top             =   1140
      Width           =   7815
      _Version        =   851968
      _ExtentX        =   13785
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Ordenar Listado Por"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkOrden 
         Height          =   255
         Left            =   6030
         TabIndex        =   11
         Top             =   360
         Width           =   1695
         _Version        =   851968
         _ExtentX        =   2990
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Descendente"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBOrden 
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Tag             =   "C.Codigo"
         Top             =   360
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1764
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Codigo"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBOrden 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   8
         Tag             =   "C.Nombre"
         Top             =   360
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1764
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Nombre"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBOrden 
         Height          =   255
         Index           =   2
         Left            =   2220
         TabIndex        =   9
         Tag             =   "Direccion"
         Top             =   360
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1764
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Direccion"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBOrden 
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   10
         Tag             =   "Localidad"
         Top             =   390
         Width           =   1005
         _Version        =   851968
         _ExtentX        =   1764
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Localidad"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBOrden 
         Height          =   255
         Index           =   4
         Left            =   4500
         TabIndex        =   32
         Tag             =   "SSaldo"
         Top             =   390
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Importe"
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GBDatosClientes 
      Height          =   1095
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   7815
      _Version        =   851968
      _ExtentX        =   13785
      _ExtentY        =   1940
      _StockProps     =   79
      Caption         =   "Datos Persona:"
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.CheckBox chkAceptados 
         Height          =   225
         Left            =   5310
         TabIndex        =   33
         Top             =   750
         Width           =   2415
         _Version        =   851968
         _ExtentX        =   4260
         _ExtentY        =   397
         _StockProps     =   79
         Caption         =   "Solo documentos aceptados"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBDatosClientes 
         Height          =   255
         Index           =   0
         Left            =   3900
         TabIndex        =   2
         Tag             =   "Codigo, Nombre y Saldo"
         Top             =   150
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Reducido"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBDatosClientes 
         Height          =   255
         Index           =   1
         Left            =   5340
         TabIndex        =   3
         Tag             =   "Codigo, Nombre, Direccion, Localidad, Telefono, Saldo"
         Top             =   150
         Width           =   1335
         _Version        =   851968
         _ExtentX        =   2355
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Normal"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton RBDatosClientes 
         Height          =   255
         Index           =   2
         Left            =   6780
         TabIndex        =   4
         Tag             =   "Codigo, Nombre, Direccion, Localidad, Telefono, Tipo Iva, Cuit, Total Debe, Total Haber, Saldo"
         Top             =   150
         Width           =   945
         _Version        =   851968
         _ExtentX        =   1667
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Completo"
         Enabled         =   0   'False
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblDatos 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00EDE803&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   7545
      End
      Begin VB.Label lblSaldos 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar los tipos de datos:"
         Height          =   195
         Index           =   2
         Left            =   1260
         TabIndex        =   5
         Top             =   165
         Width           =   2160
      End
   End
   Begin XtremeSuiteControls.PushButton PBAcciones 
      Height          =   495
      Left            =   90
      TabIndex        =   0
      Top             =   6795
      Width           =   6765
      _Version        =   851968
      _ExtentX        =   11933
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "&Imprimir Listado"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmListadoSaldosClientes.frx":04BD
   End
End
Attribute VB_Name = "frmListadoSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim vanomes, vtablaCP, vctacteCP, vid, vfacturaCP, vcobrospagos, vfdetalleCP As String
Public instanciaCP As String

Private Sub Form_Load()
On Error Resume Next

    RBDatosClientes(1).Value = True
    
    init
    
    vfecha.Value = Date

    CentrarFormulario (Me)

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub init()
If instanciaCP = "Proveedores" Then
    vctacteCP = "pcuentascorrientes"
    vid = "Idpcuentascorrientes"
    vfacturaCP = "pfactura"
    vcobrospagos = "pagos"
    vtablaCP = "proveedores"
    vfdetalleCP = "pfdetalle"
    Me.Caption = "Listado de Saldo Proveedor"

Else
    vctacteCP = "cuentascorrientes"
    vid = "id"
    vfacturaCP = "factura"
    vfdetalleCP = "fdetalle"
    vcobrospagos = "cobros"
    vtablaCP = "clientes"
     Me.Caption = "Listado de Saldo Clientes"
End If

Me.vtipoProveedor.Clear
Me.vtipoProveedor.AddItem ("Proveedor")
Me.vtipoProveedor.AddItem ("Eventual")
Me.vtipoProveedor.AddItem ("Cliente")
Me.vtipoProveedor.AddItem ("Asistido")
Me.vtipoProveedor.AddItem ("Rol1")
Me.vtipoProveedor.AddItem ("Rol2")
Me.vtipoProveedor.AddItem ("Rol3")
Me.vtipoProveedor.AddItem ("Rol4")
Me.vtipoProveedor.AddItem ("Creditos")

'
'Me.vtipoProveedor1.Clear
'Me.vtipoProveedor1.AddItem ("Proveedor")
'Me.vtipoProveedor1.AddItem ("Eventual")
'Me.vtipoProveedor1.AddItem ("Cliente")
'Me.vtipoProveedor1.AddItem ("Asistido")
'Me.vtipoProveedor1.AddItem ("Rol1")
'Me.vtipoProveedor1.AddItem ("Rol2")
'Me.vtipoProveedor1.AddItem ("Rol3")
'Me.vtipoProveedor1.AddItem ("Rol4")
'
'
'Me.vtipoProveedor2.Clear
'Me.vtipoProveedor2.AddItem ("Proveedor")
'Me.vtipoProveedor2.AddItem ("Eventual")
'Me.vtipoProveedor2.AddItem ("Cliente")
'Me.vtipoProveedor2.AddItem ("Asistido")
'Me.vtipoProveedor2.AddItem ("Rol1")
'Me.vtipoProveedor2.AddItem ("Rol2")
'Me.vtipoProveedor2.AddItem ("Rol3")
'Me.vtipoProveedor2.AddItem ("Rol4")
'
'
'Me.vtipoProveedor3.Clear
'Me.vtipoProveedor3.AddItem ("Proveedor")
'Me.vtipoProveedor3.AddItem ("Eventual")
'Me.vtipoProveedor3.AddItem ("Cliente")
'Me.vtipoProveedor3.AddItem ("Asistido")
'Me.vtipoProveedor3.AddItem ("Rol1")
'Me.vtipoProveedor3.AddItem ("Rol2")
'Me.vtipoProveedor3.AddItem ("Rol3")
'Me.vtipoProveedor3.AddItem ("Rol4")


End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub


Function armarTipoProveedor() As String
On Error Resume Next



If Err Then
    armarTipoProveedor = ""
End If
End Function

Private Sub PbAcciones_Click()
On Error Resume Next

    Dim vsql As String, vorden, vtipo, vsqlFecha, vcondi, vtipoProveedor    As String
    
    vsqlFecha = ""
    vtipo = " "
          
    'vtipoProveedor = armarTipoProveedor()
 
    If vfecha.Value Then
        vsqlFecha = " where ccc.fecha <='" + strfechaMySQL(vfecha) + "'"
        vtipo = "Saldo a la fecha: " + Str(vfecha)
    Else
        vtipo = "Saldo final"
    End If
    

vcondi = ""

If Not vlocalidad.Text = "" Then
    vcondi = " and (c.localidad like '%" + vlocalidad.Text + "%') "
End If


Dim vcorto As Integer

    If Me.chkCorto.Value Then
        vcorto = 1
    Else
        vcorto = 0
    End If
    
    Dim via As String
    
    via = ""
    
    If Me.RadSoloDoc.Value = True Then via = "Documento"
    
    If Me.Radsolofact.Value = True Then via = "Fact"
    
    
    
    
    If chkAceptados.Value Then
        vsql = FSQLSaldos(vsqlFecha, vctacteCP, Me.vtipoProveedor, vcondi, "Aceptados", , via)
    Else
        
    
        vsql = FSQLSaldos(vsqlFecha, vctacteCP, Me.vtipoProveedor, vcondi, "", vcorto, via)
    End If
    
    
    'vsql = "SELECT C.Codigo, C.Nombre, Direccion, Localidad, Telefono, C.idTipoIva AS TipoIva, Cuit, Sum(CCC.Debito) AS TotalDebito, Sum(CCC.Credito) AS TotalCredito, Sum(CCC.Debito)-Sum(CCC.Credito) AS SSaldo FROM " + c + " C LEFT JOIN " + vctacteCP + " CCC ON C.Codigo = CCC.Codigo " + vsqlFecha + " GROUP BY C.Codigo"
  
    'vSQL = "SELECT C.Codigo, C.Nombre, Direccion, Localidad, Telefono, C.idTipoIva AS TipoIva, Cuit, Sum(CCC.Debito) AS TotalDebito, Sum(CCC.Credito) AS TotalCredito, Sum(CCC.Debito)-Sum(CCC.Credito) AS Saldo FROM Clientes C LEFT JOIN  CCC ON C.Codigo = CCC.Codigo GROUP BY C.Codigo"
    
    If chkConSaldos.Value = xtpChecked Then
        vsql = Replace(vsql, "LEFT JOIN", "INNER JOIN")
    End If
    
    With Mantenimiento2.rsSaldosClientes
        If .State = 1 Then .Close
        
        .Source = vsql & " ORDER BY " & VerOrden
        .Filter = VerTipoSaldos
        
        If .State = 0 Then .Open
        .Close
        .Open
        
         .Filter = VerTipoSaldos
        '.Filter = "Saldo >0.1"
        
        
        ' .Source = vsql & " ORDER BY " & VerOrden
        
        
        If .RecordCount = 0 Then
        
            'Call PbAcciones_Click
            MsgBox "No hay datos para mostrar", vbInformation
        
            Exit Sub
        End If
    End With
    
    
    If UCase(LeerXml("Puesto")) = "PONS" Or Me.chkLrapido.Value = xtpChecked Then
    
    
    With drClientesSaldosPons
    
         ' Set .DataSource = Mantenimiento2.rsSaldosClientes
        .Sections("TituloEmpresa").Controls("Etipo").Caption = vtipo
        
        If Not Me.vlocalidad.Text = "" Then
            .Sections("TituloEmpresa").Controls("econdi").Caption = "Localidad: " + Me.vlocalidad.Text
        End If
        
        '.Sections("Detalle").Controls("txtSaldo").ForeColor = CPSaldo(0).SelectedColor
        .Sections("PieInforme").Controls("fnSaldoTotal").ForeColor = CPSaldo(1).SelectedColor
        
        .Show
        
    End With

    Else
    
    With drClientesSaldos
            .Sections("TituloEmpresa").Controls("Etipo").Caption = vtipo
        
        If Not Me.vlocalidad.Text = "" Then
            .Sections("TituloEmpresa").Controls("econdi").Caption = "Localidad: " + Me.vlocalidad.Text
        End If
        
        .Sections("Detalle").Controls("txtSaldo").ForeColor = CPSaldo(0).SelectedColor
        .Sections("PieInforme").Controls("fnSaldoTotal").ForeColor = CPSaldo(1).SelectedColor
        
        .Show
        
    End With

    End If

If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub PushButton1_Click()
  
If LeerXml("Puesto") = "PONS" Then
    drClientesSaldosPons.Show
Else
    drClientesSaldos.Show
End If
  
    End Sub

Private Sub RBDatosClientes_Click(Index As Integer)
On Error Resume Next

    lblDatos.Caption = "Se Muestra : " & RBDatosClientes(Index).Tag

If Err Then GrabarLog "RBDatosClientes_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function VerOrden()
On Error Resume Next

    Dim vTipoOrden As String, i As Integer

    For i = 0 To rbOrden.Count - 1
        If rbOrden(i).Value = True Then
            vTipoOrden = rbOrden(i).Tag
            Exit For
        End If
    Next
    
    If chkOrden.Value = xtpChecked Then
        vTipoOrden = vTipoOrden & " DESC"
    Else
        vTipoOrden = vTipoOrden & " ASC"
    End If
    
    VerOrden = vTipoOrden
    
If Err Then GrabarLog "VerOrden", Err.Number & " " & Err.Description, Me.Name
End Function
Private Function VerTipoSaldos() As String
On Error Resume Next

    Dim i As Integer

    For i = 0 To RBTipoSaldos.Count - 1
    
        If RBTipoSaldos(i).Value = True Then
        
            Select Case RBTipoSaldos(i).Caption
            
                Case "Todos"
                    VerTipoSaldos = ""
                    Exit For
                    
                Case "Deudores"
                    
                   'If Me.instanciaCP = "Proveedores" Then
                    'VerTipoSaldos = "SSaldo < -0.1"
                   'Else
                    VerTipoSaldos = "SSaldo > 0.1"
                   'End If
                    
                    Exit For
                    
                Case "Saldo"
                    VerTipoSaldos = "SSaldo > 0.1 or SSaldo < -0.1"
                    Exit For
            
                Case Else
                    
                   'If Me.instanciaCP = "Proveedores" Then
             '       VerTipoSaldos = "SSaldo > 0.1"
             '      Else
                    VerTipoSaldos = "SSaldo < -0.1"
                  ' End If
                    
                   Exit For
            
            End Select
        
        End If
    
    Next
    
If Err Then GrabarLog "VerTipoSaldos", Err.Number & " " & Err.Description, Me.Name
End Function
