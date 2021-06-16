VERSION 5.00
Object = "{39ABE45D-F077-4D34-A361-6906C77D67F7}#1.0#0"; "Fiscal150423.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.CommandBars.v13.0.0.Demo.ocx"
Object = "{AFD24A52-2823-4FBD-B75D-C282C11E1D98}#1.0#0"; "IFEpson.ocx"
Object = "{FF19AA0C-2968-41B8-A906-E80997A9C394}#208.0#0"; "WSAFIPFEOCX.ocx"
Object = "{706C3604-A82B-4400-9EE4-3433F1D8DB08}#1.8#0"; "EpsonFPHostControlX.ocx"
Begin VB.MDIForm frmPrincipal 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000F&
   ClientHeight    =   8490
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10620
   Icon            =   "frmPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Picture         =   "frmPrincipal.frx":6852
   WhatsThisHelp   =   -1  'True
   WindowState     =   2  'Maximized
   Begin EpsonFPHostControlX.EpsonFPHostControl FiscalEpson2 
      Left            =   1200
      OleObjectBlob   =   "frmPrincipal.frx":136D2
      Top             =   4440
   End
   Begin WSAFIPFEOCX.WSAFIPFEx fe 
      Left            =   180
      Top             =   225
      _ExtentX        =   1720
      _ExtentY        =   1296
   End
   Begin EPSON_Impresora_Fiscal.PrinterFiscal FiscalEpson 
      Left            =   1680
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin MSComctlLib.StatusBar barra 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   8175
      Width           =   10620
      _ExtentX        =   18733
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin FiscalPrinterLibCtl.HASAR FiscalHasar 
      Left            =   3120
      OleObjectBlob   =   "frmPrincipal.frx":1376C
      Top             =   3480
   End
   Begin XtremeCommandBars.CommandBars Botonera 
      Left            =   1530
      Top             =   120
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      ScaleMode       =   1
      VisualTheme     =   7
   End
   Begin XtremeCommandBars.ImageManager IMIconos 
      Left            =   2400
      Top             =   240
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmPrincipal.frx":13790
   End
   Begin VB.Menu cajaIngresos 
      Caption         =   "Caja-Ingresos"
      WindowList      =   -1  'True
      Begin VB.Menu IngresoDineroCaja 
         Caption         =   "Ingreso Dinero a Caja"
      End
   End
   Begin VB.Menu cajaEgresos 
      Caption         =   "Caja-Egresos"
      Begin VB.Menu pagoProveedor 
         Caption         =   "Pago a Proveedor"
      End
      Begin VB.Menu OtrosEgresos 
         Caption         =   "Egreso Dinero de Cajas"
      End
   End
   Begin VB.Menu cajalistados 
      Caption         =   "Caja-Listados"
      Begin VB.Menu movimientosPorCtas 
         Caption         =   "Movimienos por Cuentas"
      End
      Begin VB.Menu balance 
         Caption         =   "Balance"
      End
      Begin VB.Menu presupuesto 
         Caption         =   "Presupuesto"
      End
      Begin VB.Menu presupuestadovsgastos 
         Caption         =   "Presupuestasdo vs Gastos"
      End
      Begin VB.Menu l1 
         Caption         =   "-"
      End
      Begin VB.Menu cdcaja 
         Caption         =   "Cierre de Caja"
      End
      Begin VB.Menu saldocb 
         Caption         =   "Saldos de Caja  y Bancos"
      End
      Begin VB.Menu mcb 
         Caption         =   "Movimientos de Caja y Bancos"
      End
      Begin VB.Menu l2 
         Caption         =   "-"
      End
      Begin VB.Menu moviCaja 
         Caption         =   "Movientos Caja"
      End
      Begin VB.Menu moviBancos 
         Caption         =   "Movimientos de Bancos"
      End
      Begin VB.Menu lin21 
         Caption         =   "-"
      End
      Begin VB.Menu ao 
         Caption         =   "Adelantos otorgados"
      End
      Begin VB.Menu imp2 
         Caption         =   "-"
      End
      Begin VB.Menu impPorPersonas 
         Caption         =   "Importes por Personas"
      End
   End
   Begin VB.Menu cajaEventuaes 
      Caption         =   "Caja-Eventuales"
      Begin VB.Menu mantEventuales 
         Caption         =   "Mantenimiento"
      End
      Begin VB.Menu ConsultasEventuales 
         Caption         =   "Consultas"
      End
   End
   Begin VB.Menu rendiciones 
      Caption         =   "Rendiciones"
      Begin VB.Menu webcam 
         Caption         =   "WebCam"
      End
      Begin VB.Menu webcamlinea 
         Caption         =   "-"
      End
      Begin VB.Menu redifae 
         Caption         =   "F.A.E."
      End
   End
   Begin VB.Menu cajaConfiguracion 
      Caption         =   "Caja-Configuración"
      Begin VB.Menu presupuestos 
         Caption         =   "Presupuestos"
      End
      Begin VB.Menu cuentascontables 
         Caption         =   "Cuentas Contables"
      End
      Begin VB.Menu cajas 
         Caption         =   "Cajas"
      End
      Begin VB.Menu bancos 
         Caption         =   "Bancos"
      End
      Begin VB.Menu fconceptos 
         Caption         =   "Conceptos"
      End
      Begin VB.Menu l44 
         Caption         =   "-"
      End
      Begin VB.Menu empleados 
         Caption         =   "Empleados"
      End
      Begin VB.Menu lborra 
         Caption         =   "-"
      End
      Begin VB.Menu borrardatos 
         Caption         =   "Borrar Datos"
      End
   End
   Begin VB.Menu wsfe 
      Caption         =   "WS FE"
   End
   Begin VB.Menu control 
      Caption         =   "Control"
      Begin VB.Menu agenda 
         Caption         =   "Agenda"
      End
      Begin VB.Menu agendalinea 
         Caption         =   "-"
      End
      Begin VB.Menu errores 
         Caption         =   "Errores"
      End
      Begin VB.Menu ber 
         Caption         =   "-"
      End
      Begin VB.Menu tablero 
         Caption         =   "Tablero"
      End
      Begin VB.Menu t11 
         Caption         =   "-"
      End
      Begin VB.Menu tgi 
         Caption         =   "Transacciones. Gestión Integral"
      End
      Begin VB.Menu lcz1 
         Caption         =   "-"
      End
      Begin VB.Menu cierrez 
         Caption         =   "Cierre Z"
      End
      Begin VB.Menu l12 
         Caption         =   "-"
      End
      Begin VB.Menu prueba1 
         Caption         =   "Prueba No fiscal"
      End
      Begin VB.Menu lpfiscal 
         Caption         =   "-"
      End
      Begin VB.Menu pruebafiscal 
         Caption         =   "Prueba Fiscal"
      End
      Begin VB.Menu lsf1 
         Caption         =   "-"
      End
      Begin VB.Menu estadoFiscal 
         Caption         =   "Estado Fiscal"
      End
      Begin VB.Menu pruebaepson 
         Caption         =   "Prueba Epson"
      End
   End
   Begin VB.Menu listados1 
      Caption         =   "Listados"
      Begin VB.Menu lingresos 
         Caption         =   "Ingreso"
      End
   End
   Begin VB.Menu gastosfijos 
      Caption         =   "Gastos Fijos"
      Begin VB.Menu agendaGastos 
         Caption         =   "Agenda"
      End
      Begin VB.Menu lgf6 
         Caption         =   "-"
      End
      Begin VB.Menu lgf 
         Caption         =   "Listados de Gastos Fijos"
      End
   End
   Begin VB.Menu servicios 
      Caption         =   "Servicios-Comercios"
      Begin VB.Menu estadocuentasrurales 
         Caption         =   "Estado de Cuentas Servicios Rural - Urbanos"
      End
      Begin VB.Menu servcomercio 
         Caption         =   "-"
      End
      Begin VB.Menu comercios 
         Caption         =   "Comercios"
      End
   End
   Begin VB.Menu ayuda 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public WithEvents TabToolBar As TabToolBar
Attribute TabToolBar.VB_VarHelpID = -1
Dim vMenuFavoritos As Long
Dim sCmd, sCmdExt As String
Dim bAnswer As Boolean

Private Sub DockingPane1_AttachPane(ByVal Item As XtremeDockingPane.IPane)

End Sub

Private Sub cc_Click()
Call controlarFacturaDetalles
End Sub

Private Sub documentos_Click()
'dra.Show
End Sub

Private Sub fejemplo_Click()
frmEjemplo.Show
End Sub

Private Sub agenda_Click()
frmAgenda.Show
End Sub

Private Sub agendaGastos_Click()
    Call Shell("ayuda.bat")
 
    
    'Call Shell("https://www.google.com/calendar/render#main_7%7Cmonth")
End Sub

Private Sub ao_Click()

frmCtaCteC.Tag = "Proveedores"
frmCtaCteC.Show
frmCtaCteC.WindowState = vmaximizar
frmCtaCteC.init


frmCtaCteC.vdtipo.Tag = "ADT"
'frmCtaCteC.vctipo.Text = "ADT"
frmCtaCteC.vdtipo.Text = "Adelanto"


End Sub

Private Sub ayuda_Click()
frmAyuda.v = "https://www.google.com/calendar/render#main_7%7Cmonth"
frmAyuda.Show

End Sub

Private Sub balance_Click()
frmBalance.Show
End Sub

Private Sub bancos_Click()
frmBancos.Show
End Sub

Private Sub borrardatos_Click()
    frmBorrarBases.Show
End Sub

Private Sub cajas_Click()
frmBancos.Show
End Sub

Private Sub cdcaja_Click()
frmBancoCajaDetalle.Show
frmBancoCajaDetalle.vtipolistado.Text = "Agrupado por Cajas-Bancos"
frmBancoCajaDetalle.tabbc.SelectedItem = 2
End Sub

Public Sub cierrez_Click()
On Error GoTo impresora_apag

If MsgBox("Confirma querer realizar el cierre Z", vbYesNo) = vbNo Then Exit Sub

Procesar:

   If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket Hasar") Then frmPrincipal.FiscalHasar.ReporteZ
   
   
   
   If UCase(LeerXml("Impresora")) = UCase("Fiscal Ticket Epson") Then
             cierrezAle
            'frmPrincipal.FiscalEpson2.ReporteZ
   End If
    
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
    
End Sub

Private Sub cierrezAle()
Dim cCmd, sCmExt, vmensaje As String
Dim bAnswer As Boolean
bAnswer = True

With frmPrincipal.FiscalEpson2
        
        sCmd = Chr$(&H8) + Chr$(&H1)
        
        If bAnswer Then bAnswer = .AddDataField(sCmd)
        sCmdExt = Chr$(&HC) + Chr$(&H0)
        
        If bAnswer Then bAnswer = .AddDataField(sCmdExt)
        
        If bAnswer Then bAnswer = .SendCommand
        Call FPDelay
        
        If .ReturnCode <> 0 Then ShowMsg
        
End With
End Sub



Private Sub comercios_Click()
frmDeudasServicios2.Show
frmDeudasServicios2.RdComercio = True
End Sub

Private Sub cuentascontables_Click()
frmCuentas.Show
End Sub

Private Sub empleados_Click()
frmEmpleados.Show
End Sub

Private Sub errores_Click()
    frmAlarmas.Show
End Sub

Private Sub estadocuentasrurales_Click()
    frmDeudasServicios.Show
End Sub

Private Sub estadoFiscal_Click()
Dim vdatos, vmensaje  As String

vdatos = UCase(LeerXml("Impresora"))

With frmPrincipal.FiscalEpson2

    If vdatos = UCase("Fiscal Ticket epson") Then
                sCmd = Chr$(&H0) + Chr$(&H1)
                bAnswer = .AddDataField(sCmd)
                sCmdExt = Chr$(&H0) + Chr$(&H0)
                If bAnswer Then bAnswer = .AddDataField(sCmdExt)
                If bAnswer Then bAnswer = .SendCommand
                Call FPDelay
    End If

    vmensaje = "Impresora: " + Format(Hex(.PrinterStatus), "0000") + Chr(13) + _
                "Fiscal: " + Format(Hex(.FiscalStatus), "0000") + Chr(13) + _
                "RC: " + Format(Hex(.ReturnCode), "0000") + Chr(13)

    MsgBox vmensaje
End With

End Sub

Private Sub fconceptos_Click()
frmConceptos.Show
End Sub

Private Sub FiscalHasar_ErrorFiscal(ByVal flags As Long)
Debug.Print FiscalHasar.DescripcionStatusFiscal(flags)
End Sub
Private Sub FiscalHasar_EventoFiscal(ByVal flags As Long)
    Debug.Print CStr(flags)
    On Error Resume Next
    Debug.Print FiscalHasar.DescripcionStatusFiscal(flags)
End Sub
Private Sub FiscalHasar_EventoImpresora(ByVal flags As Long)

    Debug.Print FiscalHasar.DescripcionStatusImpresor(flags)
    
    Select Case flags
        Case P_JOURNAL_PAPER_LOW, P_RECEIPT_PAPER_LOW:
            Debug.Print "Falta papel"
        Case P_OFFLINE:
            Debug.Print "Impresor fuera de línea"
        Case P_PRINTER_ERROR:
            Debug.Print "Error mecánico de impresor"
        Case Else:
            Debug.Print "Otro bit de impresora"
    End Select

End Sub

Private Sub impPorPersonas_Click()
frmBancoCajaDetalle.cboAgrupado = "Personas"
frmBancoCajaDetalle.Show
End Sub

Private Sub IngresoDineroCaja_Click()

If LeerXml("Puesto") = "Comuna" Or LeerXml("Puesto") = "Caja" Or LeerXml("Puesto") = "ASOCIAL" Then
        Call frmIngresosEgresos.initEgreso

Else
        MsgBox ("Este puesto no tiene autorizaciòn." + Chr(13) + "Consulte al Servicio Tècnico")
        
End If
   
End Sub

Private Sub lgf_Click()
frmBuscarFactura.Show
Call frmBuscarFactura.PushButton14_Click
frmBuscarFactura.chkFacturaX.Value = 1
frmBuscarFactura.chkFechaTodas.Value = False

frmBuscarFactura.dtpFecha(0).Value = Date - (Day(Date) - 1)
frmBuscarFactura.dtpFecha(1).Value = DiasDelMes(frmBuscarFactura.dtpFecha(0).Value) & "/" & AjustarMes(Month(frmBuscarFactura.dtpFecha(0).Value)) & "/" & Year(frmBuscarFactura.dtpFecha(0).Value)
Call frmBuscarFactura.cmdFiltrar_Click

End Sub

Private Sub lingresos_Click()
On Error Resume Next
Dim vsql As String

vsql = "select * from factura t limit 10"
Call rsToExcel(vsql)

If Err Then Exit Sub
End Sub

Private Sub mantEventuales_Click()
frmTrabEventuales.Show
End Sub

Private Sub mcb_Click()
frmBancoCajaDetalle.vtipolistado = "Todos"
frmBancoCajaDetalle.Show
End Sub

Private Sub MDIForm_DblClick()
 On Error Resume Next
    
    Dim vPassword As String
    
    vPassword = InputBox("Migrar", "Mensaje ...")
    
    If vPassword = "WSF" Then
        frmMigracion.Show
    End If

If Err Then GrabarLog "", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub MDIForm_Load()
On Error Resume Next

'Exit Sub
   
'Me.Show

    mensaje "cargando frmPrincipal"
    
    Call verifico_ult_nrointerno_todasTablas

    mensaje "frmprincipal.1"

    cargarEntorno
    
 
    vParametrosSistema.vFechaInicio = "01/06/2010"
    vParametrosSistema.vFechaFin = Date
    
   ' Me.Picture1.Picture = LoadPicture(App.Path + "\logo3.jpg")


    Me.Caption = LeerConfig(1) & " " & vConfigGral.vempresa & " -- Usuario: " & vConfigGral.vUser

    mensaje "frmprincipal.3"
  
    CargarMenu

    mensaje "frmprincipal.4"
   
    vPFDetalle = True
        
    Me.MousePointer = vbDefault

    'frmAlarmas.Show
    
    Select Case vConfigGral.vImpresoraSeleccionada
    
        Case "Hasar", "Fiscal Hasar"
            With FiscalHasar

                .Puerto = vImpresoras.vNroPuerto
                .Modelo = vImpresoras.vModeloInterno
                .DescripcionesLargas = True
                .Comenzar

                .TratarDeCancelarTodo
                
                
                
                Call .EspecificarNombreDeFantasia(" ", " ")
                
                
                                
            End With
        
        Case "Epson"
        
        Case ""
    
    End Select
    
    
    'If vConfigGral.vIncluyeContabilidad Then frmInconsistencias.Show
        mensaje "frmprincipal.5"
    
    If LeerXml("Puesto") = "Comuna" Or LeerXml("Puesto") = "Caja" Then
        Call frmIngresosEgresos.initEgreso
    End If
    
    If UCase(LeerXml("Puesto")) = "ASOCIAL" Then
        Call frmIngresosEgresos.initIngreso
    End If
    
    
    If Trim(LeerXml("Puesto")) = "Intendente" Then
        frmBancoCajaDetalle.Show
    End If
    

    If UCase(Trim(LeerXml("Puesto"))) = "SERVICIOS" Then
        frmDeudasServicios.Show
    End If

    
    If UCase(LeerXml("Puesto")) = "KIOSCO" Then
        'Me.WindowState = 1
        Call frmRemito.Show
    End If
    
    
    If UCase(LeerXml("Puesto")) = "SERVICIOS" Then
        'Me.WindowState = 1
        Call frmDeudasServicios.Show
    End If
    
    
    
    
   ' Call garbage
   
   'frmPrincipal.Picture1.Height = 8000
    
    Unload frmSplash

    
    If Err Then GrabarLog "MDIForm_load", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cargarEntorno()
On Error Resume Next

mensaje "Entorno"

    Call CambiarCR
    Call DatosEmpresa
    Call LoadConfigRemito
    Call conexionOk
    
    mensaje "Cargar impresora"
    
    Call CargarImpresoras
If Err Then
    MsgBox "Error al cargar el entorno" + Chr(13) + Err.Description, vbCritical
    Exit Sub
End If
End Sub

Public Sub CargarMenu()
On Error Resume Next

    CommandBarsGlobalSettings.App = App
     
    Dim control As CommandBarControl
    
    
    Botonera.DeleteAll
    
    Set TabToolBar = Botonera.AddTabToolBar("Menu")
    
    Dim ToolbarTab As TabControlItem
        
    'Declaro el Grupo (+Tab, +Icono, + Nombre, + IdGrupo)
    Dim i As Integer, vNombreGrupo() As String, vValorGrupo() As Long, vNombreIcono As String, vCantidadGrupos As Integer, vIdFormularioGrupo() As Long
    
    Dim rsFormularioGrupos As New ADODB.Recordset
    
    With rsFormularioGrupos
        .CursorLocation = adUseClient
        
        Call .Open("SELECT * FROM FormularioGrupo WHERE (Habilitado = 'S') ORDER BY  idFormularioGrupo", PathDBConfig, adOpenStatic, adLockReadOnly)
        
        vCantidadGrupos = .RecordCount
        
        ReDim vIdFormularioGrupo(vCantidadGrupos)
        ReDim vNombreGrupo(vCantidadGrupos)
        ReDim vValorGrupo(vCantidadGrupos)
        
        For i = 1 To .RecordCount
            vIdFormularioGrupo(i) = .Fields("idFormularioGrupo").Value
            vNombreGrupo(i) = EsNulo(.Fields("FormularioGrupo").Value)
            vValorGrupo(i) = .Fields("ValorGrupo").Value
        
            
            Set ToolbarTab = TabToolBar.InsertCategory(vValorGrupo(i), vNombreGrupo(i))
            ToolbarTab.Image = .Fields("ValorGrupo").Value 'ID_TAB_ICON
        
            If .Fields("FormularioGrupo").Value = "Favoritos" Then
            
                vMenuFavoritos = .Fields("idFormularioGrupo").Value
            
            End If
            .MoveNext
        Next
    
    End With
    
    If rsFormularioGrupos.State = 1 Then
        rsFormularioGrupos.Close
        Set rsFormularioGrupos = Nothing
    End If

    TabToolBar.EnableDocking (xtpFlagStretched)
    TabToolBar.ShowExpandButton = False
    
    TabToolBar.MinimumWidth = 82 * vCantidadGrupos
    TabToolBar.Category(2).Color = vbRed
    
    Dim j As Integer, iCount As Integer, vToolTip As String, sqlFormularios As String
    
    Dim rsFormularios As New ADODB.Recordset
    
    For j = 1 To vCantidadGrupos
        With rsFormularios
            If .State = 1 Then .Close
            
            .CursorLocation = adUseClient
            
            sqlFormularios = "SELECT * FROM Formularios WHERE (idFormularioGrupo = " & vIdFormularioGrupo(j) & ") AND (Habilitado = 'S') ORDER BY idFormularioGrupo ASC, idFormularios ASC"
            
            Call .Open(sqlFormularios, PathDBConfig, adOpenStatic, adLockReadOnly)
        
            For iCount = 1 To .RecordCount
                vToolTip = "[" + .Fields("Descripcion").Value + "]"
           
                Set control = TabToolBar.Controls.Add(xtpControlButton, Val(vValorGrupo(j) + iCount), vToolTip)
                control.Category = vNombreGrupo(j)
                
                .MoveNext
            Next
        End With
        
        If rsFormularios.State = 1 Then
            rsFormularios.Close
            Set rsFormularios = Nothing
        End If
    Next
    
    With TabToolBar
        .TabPaintManager.Appearance = xtpTabAppearancePropertyPageSelected 'xtpTabAppearancePropertyPage2007
        .CommandBars.VisualTheme = xtpThemeWhidbey 'xtpThemeVisualStudio2008
        .CommandBars.VisualTheme = xtpThemeVisualStudio2008
        
        '.TabPaintManager.OneNoteColors = True
        .TabPaintManager.BoldSelected = True
        
        .TabPaintManager.ToolTipBehaviour = xtpTabToolTipNever
        .CommandBars.RecalcLayout
        .UpdateTabs
    End With
    
    For iCount = 1 To 322
        'IMIconos.Icons.LoadIcon (App.Path & "\Iconos\" & iCount & ".ico"), (iCount), (xtpImageNormal)
    Next iCount

  '  IMIconos.Icons.LoadIcon App.Path & "\Iconos\App.ico", ID_TAB_ICON, xtpImageNormal
    
  '  Botonera.Icons = IMIconos.Icons
    
  ' sacar comentarios
    
    TabToolBar.SetIconSize Val(LeerConfig(19)), Val(LeerConfig(19))
       
    Botonera.Options.UseDisabledIcons = True
    
    Botonera.EnableCustomization (True)

If Err Then GrabarLog "CargarMenu", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub MDIForm_Terminate()
End
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub
Private Sub Botonera_ExecuteVieja(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next
    
    Dim Formulario As Form, vNombreForm As String, vFormActivo As Boolean, vIdFormularioActivo As Long

    vNombreForm = TraerDato("Formularios", "(Descripcion = '" & control.ToolTipText & "')", "Formulario", PathDBConfig)
    vParametro = TraerDato("Formularios", "(Descripcion = '" & control.ToolTipText & "')", "Parametro", PathDBConfig)
    
    vIdFormularioActivo = Val(TraerDato("Formularios", "(Descripcion = '" & control.ToolTipText & "')", "idFormularios", PathDBConfig))
            
    vFormActivo = CBool(TraerDato("FormActivos", "(idFormularios = " & vIdFormularioActivo & ") AND (idUsuarios = " & vConfigGral.vIdUsuario & ") ", "idFormActivos", PathDBConfig))
                
    If Not vFormActivo = True Then
        Call EjecutarScript("INSERT INTO   FormActivos (idUsuarios, idFormularios) VALUES (" & vConfigGral.vIdUsuario & "," & vIdFormularioActivo & ")", PathDBConfig)
        Call EjecutarScript("INSERT INTO FormVisitados (idUsuarios, idFormularios) VALUES (" & vConfigGral.vIdUsuario & "," & vIdFormularioActivo & ")", PathDBConfig)
        
        Set Formulario = CallByName(Forms, "Add", VbMethod, vNombreForm)
        
        CallByName Formulario, "Tag", VbLet, vIdFormularioActivo

    Else
    
        MsgBox "El FORMULARIO ya se encuenta activo para este USUARIO", vbExclamation, "Mensaje ..."
    
    End If
    
If Err Then GrabarLog "Botonera_Execute", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Botonera_Execute(ByVal control As XtremeCommandBars.ICommandBarControl)
On Error Resume Next
    
    Dim Formulario As Form, vNombreForm As String, vFormActivo As Boolean, vIdFormularioActivo As Long
    
    vNombreForm = TraerDato("Formularios", "(Descripcion = '" & Mid(control.ToolTipText, 2, Len(control.ToolTipText) - 2) & "')", "Formulario", PathDBConfig)
    vParametro = TraerDato("Formularios", "(Descripcion = '" & control.ToolTipText & "')", "Parametro", PathDBConfig)
    
    vIdFormularioActivo = Val(TraerDato("Formularios", "(Descripcion = '" & control.ToolTipText & "')", "idFormularios", PathDBConfig))
            
    vFormActivo = CBool(TraerDato("FormActivos", "(idFormularios = " & vIdFormularioActivo & ") AND (idUsuarios = " & vConfigGral.vIdUsuario & ") ", "idFormActivos", PathDBConfig))
                
    If Not vFormActivo = True Then
        'Call EjecutarScript("INSERT INTO   FormActivos (idUsuarios, idFormularios) VALUES (" & vConfigGral.vIdUsuario & "," & vIdFormularioActivo & ")", PathDBConfig)
        Call EjecutarScript("INSERT INTO FormVisitados (idUsuarios, idFormularios) VALUES (" & vConfigGral.vIdUsuario & "," & vIdFormularioActivo & ")", PathDBConfig)
        
                
        Select Case vNombreForm
        
            'Articulos
            Case "frmArticulo"
                frmArticulos.Show
            
            Case "frmConsulta"
                'frmConsulta.Show
            
            Case "frmArticuloProveedor"
                frmArticuloProveedor.Show
                
            Case "frmActualizacionPrecio"
                frmActualizacionPrecio.Show
                
            Case "frmCobros"
                frmCobros.cpInstancia = "cobro"
                frmCobros.Show
                 frmCobros.WindowState = vmaximizar

            Case "frmRubros"
                frmRubros.Show
            
            'Clientes
            Case "frmClientes"
                frmClientes.Show
            Case "frmCtaCteC"
                
                
                frmCtaCteC.Tag = "Clientes"
                frmCtaCteC.Show
                frmCtaCteC.WindowState = vmaximizar

            Case "frmCreditos"
                'frmCreditos.Show
            Case "frmListadoSaldos"
                frmListadoSaldos.instanciaCP = "Clientes"
                frmListadoSaldos.Show
            Case "frmEnvasesAdeudados"
                'frmEnvasesAdeudados.Show
            Case "frmUltimaCompra"
                'frmUltimaCompra.Show
            Case "frmQuebrantos"
                'frmQuebrantos.Show
            Case "frmLibreta"
                'frmLibreta.Show
        
            'Proveedores
            Case "frmProveedores"
                frmProveedores.Show
                frmProveedores.Tag = vIdFormularioActivo
            Case "frmSaldosProveedores"
                'frmSaldosProveedores.Show
                'frmSaldosProveedores.Tag = vIdFormularioActivo
                frmListadoSaldos.instanciaCP = "Proveedores"
                frmListadoSaldos.Show
                
                
            Case "frmCtaCteP"
                
               ' frmCtaCteP.Tag = vIdFormularioActivo
                
                frmCtaCteC.Tag = "Proveedores"
                frmCtaCteC.Show
                frmCtaCteC.WindowState = vmaximizar
                frmCtaCteC.init
                
            'Empleados
            Case "frmEmpleados"
                frmEmpleados.Show
                frmEmpleados.Tag = vIdFormularioActivo
 
            Case "frmLiquidacionSueldos"
                'frmLiquidacionSueldos.Show
                'frmLiquidacionSueldos.Tag = vIdFormularioActivo
 
            Case "frmClienteRepartidor"
                'frmClienteRepartidor.Show
                'frmClienteRepartidor.Tag = vIdFormularioActivo
                 
            'Case "frmControlEnvases"
            '    frmControlEnvases.Show
            '    frmControlEnvases.Tag = vIdFormularioActivo
                 
            Case "frmListadoEnvaceRepartidor"
                'frmListadoEnvaceRepartidor.Show
                'frmListadoEnvaceRepartidor.Tag = vIdFormularioActivo
                            
            'Ventas/Comras
            Case "frmRemito"
                If vConfigGral.vIncluyeResto = True Then
                    frmRemitoResto.Show
                    frmRemitoResto.Tag = vIdFormularioActivo
                Else
                    'Panic: Poner un if para usar con algunos clientes
                    'Botonera.DeleteAll
                    frmRemito.Show
                    frmRemito.Tag = vIdFormularioActivo
                End If
                
            Case "frmBuscarFactura"
                frmBuscarFactura.cpFactura = "factura"
                frmBuscarFactura.CP = "clientes"
                
                frmBuscarFactura.Show
                frmBuscarFactura.Tag = vIdFormularioActivo
                frmBuscarFactura.lblDocumento(0).Caption = "> Cliente:"
                frmBuscarFactura.Caption = "Listado de documentos de Ventas"
                

        
            Case "frmCierresXZ"
                frmCierresXZ.Show
                frmCierresXZ.Tag = vIdFormularioActivo
                
            Case "frmControlFacturacion"
                'frmControlFacturacion.Show
                'frmControlFacturacion.Tag = vIdFormularioActivo
            
            Case "frmCompras"
    
                frmCompras.Show
                frmCompras.Tag = vIdFormularioActivo
            
                

                
            Case "frmBuscarCompra"
                
                frmBuscarFactura.CP = "proveedores"
                frmBuscarFactura.cpFactura = "pfactura"
                
                frmBuscarFactura.Show
                frmBuscarFactura.Tag = vIdFormularioActivo
                                
                frmBuscarFactura.lblDocumento(0).Caption = "> Proveedores:"
                frmBuscarFactura.Caption = "Listado de documentos de Compras"
                
                
            Case "frmPagos" ' llamo al frmcobros
                frmCobros.cpInstancia = "pagos" 'Alfredo: seteo la variable para que inicie como un pago a proveedor
                frmCobros.Show
                 frmCobros.WindowState = vmaximizar
                frmCobros.Tag = vIdFormularioActivo
                
            'Cheques
            Case "frmChequesAlta"
                frmChequesAlta.vViene = ""
                frmChequesAlta.Show
                frmChequesAlta.Tag = vIdFormularioActivo
 
            Case "frmCheques"
                frmCheques.Show
                frmCheques.WindowState = vmaximizar
                frmCheques.Tag = vIdFormularioActivo
            
            'Estadisticas
            Case "frmEstadisticaProducto"
                'frmEstadisticaProducto.Show
                'frmEstadisticaProducto.Tag = vIdFormularioActivo
 
            Case "frmEstadisticaCliente"
                'frmEstadisticaCliente.Show
                'frmEstadisticaCliente.Tag = vIdFormularioActivo
 
            Case "frmEstadisticasGral"
                'frmEstadisticasGral.Show
                'frmEstadisticasGral.Tag = vIdFormularioActivo
 
            Case "frmSaldosTotales"
                frmSaldosTotales.Show
                frmSaldosTotales.Tag = vIdFormularioActivo

            'Listados de Iva
            Case "frmIvaVenta"
                frmIvaVenta.Show
                frmIvaVenta.Tag = vIdFormularioActivo
 
            Case "frmIvaCompra"
                frmIvaCompra.Show
                frmIvaCompra.Tag = vIdFormularioActivo
 
            'Caja
            Case "frmCaja"
                frmCaja.Show
                frmCaja.Tag = vIdFormularioActivo
            Case "frmEstructuraCaja"
                frmEstructuraCaja.Show
                frmEstructuraCaja.Tag = vIdFormularioActivo
            Case "frmGastosConcepto"
                'frmGastosConcepto.Show
                'frmGastosConcepto.Tag = vIdFormularioActivo
            
            'Bancos
            Case "frmBancosAlta"
                frmBancosAlta.Show
                frmBancosAlta.Tag = vIdFormularioActivo
            Case "frmBancosCuentaAlta"
                frmBancosCuentaAlta.Show
                frmBancosCuentaAlta.Tag = vIdFormularioActivo
            
            Case "frmBancos"
                frmBancos.Show
                frmBancos.Tag = vIdFormularioActivo
                
            Case "frmBancosMovimientos"
                frmBancosMovimientos.Show
                frmBancosMovimientos.Tag = vIdFormularioActivo
                
            Case "frmCajaMovimientos"
                frmCajaMovimientos.Show
                frmCajaMovimientos.Tag = vIdFormularioActivo
                
            Case "frmIngresosEgresos"
                frmIngresosEgresos.Show
                frmIngresosEgresos.Tag = vIdFormularioActivo
                            

            Case "frmBancoCajaDetalle"
                frmBancoCajaDetalle.Show
                frmBancoCajaDetalle.Tag = vIdFormularioActivo
                                                                
            'Contabilidad
            Case "frmPeriodoContable"
                frmPeriodoContable.Show
            
            
            Case "frmCuentas"
                frmCuentas.Show
            Case "frmAsientosAlta"
                frmAsientosAlta.Show
            Case "frmAsientos"
                frmAsientos.Show
            Case "frmMovientosDiarios"
                frmMovientosDiarios.Show
            Case "frmMovimientosCuentas"
                frmMovimientosCuentas.Show
            Case "frmBalance"
                frmBalance.Show
                
            Case "frmResultados"
                frmResultados.Show
            
            'Archivo
            
            Case "frmMigracion"
                frmMigracion.Show
                frmMigracion.Tag = "frmMigracion"
            
            Case "frmBackup"
                frmBackup.Show
                frmBackup.Tag = vIdFormularioActivo
             
            'Case "frmBackupCD"
            '    frmBackupCD.Show
            '    frmBackupCD.Tag = vIdFormularioActivo
             
            Case "frmAgenda"
                frmAgenda.Show
                frmAgenda.Tag = vIdFormularioActivo
             
            Case "frmGuia"
                'frmGuia.Show
                'frmGuia.Tag = vIdFormularioActivo
             
            Case "frmBorrarBases"
                frmBorrarBases.Show
                frmBorrarBases.Tag = vIdFormularioActivo
             
            Case "frmConfigurar"
                'frmConfigurar.Show
                'frmConfigurar.Tag = vIdFormularioActivo
             
             Case "frmCotizacion"
                frmCotizacion.Show
                
            'Ayuda
            Case "frmAyuda"
                'frmAyuda.Show
                'frmAyuda.Tag = vIdFormularioActivo
             
            Case "frmAbout"
                'frmAbout.Show
                'frmAbout.Tag = vIdFormularioActivo
            
            'Salir
            Case "frmOut"
                frmOut.Show
            Case "frmOut"
                 End
            Case "frmMonitor"
                frmMonitor.Show
               
        End Select


    Else
    
        MsgBox "El FORMULARIO ya se encuenta activo para este USUARIO", vbExclamation, "Mensaje ..."
    
    End If

If Err Then
   Exit Sub
    MsgBox "Error inesperado al cargar el menú principal" + Err.Description, vbCritical
    GrabarLog "Botonera_Execute", Err.Number & " " & Err.Description, Me.Name
End If
End Sub

Private Sub scp_Click()
    frmSaldosClientes.Show
End Sub

Private Sub moviBancos_Click()
    frmBancosMovimientos.Show
End Sub

Private Sub moviCaja_Click()
    frmCajaMovimientos.Show
End Sub

Private Sub movimientosPorCtas_Click()
    frmMovimientosCuentas.Show
End Sub

Private Sub OtrosEgresos_Click()

If LeerXml("Puesto") = "Comuna" Or LeerXml("Puesto") = "Caja" Or LeerXml("Puesto") = "ASOCIAL" Then
        Call frmIngresosEgresos.initIngreso
Else
        MsgBox ("Este puesto no tiene autorizaciòn." + Chr(13) + "Consulte al Servicio Tècnico")
End If
End Sub

Private Sub pagoProveedor_Click()
  frmCobros.cpInstancia = "pagos" 'Alfredo: seteo la variable para que inicie como un pago a proveedor
  frmCobros.Show
End Sub

Private Sub presupuestadovsgastos_Click()
frmBalance.Show
frmBalance.rdPresupuestado.Value = True
End Sub

Private Sub presupuesto_Click()
frmPresupuesto.Show
frmPresupuesto.tab.SelectedItem = 1
End Sub

Private Sub presupuestos_Click()
On Error Resume Next
    frmPresupuesto.Show
If Err Then Exit Sub
End Sub

Private Sub prueba1_Click()
Dim msg As String
Dim j As Integer

On Error GoTo impresora_apag
       
Procesar:
    
    
    frmPrincipal.FiscalHasar.AbrirComprobanteNoFiscal
    
    For j = 1 To 1
        frmPrincipal.FiscalHasar.ImprimirTextoNoFiscal "Linea de Texto No Fiscal @@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
    Next j
    
    frmPrincipal.FiscalHasar.CerrarComprobanteNoFiscal
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

Private Sub pruebaepson_Click()

Dim cCmd, sCmExt, vmensaje As String
Dim bAnswer As Boolean
bAnswer = True


With frmPrincipal.FiscalEpson2
         
         
                            .ClosePort
                            
                            Call FPDelay
                            
                            .CommPort = LeerXml("Puerto")
                           ' .BaudRate = 3
                            
                            .ProtocolType = protocol_Extended
                            
                            
                            vmensaje = " Puerto : " + Str(.CommPort) + "  - " + Str(.BaudRate)
                            
                            MsgBox vmensaje
                            
                            
                            If (.OpenPort) Then
                                Call FPDelay
                            Else
                                MsgBox "2- El controlador fiscal no está conectado. " + Chr(13) + _
                                "Conecte el controlador y vuelva a ingresar a este módulo"
                            End If
         
         
            sCmd = Chr$(&HA) + Chr$(&H1)
            If bAnswer Then bAnswer = .AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H0)
            If bAnswer Then bAnswer = .AddDataField(sCmdExt)
            If bAnswer Then bAnswer = .SendCommand
            Call FPDelay
            If .ReturnCode <> 0 Then ShowMsg
            
            'Item
            sCmd = Chr$(&HA) + Chr$(&H2)
            If bAnswer Then bAnswer = .AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H0)
            If bAnswer Then bAnswer = .AddDataField(sCmdExt)
            If bAnswer Then bAnswer = .AddDataField("a")
            If bAnswer Then bAnswer = .AddDataField("b")
            If bAnswer Then bAnswer = .AddDataField("c")
            If bAnswer Then bAnswer = .AddDataField("d")
            If bAnswer Then bAnswer = .AddDataField("Descripción Item")
            If bAnswer Then bAnswer = .AddDataField("10000")
            If bAnswer Then bAnswer = .AddDataField("1000")
            If bAnswer Then bAnswer = .AddDataField("2100")
            If bAnswer Then bAnswer = .AddDataField("")
            If bAnswer Then bAnswer = .AddDataField("")
            If bAnswer Then bAnswer = .SendCommand ' comando de items
            Call FPDelay
            If .ReturnCode <> 0 Then ShowMsg
            
            'Payment
            sCmd = Chr$(&HA) + Chr$(&H5)
            If bAnswer Then bAnswer = .AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H0)
            If bAnswer Then bAnswer = .AddDataField(sCmdExt)
            If bAnswer Then bAnswer = .AddDataField("")
            If bAnswer Then bAnswer = .AddDataField("EFECTIVO")
            If bAnswer Then bAnswer = .AddDataField("500")
            If bAnswer Then bAnswer = .SendCommand
            Call FPDelay
            If .ReturnCode <> 0 Then ShowMsg
            
            'Close
            sCmd = Chr$(&HA) + Chr$(&H6)
            If bAnswer Then bAnswer = .AddDataField(sCmd)
            sCmdExt = Chr$(&H0) + Chr$(&H1)
            If bAnswer Then bAnswer = .AddDataField(sCmdExt)
            If bAnswer Then bAnswer = .AddDataField(1)
            If bAnswer Then bAnswer = .AddDataField("-")
            If bAnswer Then bAnswer = .AddDataField(2)
            If bAnswer Then bAnswer = .AddDataField("-")
            If bAnswer Then bAnswer = .AddDataField(3)
            If bAnswer Then bAnswer = .AddDataField("-")
            If bAnswer Then bAnswer = .SendCommand
            Call FPDelay
            If .ReturnCode <> 0 Then ShowMsg

End With


End Sub

Private Sub pruebafiscal_Click()
Dim msg As String

On Error GoTo impresora_apag
       
Procesar:
    
   ' log.AddItem "1"
    
    With frmPrincipal.FiscalHasar
                     .DatosCliente "Nombre del Cliente", "20249182940", TIPO_CUIT, RESPONSABLE_INSCRIPTO, "Domicilio: Donde Siempre"
                     
                     'HASAR2.InformacionRemito(1) = "1234-56789012"
                     
                  '   log.AddItem "22"
                    
                     .AbrirComprobanteFiscal TICKET_FACTURA_A
                     
                   '  log.AddItem "3"
                     
                     .ImprimirItem "Producto a la Venta: Uno", 1, 1, 21, 0
                   '  log.AddItem "4"
                     
                     .ImprimirPago "Efectivo", 1
                  '   log.AddItem "5"
                     .CerrarComprobanteFiscal
                   '  log.AddItem "6"

    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description + " codigo: " + Str(Err.Number), vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
        .CerrarComprobanteFiscal
    
    End If
  End With

End Sub

Private Sub redifae_Click()

    frmFiltro.vcomando = "fae"
    frmFiltro.Show
    
    'Call faedetalle
End Sub

Private Sub saldocb_Click()
    frmBancoCajaDetalle.Show
    frmBancoCajaDetalle.vtipolistado = "Agrupado por Cajas-Bancos"
    Call frmBancoCajaDetalle.cmdFiltrar_Click
End Sub

Private Sub tablero_Click()
      '  Call frmTablero.Show
      '  Call frmTablero.Form_Initialize
      '  Call frmTablero.PusActualizar_Click
End Sub

Private Sub tgi_Click()
    frmTransaccionMantenimiento.Show
End Sub


Private Sub webcam_Click()
frmWebCam.Show
End Sub

Private Sub wsfe_Click()
frmFeStatus.Show
End Sub
