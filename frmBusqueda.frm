VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#13.0#0"; "Codejock.DockingPane.v13.0.0.Demo.ocx"
Begin VB.Form frmBusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccion de ..."
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   10890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   10890
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   5715
      Left            =   30
      TabIndex        =   1
      Top             =   420
      Width           =   10815
      _Version        =   851968
      _ExtentX        =   19076
      _ExtentY        =   10081
      _StockProps     =   68
      ItemCount       =   1
      Item(0).Caption =   "Buscar por:"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "GroupBox1"
      Item(0).Control(1)=   "txtBusqueda"
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   4695
         Left            =   90
         TabIndex        =   2
         Top             =   870
         Width           =   10635
         _Version        =   851968
         _ExtentX        =   18759
         _ExtentY        =   8281
         _StockProps     =   79
         BackColor       =   -2147483644
         UseVisualStyle  =   -1  'True
         Begin MSDataGridLib.DataGrid dgBusqueda 
            Height          =   4470
            Left            =   60
            TabIndex        =   3
            Top             =   150
            Width           =   10515
            _ExtentX        =   18547
            _ExtentY        =   7885
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777215
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   15
            RowDividerStyle =   4
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
      End
      Begin XtremeSuiteControls.FlatEdit txtBusqueda 
         Height          =   315
         Left            =   120
         TabIndex        =   4
         Top             =   480
         Width           =   10515
         _Version        =   851968
         _ExtentX        =   18547
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":1A188
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":209EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":2724C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":2DAAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":34310
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":3AB72
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBusqueda.frx":413D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Botonera 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10890
      _ExtentX        =   19209
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo Registro"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar Registro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Borrar Registro"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Seleccionar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Ayuda"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin XtremeDockingPane.DockingPane DockingPane1 
      Left            =   0
      Top             =   540
      _Version        =   851968
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsBusqueda As ADODB.Recordset
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Dim arrPanes(1 To 4) As frmTaskPanel
Private Sub Nuevo()
On Error Resume Next
    
    '20100519-1300
    frmAltas.vaccion = "Nuevo"
    
    Select Case vVieneBusqueda
    
        Case "CodigoPostal"
            frmAltas.vcampos = 2
        
        Case "Vendedor"

        Case "Reparto"
            frmAltas.vcampos = 1
            
        Case "TipoIva"
            frmAltas.vcampos = 1
            
        Case "TipoCliente"
            frmAltas.vcampos = 1
        
        Case "Actividad"
            frmAltas.vcampos = 1
        
        Case "Lista"
            frmAltas.vcampos = 1

        Case "EstadoCliente"
            frmAltas.vcampos = 3

        
        Case "Rubro"
            frmAltas.vcampos = 1
            
        Case "Proveedor"
            frmProveedoresAlta.Show
            frmProveedoresAlta.vaccion = "Nuevo"

        Case "Fabricante"
            frmAltas.vcampos = 1

        Case "PorcentajeIva"
            frmAltas.vcampos = 2

        Case "Cotizacion"
            frmAltas.vcampos = 2
        
        Case "Mozo"
            frmAltas.vcampos = 1

        Case "SubRubro"
            frmAltas.vcampos = 1
        
        Case "EstadoCheque"
            frmAltas.vcampos = 1
            
        Case "TipoMovimientosBanco"
            frmAltas.vcampos = 5
        
        Case "CodigoCliente"
            frmClientesAlta.Show
            frmClientesAlta.vaccion = "Nuevo"
            frmClientesAlta.vVieneClientesAlta = Me.Name
    
    End Select
    
    
    If Not vVieneBusqueda = "CodigoCliente" And Not vVieneBusqueda = "Proveedor" Then
    
        frmAltas.Show

    End If
If Err Then GrabarLog "Nuevo", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Modificar()
On Error Resume Next

If Err Then GrabarLog "Modificar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub BorrarRegistro()
On Error Resume Next
    
    Call BorrarRecordset(rsBusqueda, True)

If Err Then GrabarLog "BorrarRegistro", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Seleccionar()
On Error Resume Next
    
    dgBusqueda_DblClick

If Err Then GrabarLog "Seleccionar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Buscar()
On Error Resume Next


If Err Then GrabarLog "Buscar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub dgBusqueda_DblClick()

' Alfredo: siempre vas a tener que hacer un case mas para el formulario que estas programando
On Error Resume Next

    With rsBusqueda
    
        If Not .EOF = True And Not .BOF = True Then
            
            Select Case vVuelveBusqueda
                
                Case "frmAsientosAlta"
                    Select Case vVieneBusqueda
                    
                        Case "CodigoCuenta"
                            frmAsientosAlta.txtModificar(0).Text = EsNulo(.Fields("CodigoCuenta").Value)
                            frmAsientosAlta.txtModificar(1).Text = EsNulo(.Fields("Cuenta").Value)
                            frmAsientosAlta.txtModificar(2).SetFocus
                    End Select
                
                Case "frmBalance"
                    
                    Select Case vVieneBusqueda

                        Case "CodigoCuentaD"
                            frmBalance.txtCuentaContable(0).Text = EsNulo(.Fields("CodigoCuenta").Value)
                            frmBalance.txtCuentaContable(1).Text = EsNulo(.Fields("Cuenta").Value)
                                
                        Case "CodigoCuentaH"
                            frmBalance.txtCuentaContable(2).Text = EsNulo(.Fields("CodigoCuenta").Value)
                            frmBalance.txtCuentaContable(3).Text = EsNulo(.Fields("Cuenta").Value)
                    End Select
                                
                Case "frmMovimientosCuentas"
                    
                    Case "CodigoCuentaD"
                        frmMovimientosCuentas.vnro_asto0.Text = EsNulo(.Fields("CodigoCuenta").Value)
                        frmMovimientosCuentas.vnro_asto1.Text = EsNulo(.Fields("Cuenta").Value)
                                
                    Case "CodigoCuentaH"
                        frmMovimientosCuentas.vnro_asto2.Text = EsNulo(.Fields("CodigoCuenta").Value)
                        frmMovimientosCuentas.vnro_asto3.Text = EsNulo(.Fields("Cuenta").Value)
                    
                
                Case "frmCompras"
                        
                    Select Case vVieneBusqueda
                    
                        Case "Banco"
                            frmCompras.txtCaja(0).Text = EsNulo(.Fields("idBancos").Value)
                            frmCompras.txtCaja(1).Text = EsNulo(.Fields("Descripcion").Value)
                        
                        Case "BancoCuenta"
                            frmCompras.txtCaja(2).Text = EsNulo(.Fields("idBancosCuentas").Value)
                            frmCompras.txtCaja(3).Text = EsNulo(.Fields("Cuenta").Value)


                        Case "TipoValor"
                            frmCompras.txtCaja(0).Text = EsNulo(.Fields("idTipoValor").Value)
                            frmCompras.txtCaja(1).Text = EsNulo(.Fields("Descripcion").Value)
                        
                        Case "TipoMovimientos"
                         frmCompras.txtTipoMovimiento(0).Text = EsNulo(.Fields("codigo").Value)
                         frmCompras.txtTipoMovimiento(1).Text = EsNulo(.Fields("TipoMovimiento").Value)
                            
                        Case "compra-caja"
                         frmCompras.txtBancoCheque(0).Text = EsNulo(.Fields("idBancos").Value)
                         frmCompras.txtBancoCheque(1).Text = EsNulo(.Fields("Descripcion").Value)
                            
                    
                    End Select
                Case "frmCobros"
                        
                        Select Case vVieneBusqueda
                        
                        
                            Case "Caja"
                                frmCobros.txtBancoCheque(4).Text = EsNulo(.Fields("idBancos").Value)
                                frmCobros.txtBancoCheque(5).Text = EsNulo(.Fields("Descripcion").Value)
                        
                            Case "Banco"
                                frmCobros.txtBancoCheque(0).Text = EsNulo(.Fields("idBancos").Value)
                                frmCobros.txtBancoCheque(1).Text = EsNulo(.Fields("Descripcion").Value)
                            
                            Case "BancoCuenta"
                                frmCobros.txtBancoCheque(2).Text = EsNulo(.Fields("idBancosCuentas").Value)
                                frmCobros.txtBancoCheque(3).Text = EsNulo(.Fields("Cuenta").Value)

                            Case "caja-importe-cobro"
                                frmCobros.txtBancoCheque(6).Text = EsNulo(.Fields("idBancos").Value)
                                frmCobros.txtBancoCheque(7).Text = EsNulo(.Fields("Descripcion").Value)


                            Case "BancoDeposito"
                                frmCobros.txtDepositoBanco(0).Text = EsNulo(.Fields("idBancos").Value)
                                frmCobros.txtDepositoBanco(1).Text = EsNulo(.Fields("Descripcion").Value)
                            
                            Case "BancoCuentaDeposito"
                                frmCobros.txtDepositoBanco(2).Text = EsNulo(.Fields("idBancosCuentas").Value)
                                frmCobros.txtDepositoBanco(3).Text = EsNulo(.Fields("Cuenta").Value)
                                frmCobros.txtDepositoImporte.SetFocus
                                
                        End Select
                        frmCobros.WindowState = vmaximizar
                Case "frmPagos"
                        
                        Select Case vVieneBusqueda
                        
                            Case "Banco"
                                frmPagos.txtBancoCheque(0).Text = EsNulo(.Fields("idBancos").Value)
                                frmPagos.txtBancoCheque(1).Text = EsNulo(.Fields("Descripcion").Value)
                            
                            Case "BancoCuenta"
                                frmPagos.txtBancoCheque(2).Text = EsNulo(.Fields("idBancosCuentas").Value)
                                frmPagos.txtBancoCheque(3).Text = EsNulo(.Fields("Cuenta").Value)

                            Case "BancoDeposito"
                                frmPagos.txtDepositoBanco(0).Text = EsNulo(.Fields("idBancos").Value)
                                frmPagos.txtDepositoBanco(1).Text = EsNulo(.Fields("Descripcion").Value)
                        
                            Case "BancoCuentaDeposito"
                                frmPagos.txtDepositoBanco(2).Text = EsNulo(.Fields("idBancosCuentas").Value)
                                frmPagos.txtDepositoBanco(3).Text = EsNulo(.Fields("Cuenta").Value)
                        
                        End Select
                    

                    
                Case "frmBuscarCompra"
                    Select Case vVieneBusqueda
                    
                        Case "Proveedor"
                            frmBuscarCompra.txtProveedor(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmBuscarCompra.txtProveedor(1).Text = EsNulo(.Fields("Nombre").Value)
                            frmBuscarCompra.cmdBuscaryCalcular.SetFocus
                            
                    End Select
                
                Case "frmChequesAlta" '  Alfredo: acá verifica el nombre del formulario donde tiene que ir
                    Select Case vVieneBusqueda 'Alfredo: acá discrimina cual del los botones buscar en tabla es
                    
                        Case "CodigoCliente"
                            frmChequesAlta.txtFicha(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmChequesAlta.txtFicha(1).Text = EsNulo(.Fields("Nombre").Value)
                            frmChequesAlta.txtFicha(2).Text = ""
                            frmChequesAlta.txtFicha(3).Text = ""
                            frmChequesAlta.txtFicha(0).Enabled = True
                            frmChequesAlta.txtFicha(1).Enabled = True
                            frmChequesAlta.txtFicha(2).Enabled = False
                            frmChequesAlta.txtFicha(3).Enabled = False
                            
                            frmChequesAlta.txtFicha(4).SetFocus
                        Case "Proveedor"
                            frmChequesAlta.txtFicha(2).Text = EsNulo(.Fields("Codigo").Value)
                            frmChequesAlta.txtFicha(3).Text = EsNulo(.Fields("Nombre").Value)
                            frmChequesAlta.txtFicha(0).Enabled = False
                            frmChequesAlta.txtFicha(1).Enabled = False
                            frmChequesAlta.txtFicha(2).Enabled = True
                            frmChequesAlta.txtFicha(3).Enabled = True
                            
                            frmChequesAlta.txtFicha(2).Text = EsNulo(.Fields("Codigo").Value)
                            frmChequesAlta.txtFicha(3).Text = EsNulo(.Fields("Nombre").Value)
                    
                            frmChequesAlta.txtFicha(4).SetFocus
                    
                        Case "EstadoCheque"
                            frmChequesAlta.txtFicha(7).Text = EsNulo(.Fields("idEstadoCheque").Value)
                            frmChequesAlta.txtFicha(8).Text = EsNulo(.Fields("Descripcion").Value)

                        Case "Banco"
                            frmChequesAlta.txtFicha(9).Text = EsNulo(.Fields("idBancos").Value)
                            frmChequesAlta.txtFicha(10).Text = EsNulo(.Fields("Descripcion").Value)
                    
                        Case "BancoCuenta"
                            frmChequesAlta.txtFicha(11).Text = EsNulo(.Fields("idBancosCuentas").Value)
                            frmChequesAlta.txtFicha(12).Text = EsNulo(.Fields("Cuenta").Value)
                    End Select




                Case "frmCheques" '  Alfredo: acá verifica el nombre del formulario donde tiene que ir
                    Select Case vVieneBusqueda 'Alfredo: acá discrimina cual del los botones buscar en tabla es
                    
                        Case "CodigoCliente"
                            frmCheques.txtFicha(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmCheques.txtFicha(1).Text = EsNulo(.Fields("Nombre").Value)
                            frmCheques.txtFicha(2).Text = ""
                            frmCheques.txtFicha(3).Text = ""
                            frmCheques.txtFicha(0).Enabled = True
                            frmCheques.txtFicha(1).Enabled = True
                            frmCheques.txtFicha(2).Enabled = False
                            frmCheques.txtFicha(3).Enabled = False
                            
                            frmCheques.txtFicha(4).SetFocus
                        Case "Proveedor"
                            frmCheques.txtFicha(2).Text = EsNulo(.Fields("Codigo").Value)
                            frmCheques.txtFicha(3).Text = EsNulo(.Fields("Nombre").Value)
                            frmCheques.txtFicha(0).Enabled = False
                            frmCheques.txtFicha(1).Enabled = False
                            frmCheques.txtFicha(2).Enabled = True
                            frmCheques.txtFicha(3).Enabled = True
                            
                            frmCheques.txtFicha(2).Text = EsNulo(.Fields("Codigo").Value)
                            frmCheques.txtFicha(3).Text = EsNulo(.Fields("Nombre").Value)
                    
                            frmCheques.txtFicha(4).SetFocus
                    
                        Case "EstadoCheque"
                            frmCheques.txtFicha(7).Text = EsNulo(.Fields("idEstadoCheque").Value)
                            frmCheques.txtFicha(8).Text = EsNulo(.Fields("Descripcion").Value)

                        Case "Banco"
                            frmCheques.txtFicha(9).Text = EsNulo(.Fields("idBancos").Value)
                            frmCheques.txtFicha(10).Text = EsNulo(.Fields("Descripcion").Value)
                    
                        Case "BancoCuenta"
                            frmCheques.txtFicha(11).Text = EsNulo(.Fields("idBancosCuentas").Value)
                            frmCheques.txtFicha(12).Text = EsNulo(.Fields("Cuenta").Value)
                    End Select

                            frmCheques.WindowState = vmaximizar
                
                Case "frmIngresosEgresos"
                    Select Case vVieneBusqueda
                    
                        Case "TipoMovimientosBanco"
                            frmIngresosEgresos.txtAlta(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmIngresosEgresos.txtAlta(1).Text = EsNulo(.Fields("TipoMovimiento").Value)
                            
                            If EsNulo(.Fields("IngresoEgreso").Value) = "I" Then
                                frmIngresosEgresos.RBIngresoEgresoCaja(0).Value = True
                            Else
                                frmIngresosEgresos.RBIngresoEgresoCaja(1).Value = True
                            End If
                            
                            frmIngresosEgresos.txtAlta(2).SetFocus
                            
                        Case "TipoValor"
                            frmIngresosEgresos.txtAlta(3).Text = EsNulo(.Fields("idTipoValor").Value)
                            frmIngresosEgresos.txtAlta(4).Text = EsNulo(.Fields("TipoValor").Value)
                            
                            If Not UCase(EsNulo(.Fields("idTipoValor").Value)) = "CH" Then
                                frmIngresosEgresos.txtAlta(5).Text = ""
                                frmIngresosEgresos.dtpValor.Text = ""
                                frmIngresosEgresos.txtAlta(6).SetFocus
                            Else
                                frmIngresosEgresos.txtAlta(5).SetFocus
                            End If

                         Case "CajaBanco"
                            frmIngresosEgresos.txtAlta(6).Text = EsNulo(.Fields("idBancos").Value)
                            frmIngresosEgresos.txtAlta(7).Text = EsNulo(.Fields("Descripcion").Value)
                            frmIngresosEgresos.txtAlta(8).Text = ""
                            frmIngresosEgresos.txtAlta(9).Text = ""
                            frmIngresosEgresos.txtAlta(10).Text = ""
                            frmIngresosEgresos.txtAlta(11).Text = ""
                            
                            If EsNulo(.Fields("EsCaja").Value) = "S" Then
                                frmIngresosEgresos.pbCarga(3).Enabled = False
                                frmIngresosEgresos.txtAlta(8).Enabled = False
                                frmIngresosEgresos.txtAlta(9).Enabled = False
                                frmIngresosEgresos.txtAlta(10).Text = EsNulo(.Fields("CuentaContableAsociada").Value)
                            Else
                                frmIngresosEgresos.pbCarga(3).Enabled = True
                                frmIngresosEgresos.txtAlta(8).Enabled = True
                                frmIngresosEgresos.txtAlta(9).Enabled = True
                            End If
                                
                         Case "BancoCuenta"
                            frmIngresosEgresos.txtAlta(8).Text = EsNulo(.Fields("idBancosCuentas").Value)
                            frmIngresosEgresos.txtAlta(9).Text = EsNulo(.Fields("Cuenta").Value)
                            frmIngresosEgresos.txtAlta(10).Text = EsNulo(.Fields("CuentaContableAsociada").Value)
                            
                        Case "CodigoCuenta"
                            frmIngresosEgresos.txtAlta(10).Text = EsNulo(.Fields("CodigoCuenta").Value)
                            frmIngresosEgresos.txtAlta(11).Text = EsNulo(.Fields("Cuenta").Value)
                            frmIngresosEgresos.txtAlta(12).SetFocus
                    End Select
                
                
                Case "frmActualizacionPrecio"
                    
                    Select Case vVieneBusqueda
                        
                        Case "SubRubroD"
                            frmActualizacionPrecio.txtFiltro(1).Text = EsNulo(.Fields("idSubRubros").Value)
                            frmActualizacionPrecio.txtFiltro(2).Text = EsNulo(.Fields("SubRubro").Value)
                    
                        Case "SubRubroH"
                            frmActualizacionPrecio.txtFiltro(3).Text = EsNulo(.Fields("idSubRubros").Value)
                            frmActualizacionPrecio.txtFiltro(4).Text = EsNulo(.Fields("SubRubro").Value)
                            
                        Case "RubroD"
                            frmActualizacionPrecio.txtFiltro(5).Text = EsNulo(.Fields("idRubros").Value)
                            frmActualizacionPrecio.txtFiltro(6).Text = EsNulo(.Fields("Rubro").Value)
                        
                        Case "RubroH"
                            frmActualizacionPrecio.txtFiltro(7).Text = EsNulo(.Fields("idRubros").Value)
                            frmActualizacionPrecio.txtFiltro(8).Text = EsNulo(.Fields("Rubro").Value)
                            
                        Case "ProveedorD"
                            frmActualizacionPrecio.txtFiltro(9).Text = EsNulo(.Fields("Codigo").Value)
                            frmActualizacionPrecio.txtFiltro(10).Text = EsNulo(.Fields("Nombre").Value)
                        
                        Case "ProveedorH"
                            frmActualizacionPrecio.txtFiltro(11).Text = EsNulo(.Fields("Codigo").Value)
                            frmActualizacionPrecio.txtFiltro(12).Text = EsNulo(.Fields("Nombre").Value)
                    
                    End Select
                
                Case "frmRemitoResto"
                    Select Case vVieneBusqueda
                    
                        Case "CodigoCliente"
                            Call frmRemitoResto.CargarDatosCliente(EsNulo(.Fields("Codigo").Value), False)
                        Case "Mozo"
                            frmRemitoResto.txtEmpleado(0).Text = EsNulo(.Fields("idMozos").Value)
                            frmRemitoResto.txtEmpleado(1).Text = EsNulo(.Fields("Mozo").Value)
                    End Select
                    
                Case "frmBancosMovimientos"
                    Select Case vVieneBusqueda
                    
                        Case "Banco"
                            frmBancosMovimientos.txtBanco(0).Text = EsNulo(.Fields("idBancos").Value)
                            frmBancosMovimientos.txtBanco(1).Text = EsNulo(.Fields("Descripcion").Value)
                            
                        Case "BancoCuenta"
                            frmBancosMovimientos.txtBanco(2).Text = Val(.Fields("idBancosCuentas").Value)
                            frmBancosMovimientos.txtBanco(3).Text = EsNulo(.Fields("Cuenta").Value)
                            
                    End Select
                        
                Case "frmCaja"
                    Select Case vVieneBusqueda
                    
                        Case "Proveedor"
                            frmCaja.txtProveedor(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmCaja.txtProveedor(1).Text = EsNulo(.Fields("Nombre").Value)
                            frmCaja.txtProveedor(2).Text = EsNulo(.Fields("Direccion").Value)
                            frmCaja.txtProveedor(3).Text = EsNulo(.Fields("Localidad").Value)
                            frmCaja.txtProveedor(4).Text = EsNulo(.Fields("Telefono").Value) & EsNulo(.Fields("Celular").Value)
                            frmCaja.txtProveedor(5).Text = EsNulo(.Fields("Cuit").Value)
                            frmCaja.cboTipoIva(0).Tag = EsNulo(.Fields("idTipoIva").Value)
                            frmCaja.cboTipoIva(0).Text = TraerDato("TipoIva", "idTipoIva = " & EsNulo(.Fields("idTipoIva").Value) & "", "TipoIva")
                    
                        Case "CodigoCliente"
                            frmCaja.txtCliente(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmCaja.txtCliente(1).Text = EsNulo(.Fields("Nombre").Value)
                            frmCaja.txtCliente(2).Text = EsNulo(.Fields("Direccion").Value)
                            frmCaja.txtCliente(3).Text = EsNulo(.Fields("Localidad").Value)
                            frmCaja.txtCliente(4).Text = EsNulo(.Fields("Telefono").Value) & EsNulo(.Fields("Celular").Value)
                            frmCaja.txtCliente(5).Text = EsNulo(.Fields("Cuit").Value)
                            frmCaja.cboTipoIva(1).Tag = EsNulo(.Fields("idTipoIva").Value)
                            frmCaja.cboTipoIva(1).Text = TraerDato("TipoIva", "idTipoIva = " & EsNulo(.Fields("idTipoIva").Value) & "", "TipoIva")
                            

                    End Select
                    
                Case "frmCajaMovimientos"
                    Select Case vVieneBusqueda
                    
                        Case "Caja"
                            frmCajaMovimientos.txtCaja(0).Text = EsNulo(.Fields("idBancos").Value)
                            frmCajaMovimientos.txtCaja(1).Text = EsNulo(.Fields("Descripcion").Value)
                            
                    End Select
                
                Case "frmBancosAlta"
                        frmBancosAlta.txtAlta(3).Text = EsNulo(.Fields("CodigoCuenta").Value)
                        frmBancosAlta.txtAlta(4).Text = EsNulo(.Fields("Cuenta").Value)
                
         
                    Case "frmBancosCuentaAlta"
                        Select Case vVieneBusqueda
                            
                            Case "Banco"
                                frmBancosCuentaAlta.txtAlta(0).Text = EsNulo(.Fields("idBancos").Value)
                                frmBancosCuentaAlta.txtAlta(1).Text = EsNulo(.Fields("Descripcion").Value)

                            Case "CodigoCuenta"
                                
                                frmBancosCuentaAlta.txtAlta(4).Text = EsNulo(.Fields("CodigoCuenta").Value)
                                frmBancosCuentaAlta.txtAlta(5).Text = EsNulo(.Fields("Cuenta").Value)

                            Case "TipoCuentaBanco"
                                frmBancosCuentaAlta.txtAlta(6).Text = EsNulo(.Fields("idTipoCuentaBanco").Value)
                                frmBancosCuentaAlta.txtAlta(7).Text = EsNulo(.Fields("TipoCuentaBanco").Value)
                
                        End Select
                
                Case "frmEmpleadosAlta"
                
                    Select Case vVieneBusqueda
            
                        Case "CodigoPostal"
                            frmEmpleadosAlta.txtFicha(1).Text = EsNulo(.Fields("CodigoPostal").Value)
                            frmEmpleadosAlta.txtFicha(2).Text = EsNulo(.Fields("Localidad").Value)
                            frmEmpleadosAlta.txtFicha(3).Text = EsNulo(.Fields("Provincia").Value)

                        Case "TipoIva"
                            frmEmpleadosAlta.txtDatosComerciales(0).Text = EsNulo(.Fields("idTipoIva").Value)
                            frmEmpleadosAlta.txtDatosComerciales(1).Text = EsNulo(.Fields("TipoIva").Value)
                    
                        Case "Actividad"
                            frmEmpleadosAlta.txtDatosComerciales(3).Text = EsNulo(.Fields("idActividades").Value)
                            frmEmpleadosAlta.txtDatosComerciales(4).Text = EsNulo(.Fields("Descripcion").Value)
                    
                        Case "Lista"
                            frmEmpleadosAlta.txtDatosComerciales(5).Text = EsNulo(.Fields("idListas").Value)
                            frmEmpleadosAlta.txtDatosComerciales(6).Text = EsNulo(.Fields("Lista").Value)

                        Case "EstadoCliente"
                            frmEmpleadosAlta.txtOtrosDatos(4).Text = EsNulo(.Fields("idEstados").Value)
                            frmEmpleadosAlta.txtOtrosDatos(5).Text = EsNulo(.Fields("Estado").Value)

                    End Select
                    
                Case "frmClientesAlta"
                
                    Select Case vVieneBusqueda
            
                        Case "CodigoPostal"
                            frmClientesAlta.txtFicha(1).Text = EsNulo(.Fields("CodigoPostal").Value)
                            frmClientesAlta.txtFicha(2).Text = EsNulo(.Fields("Localidad").Value)
                            frmClientesAlta.txtFicha(3).Text = EsNulo(.Fields("Provincia").Value)
                            
                        Case "Vendedor"
                            frmClientesAlta.txtFicha(8).Text = EsNulo(.Fields("codigo").Value)
                            frmClientesAlta.txtFicha(9).Text = EsNulo(.Fields("Nombre").Value)
                        
                        Case "Reparto"
                            frmClientesAlta.txtFicha(10).Text = EsNulo(.Fields("nreparto").Value)
                            frmClientesAlta.txtFicha(11).Text = EsNulo(.Fields("descrip").Value)

                        Case "TipoDocumento"
                            frmClientesAlta.txtFicha(12).Text = EsNulo(.Fields("idTipoDocumentos").Value)
                            frmClientesAlta.txtFicha(13).Text = EsNulo(.Fields("Tipo").Value)
                            
                        Case "TipoIva"
                            frmClientesAlta.txtDatosComerciales(0).Text = EsNulo(.Fields("idTipoIva").Value)
                            frmClientesAlta.txtDatosComerciales(1).Text = EsNulo(.Fields("TipoIva").Value)
                    
                        Case "TipoCliente"
                            frmClientesAlta.txtDatosComerciales(3).Text = EsNulo(.Fields("idTipoClientes").Value)
                            frmClientesAlta.txtDatosComerciales(4).Text = EsNulo(.Fields("Descripcion").Value)
                    
                        Case "Actividad"
                            frmClientesAlta.txtDatosComerciales(5).Text = EsNulo(.Fields("idActividades").Value)
                            frmClientesAlta.txtDatosComerciales(6).Text = EsNulo(.Fields("Descripcion").Value)
                    
                        Case "Lista"
                            frmClientesAlta.txtDatosComerciales(7).Text = EsNulo(.Fields("idListas").Value)
                            frmClientesAlta.txtDatosComerciales(8).Text = EsNulo(.Fields("Lista").Value)

                        Case "EstadoCliente"
                            frmClientesAlta.txtOtrosDatos(4).Text = EsNulo(.Fields("idEstados").Value)
                            frmClientesAlta.txtOtrosDatos(5).Text = EsNulo(.Fields("Estado").Value)

                            
                        Case "CodigoArticulo"
                            frmClientesAlta.txtArticulos(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmClientesAlta.txtArticulos(1).Text = EsNulo(.Fields("Descrip").Value)
                            frmClientesAlta.txtArticulos(2).SetFocus
                    
                    End Select
            
            
                Case "frmProveedoresAlta"
                    Select Case vVieneBusqueda
            
                        Case "CodigoPostal"
                            frmProveedoresAlta.txtFicha(1).Text = EsNulo(.Fields("CodigoPostal").Value)
                            frmProveedoresAlta.txtFicha(2).Text = EsNulo(.Fields("Localidad").Value)
                            frmProveedoresAlta.txtFicha(3).Text = EsNulo(.Fields("Provincia").Value)
                            
                        Case "Vendedor"
                            frmProveedoresAlta.txtFicha(8).Text = EsNulo(.Fields("codigo").Value)
                            frmProveedoresAlta.txtFicha(9).Text = EsNulo(.Fields("Nombre").Value)
                        
                        Case "Reparto"
                            frmProveedoresAlta.txtFicha(10).Text = EsNulo(.Fields("nreparto").Value)
                            frmProveedoresAlta.txtFicha(11).Text = EsNulo(.Fields("descrip").Value)
                    
                        Case "TipoIva"
                            frmProveedoresAlta.txtDatosComerciales(0).Text = EsNulo(.Fields("idTipoIva").Value)
                            frmProveedoresAlta.txtDatosComerciales(1).Text = EsNulo(.Fields("TipoIva").Value)
                    
                        Case "TipoCliente"
                            frmProveedoresAlta.txtDatosComerciales(3).Text = EsNulo(.Fields("idTipoClientes").Value)
                            frmProveedoresAlta.txtDatosComerciales(4).Text = EsNulo(.Fields("Descripcion").Value)
                    
                        Case "Actividad"
                            frmProveedoresAlta.txtDatosComerciales(5).Text = EsNulo(.Fields("idActividades").Value)
                            frmProveedoresAlta.txtDatosComerciales(6).Text = EsNulo(.Fields("Descripcion").Value)
                    
                        Case "Lista"
                            frmProveedoresAlta.txtDatosComerciales(7).Text = EsNulo(.Fields("idListas").Value)
                            frmProveedoresAlta.txtDatosComerciales(8).Text = EsNulo(.Fields("Lista").Value)

                        Case "EstadoCliente"
                            frmProveedoresAlta.txtOtrosDatos(4).Text = EsNulo(.Fields("idEstados").Value)
                            frmProveedoresAlta.txtOtrosDatos(5).Text = EsNulo(.Fields("Estado").Value)

                    End Select
                    
                Case "frmEmpleadosAlta"
                
                Case "frmArticulosAlta"
                    Select Case vVieneBusqueda
                    
                        Case "PorcentajeIva"
                            frmArticulosAlta.txtFicha(0).Text = EsNulo(.Fields("idPorcentajeIva").Value)
                            frmArticulosAlta.txtFicha(1).Text = EsNulo(.Fields("Descripcion").Value)
                            frmArticulosAlta.txtFicha(2).Text = EsNulo(.Fields("Porcentaje").Value)
                        
                        Case "Proveedor"
                            frmArticulosAlta.txtFicha(3).Text = EsNulo(.Fields("Codigo").Value)
                            frmArticulosAlta.txtFicha(4).Text = EsNulo(.Fields("Nombre").Value)
                        
                        Case "ProveedorPrecio"
                            frmArticulosAlta.txtProveedores(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmArticulosAlta.txtProveedores(1).Text = EsNulo(.Fields("Nombre").Value)
                            
                        Case "Fabricante"
                            frmArticulosAlta.txtFicha(5).Text = EsNulo(.Fields("idFabricantes").Value)
                            frmArticulosAlta.txtFicha(6).Text = EsNulo(.Fields("Nombre").Value)
                        
                        Case "SubRubro"
                            frmArticulosAlta.txtAlta(2).Text = EsNulo(.Fields("idSubRubros").Value)
                            frmArticulosAlta.txtAlta(3).Text = EsNulo(.Fields("SubRubro").Value)
                        
                        Case "Rubro"
                            frmArticulosAlta.txtAlta(4).Text = EsNulo(.Fields("idRubros").Value)
                            frmArticulosAlta.txtAlta(5).Text = EsNulo(.Fields("Rubro").Value)
                    End Select

                Case "frmImprimir"
                
                    Select Case vVieneBusqueda
                    
                        Case "CodigoCliente", "CodigoClienteD", "CodigoClienteH"
                            If vVieneBusqueda = "CodigoClienteD" Then
                                frmImprimir.txtIntervalos(0).Text = EsNulo(.Fields("Codigo").Value)
                            Else
                                frmImprimir.txtIntervalos(1).Text = EsNulo(.Fields("Codigo").Value)
                            End If
                        
                        Case "Proveedor", "ProveedorD", "ProveedorH"
                            If vVieneBusqueda = "ProveedorD" Then
                                frmImprimir.txtIntervalos(3).Text = EsNulo(.Fields("Codigo").Value)
                            Else
                                frmImprimir.txtIntervalos(4).Text = EsNulo(.Fields("Nombre").Value)
                            End If
                            
                        Case "Fabricante", "FabricanteD", "FabricanteH"
                            If vVieneBusqueda = "FabricanteD" Then
                                frmImprimir.txtIntervalos(5).Text = EsNulo(.Fields("idFabricantes").Value)
                            Else
                                frmImprimir.txtIntervalos(6).Text = EsNulo(.Fields("Nombre").Value)
                            End If
                        
                        Case "CodigoPostal", "CodigoPostalD", "CodigoPostalH"
                            Select Case vVieneBusqueda
                            
                                Case "CodigoPostal"
                                
                                Case "CodigoPostalD"
                                    frmImprimir.txtIntervalos(6).Text = EsNulo(.Fields("CodigoPostal").Value)
                                    frmImprimir.txtIntervalos(7).Text = EsNulo(TraerDato("Localidades", "CodigoPostal = '" & .Fields("CodigoPostal").Value & "'", "Localidad"))
                                Case "CodigoPostalH"
                                    frmImprimir.txtIntervalos(8).Text = EsNulo(.Fields("CodigoPostal").Value)
                                    frmImprimir.txtIntervalos(9).Text = EsNulo(TraerDato("Localidades", "CodigoPostal = '" & .Fields("CodigoPostal").Value & "'", "Localidad"))
                            
                            End Select
                        
                        Case "EstadoCliente", "EstadoClienteD", "EstadoClienteH"
                            Select Case vVieneBusqueda
                                
                                Case "EstadoCliente"
                                    frmImprimir.txtIntervalos(10).Text = EsNulo(.Fields("idEstados").Value)
                                    frmImprimir.txtIntervalos(11).Text = EsNulo(TraerDato("Estados", "idEstados = '" & .Fields("idEstados").Value & "'", "Estado"))
                                Case "EstadoClienteD"
                                    frmImprimir.txtIntervalos(5).Text = EsNulo(.Fields("idFabricantes").Value)
                                Case "EstadoClienteH"
                                    frmImprimir.txtIntervalos(5).Text = EsNulo(.Fields("idFabricantes").Value)
                            End Select
                            
                        Case "TipoCliente", "TipoClienteD", "TipoClienteH"
                            Select Case vVieneBusqueda
                                
                                Case "TipoCliente"
                                    'frmImprimir.txtIntervalos(0).Text = EsNulo(.Fields("idTipoClientes").Value)
                                    'frmImprimir.txtIntervalos(0).Text = EsNulo(TraerDato("TipoClientes", "idTipoClientes = '" & .Fields("idTipoClientes").Value & "'", "TipoCliente"))
                                
                                Case "TipoClienteD"
                                    frmImprimir.txtIntervalos(12).Text = EsNulo(.Fields("idTipoClientes").Value)
                                    frmImprimir.txtIntervalos(13).Text = EsNulo(TraerDato("TipoClientes", "idTipoClientes = '" & .Fields("idTipoClientes").Value & "'", "Descripcion"))
                                Case "TipoClienteH"
                                    frmImprimir.txtIntervalos(14).Text = EsNulo(.Fields("idTipoClientes").Value)
                                    frmImprimir.txtIntervalos(15).Text = EsNulo(TraerDato("TipoClientes", "idTipoClientes = '" & .Fields("idTipoClientes").Value & "'", "Descripcion"))
                            End Select
                    
                        Case "Actividad", "ActividadD", "ActividadH"
                            Select Case vVieneBusqueda
                                
                                Case "Actividad"
                                    'frmImprimir.txtIntervalos(16).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(17).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                Case "ActividadD"
                                    frmImprimir.txtIntervalos(16).Text = EsNulo(.Fields("idActividades").Value)
                                    frmImprimir.txtIntervalos(17).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                
                                Case "ActividadH"
                                    frmImprimir.txtIntervalos(18).Text = EsNulo(.Fields("idActividades").Value)
                                    frmImprimir.txtIntervalos(19).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                            End Select
                        
                        Case "Vendedor", "VendedorD", "VendedorH"
                            Select Case vVieneBusqueda
                                
                                Case "Vendedor"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                
                                Case "VendedorD"
                                    frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("Codigo").Value)
                                    frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Empleados", "Codigo = '" & .Fields("Codigo").Value & "'", "Nombre"))
                                
                                Case "VendedorH"
                                    frmImprimir.txtIntervalos(22).Text = EsNulo(.Fields("Codigo").Value)
                                    frmImprimir.txtIntervalos(23).Text = EsNulo(TraerDato("Empleados", "Codigo = '" & .Fields("Codigo").Value & "'", "Nombre"))
                            End Select
                        
                        Case "CodigoArticulo", "CodigoArticuloD", "CodigoArticuloH"
                            Select Case vVieneBusqueda
                                
                                Case "CodigoArticulo"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                
                                Case "CodigoArticuloD"
                                    frmImprimir.txtArticulos(0).Text = EsNulo(.Fields("Codigo").Value)
                                
                                Case "CodigoArticuloH"
                                    frmImprimir.txtArticulos(1).Text = EsNulo(.Fields("Codigo").Value)
                            End Select
                        
                        Case "CodigoProveedor", "CodigoProveedorD", "CodigoProveedorH"
                            Select Case vVieneBusqueda
                                
                                Case "CodigoProveedor"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                
                                Case "CodigoProveedorD"
                                    frmImprimir.txtArticulos(4).Text = EsNulo(.Fields("Codigo").Value)
                                    frmImprimir.txtArticulos(5).Text = EsNulo(.Fields("Nombre").Value)
                                
                                Case "CodigoProveedorH"
                                    frmImprimir.txtArticulos(6).Text = EsNulo(.Fields("Codigo").Value)
                                    frmImprimir.txtArticulos(7).Text = EsNulo(.Fields("Nombre").Value)
                            End Select
                    
                        Case "Rubro", "RubroD", "RubroH"
                            Select Case vVieneBusqueda
                                
                                Case "Rubro"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                
                                Case "RubroD"
                                    frmImprimir.txtArticulos(8).Text = EsNulo(.Fields("idRubros").Value)
                                    frmImprimir.txtArticulos(9).Text = EsNulo(.Fields("Rubro").Value)
                                
                                Case "RubroH"
                                    frmImprimir.txtArticulos(10).Text = EsNulo(.Fields("idRubros").Value)
                                    frmImprimir.txtArticulos(11).Text = EsNulo(.Fields("Rubro").Value)
                            End Select
                            
                        Case "SubRubro", "SubRubroD", "SubRubroH"
                            Select Case vVieneBusqueda
                                
                                Case "SubRubro"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                Case "SubRubroD"
                                    frmImprimir.txtArticulos(12).Text = EsNulo(.Fields("idSubRubros").Value)
                                    frmImprimir.txtArticulos(13).Text = EsNulo(.Fields("SubRubro").Value)
                                
                                Case "SubRubroH"
                                    frmImprimir.txtArticulos(14).Text = EsNulo(.Fields("idSubRubros").Value)
                                    frmImprimir.txtArticulos(15).Text = EsNulo(.Fields("SubRubro").Value)
                            End Select
                        
                        Case "PorcentajeIva", "PorcentajeIvaD", "PorcentajeIvaH"
                            Select Case vVieneBusqueda
                                
                                Case "PorcentajeIva"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                
                                Case "PorcentajeIvaD"
                                    frmImprimir.txtArticulos(16).Text = EsNulo(.Fields("idPorcentajeIva").Value)
                                    frmImprimir.txtArticulos(17).Text = EsNulo(.Fields("Porcentaje").Value)
                                
                                Case "PorcentajeIvaH"
                                    frmImprimir.txtArticulos(18).Text = EsNulo(.Fields("idPorcentajeIva").Value)
                                    frmImprimir.txtArticulos(19).Text = EsNulo(.Fields("Porcentaje").Value)
                            End Select
                    
                        Case "CodigoCuenta", "CodigoCuentaD", "CodigoCuentaH"
                            Select Case vVieneBusqueda
                                
                                Case "CodigoCuenta"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                Case "CodigoCuentaD"
                                    frmImprimir.txtContabilidad(0).Text = EsNulo(.Fields("CodigoCuenta").Value)
                                    frmImprimir.txtContabilidad(1).Text = EsNulo(.Fields("Cuenta").Value)
                                
                                Case "CodigoCuentaH"
                                    frmImprimir.txtContabilidad(2).Text = EsNulo(.Fields("CodigoCuenta").Value)
                                    frmImprimir.txtContabilidad(3).Text = EsNulo(.Fields("Cuenta").Value)
                            End Select
                        
                        Case "Banco", "BancoD", "BancoH"
                            Select Case vVieneBusqueda
                                
                                Case "Banco"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                Case "BancoD"
                                    frmImprimir.txtBancoCajaDetalle(0).Text = EsNulo(.Fields("idBancos").Value)
                                    frmImprimir.txtBancoCajaDetalle(1).Text = EsNulo(.Fields("Descripcion").Value)
                                
                                Case "BancoH"
                                    frmImprimir.txtBancoCajaDetalle(2).Text = EsNulo(.Fields("idBancos").Value)
                                    frmImprimir.txtBancoCajaDetalle(3).Text = EsNulo(.Fields("Descripcion").Value)
                            End Select
                    
                        Case "BancoCuenta", "BancoCuentaD", "BancoCuentaH"
                            Select Case vVieneBusqueda
                                
                                Case "BancoCuenta"
                                    'frmImprimir.txtIntervalos(20).Text = EsNulo(.Fields("idActividades").Value)
                                    'frmImprimir.txtIntervalos(21).Text = EsNulo(TraerDato("Actividades", "idActividades = '" & .Fields("idActividades").Value & "'", "Descripcion"))
                                    
                                Case "BancoCuentaD"
                                    frmImprimir.txtBancoCajaDetalle(4).Text = EsNulo(.Fields("idBancosCuentas").Value)
                                    frmImprimir.txtBancoCajaDetalle(5).Text = EsNulo(.Fields("Cuenta").Value)
                                
                                Case "BancoCuentaH"
                                    frmImprimir.txtBancoCajaDetalle(6).Text = EsNulo(.Fields("idTipoValor").Value)
                                    frmImprimir.txtBancoCajaDetalle(7).Text = EsNulo(.Fields("TipoValor").Value)
                            End Select
                    
                                        
                    End Select
                    
                Case "frmRemito"
                            
 
                            
                    Select Case vVieneBusqueda
                        Case "Vendedor"
                            frmRemito.txtEmpleados(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmRemito.txtEmpleados(1).Text = EsNulo(.Fields("Nombre").Value)
                    
                        Case "TipoMovimientos"
                            frmRemito.txtTipoMovimiento(0).Text = EsNulo(.Fields("Codigo").Value)
                            frmRemito.txtTipoMovimiento(1).Text = EsNulo(.Fields("TipoMovimiento").Value)
                            frmRemito.txtIB(0).SetFocus
                        
                        Case "CajaBanco"
                            frmRemito.txtCaja(0).Text = EsNulo(.Fields("idBancos").Value)
                            frmRemito.txtCaja(1).Text = EsNulo(.Fields("Descripcion").Value)
                            frmRemito.pbCarga(3).SetFocus
                            
                        Case "BancoCuenta"
                            frmRemito.txtCaja(2).Text = EsNulo(.Fields("idBancosCuentas").Value)
                            frmRemito.txtCaja(3).Text = EsNulo(.Fields("Cuenta").Value)
                            frmRemito.pbCarga(4).SetFocus
                            
                        Case "TipoValor"
                            frmRemito.txtCaja(4).Text = EsNulo(.Fields("idBancosCuentas").Value)
                            frmRemito.txtCaja(5).Text = EsNulo(.Fields("Cuenta").Value)
                                     
                    End Select
            
            End Select
        
        End If
    
    End With

    Unload Me

If Err Then GrabarLog "dgBusqueda_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub dgBusqueda_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    Call OrdenarDataGrid(ColIndex, rsBusqueda, dgBusqueda)

If Err Then GrabarLog "dgBusqueda_HeadClick", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Form_Load()
On Error Resume Next
    Me.Show
    SeleccionarModelo
    txtBusqueda.SetFocus
    
    'MenuPanelConfigurar
    
    
    Call CentrarFormulario(Me)
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub CreatePane(Title As String, Direction As DockingDirection, PaneIndex As Integer)

Dim Pane As XtremeDockingPane.Pane
Set Pane = DockingPane1.CreatePane(PaneIndex, 100, 200, Direction)
    Pane.Title = Title

Set arrPanes(PaneIndex) = New frmTaskPanel
    Pane.Handle = arrPanes(PaneIndex).hWnd
End Sub

Private Sub MenuPanelConfigurar()
DockingPane1.VisualTheme = ThemeOffice2007
DockingPane1.Options.DefaultPaneOptions = PaneNoCloseable Or PaneNoDockable Or PaneNoFloatable

Call CreatePane("Pane A", DockLeftOf, 1)
End Sub
Private Sub SeleccionarModelo()
' paso 1: indicar de que botòn viene el llamado. Se chequea por el .tag del botón
On Error Resume Next
    
    Select Case vVieneBusqueda ' Alfredo: acá controla si coincide lo que  pusiste en el tag del botón
    ' Alfredo: si hay una tabla nueva, tenés q agregar un case con los datos correspondientes
    
        Case "CodigoPostal", "CodigoPostalD", "CodigoPostalH"
            vGrabarTabla = "Localidades"
        
        Case "Vendedor", "VendedorD", "VendedorH"
            vGrabarTabla = "Empleados"
            
        Case "Reparto", "RepartoD", "RepartoH"
            vGrabarTabla = "CliReparto"
            
        Case "TipoIva", "TipoIvaD", "TipoIvaH"
            vGrabarTabla = "TipoIva"
            
        Case "TipoCliente", "TipoClienteD", "TipoClienteH"
            vGrabarTabla = "TipoClientes"
            
        Case "Actividad", "ActividadD", "ActividadH"
            vGrabarTabla = "Actividades"
        
        Case "Lista", "ListaD", "ListaH"
            vGrabarTabla = "Listas"
            
        Case "EstadoCliente", "EstadoClienteD", "EstadoClienteH"
            vGrabarTabla = "Estados"
            
        Case "Rubro", "RubroD", "RubroH"
            vGrabarTabla = "Rubros"
    
        Case "SubRubro", "SubRubroD", "SubRubroH"
            vGrabarTabla = "SubRubros"
            
        Case "Proveedor", "ProveedorD", "ProveedorH", "CodigoProveedor", "CodigoProveedorD", "CodigoProveedorH", "ProveedorPrecio"
            vGrabarTabla = "Proveedores"
        
        Case "Fabricante", "FabricanteD", "FabricanteH", "CodigoFabricante", "CodigoFabricanteD", "CodigoFabricanteH"
            vGrabarTabla = "Fabricantes"
    
        Case "PorcentajeIva", "PorcentajeIvaD", "PorcentajeIvaH"
            vGrabarTabla = "PorcentajeIva"
    
        Case "CodigoCliente", "CodigoClienteD", "CodigoClienteH", "cliente"
            vGrabarTabla = "Clientes"
        
        Case "CodigoArticulo", "CodigoArticuloD", "CodigoArticuloH"
            vGrabarTabla = "Articulos"
    
        Case "Cotizacion"
            vGrabarTabla = "Cotizaciones"
    
        Case "CodigoCuenta", "CodigoCuentaD", "CodigoCuentaH"
            vGrabarTabla = "Cuentas"
    
        Case "Banco", "BancoD", "BancoH"
            vGrabarTabla = "Bancos WHERE (not EsCaja = 'S')"
    
        Case "Caja", "caja-importe-cobro", "compra-caja"
            vGrabarTabla = "Bancos WHERE (EsCaja = 'S')"
            
        Case "CajaBanco"
            vGrabarTabla = "Bancos"

        Case "TipoCuentaBanco", "TipoCuentaBancoD", "TipoCuentaBancoH"
            vGrabarTabla = "Tipocuentabanco"
    
        Case "BancoCuenta", "BancoCuentaD", "BancoCuentaH"
        
            Select Case vVuelveBusqueda
                Case "frmPagos"
                    vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmPagos.txtBancoCheque(0).Text) & "'"
                    
                Case "frmIngresosEgresos"
                    vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmIngresosEgresos.txtAlta(6).Text) & "'"
                
                Case "frmCompras"
                    vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmCompras.txtCaja(0).Text) & "'"
                
                Case "frmImprimir"
                    Select Case vVieneBusqueda
                    
                        Case "BancoCuentaD"
                            vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmImprimir.txtBancoCajaDetalle(0).Text) & "'"
                        
                        Case "BancoCuentaH"
                            vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmImprimir.txtBancoCajaDetalle(2).Text) & "'"
                    End Select
                    
                Case "frmCobros"
                    vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmCobros.txtBancoCheque(0).Text) & "'"
                
                Case "frmChequesAlta"
                    vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmChequesAlta.txtFicha(9).Text) & "'"
                
                Case "frmBancosMovimientos"
                    vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmBancosMovimientos.txtBanco(0).Text) & "'"
                
                Case Else
                    vGrabarTabla = "Bancoscuentas"
            
            End Select
                    
        Case "BancoDeposito"
            vGrabarTabla = "Bancos"
        
        Case "BancoCuentaDeposito"
            
            Select Case vVuelveBusqueda
                            
                Case "frmCobros"
                    vGrabarTabla = "Bancoscuentas WHERE idBancos = '" & Trim(frmCobros.txtDepositoBanco(0).Text) & "'"
                
                Case Else
                    vGrabarTabla = "Bancoscuentas"
            
            End Select
        
        Case "Mozo"
            vGrabarTabla = "Mozos"

        Case "TipoMovimientos"
            vGrabarTabla = "TipoMovimientos"
            
        Case "TipoValor"
            vGrabarTabla = "TipoValor"
        
        Case "TipoDocumento"
            vGrabarTabla = "TipoDocumentos"
    
        Case "TipoMovimientosBanco"
            vGrabarTabla = "TipoMovimientos"
    
        Case "EstadoCheque"
            vGrabarTabla = "EstadoCheque"
    
        Case "BancoCheque"
            vGrabarTabla = "BancoCheque"
            
    End Select
    
    CargarRegistros
            
If Err Then GrabarLog "SeleccionarModelo", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    vVieneBusqueda = ""
    vVuelveBusqueda = ""
    
    Set dgBusqueda.DataSource = Nothing

    If rsBusqueda.State = 1 Then
        rsBusqueda.Close
        Set rsBusqueda = Nothing
    End If
    
    Unload Me
If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Botonera_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    
    Select Case Button.Index
                          
        Case 1
            Nuevo
        
        Case 2
            Modificar
        
        Case 3
            BorrarRegistro
        
        Case 4
            'Separador
        
        Case 5
            Buscar
        
        Case 6
            Seleccionar
        
        Case 7
            'Separador
        
        Case 8
            VerAyuda (Me.Name)
        
        Case 9
            Unload Me
        
        
    End Select

If Err Then GrabarLog "Botonera_ButtonClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Function CargarRegistros() As Boolean

' Alfredo: en el caso de que estes cargando una tabla nueva vas a tener que crear un case mas
On Error Resume Next
    
    Set rsBusqueda = New ADODB.Recordset
    Dim sqlBusqueda As String, i As Integer

    sqlBusqueda = "SELECT * FROM " & vGrabarTabla & " ORDER BY 1"
    
    With rsBusqueda
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        Call .Open(sqlBusqueda, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        CargarRegistros = Not .EOF
        
    End With

    Select Case vVieneBusqueda
        
        Case "TipoDocumento", "TipoDocumentoD", "TipoDocumentoH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                

                .Columns(0).Width = 1500
                .Columns(0).Caption = "Tipo"
                .Columns(1).Width = 6100
                .Columns(1).Caption = "Descripcion"
                .Columns(2).Width = 0

            
            End With
        Case "CodigoPostal", "CodigoPostalD", "CodigoPostalH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 0
                .Columns(1).Width = 1500
                .Columns(1).Caption = "Codigo Postal"
                .Columns(2).Width = 3000
                .Columns(2).Caption = "Localidad"
                .Columns(3).Width = 3000
                .Columns(3).Caption = "Provincia"
            
            End With
        
        Case "Vendedor", "VendedorD", "VendedorH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 0
                
                .Columns(1).Width = 1000
                .Columns(1).Caption = "Codigo"
                .Columns(2).Width = 0
                .Columns(3).Width = 6600
                .Columns(3).Caption = "Nombre del Vendedor"
                
                For i = 4 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With

        Case "Reparto", "RepartoD", "RepartoH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 0
                .Columns(1).Width = 1000
                .Columns(1).Caption = "Codigo"
                .Columns(2).Width = 6600
                .Columns(2).Caption = "Reparto"
                
            End With
            
        Case "TipoIva", "TipoIvaD", "TipoIvaH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Tipo Iva"
            End With

        Case "TipoCliente", "TipoClienteD", "TipoClienteH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Tipo de Cliente"
            End With
        
        Case "Actividad", "ActividadD", "ActividadH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 5000
                .Columns(1).Caption = "Actividades"
            End With
        
        Case "Lista", "ListaD", "ListaH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Lista de Precio"
            End With
            
        Case "EstadoCliente", "EstadoClienteD", "EstadoClienteH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Estado"
            End With
            
        Case "Rubro", "RubroD", "RubroH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Descripcion de Rubro"
            End With
        
        Case "SubRubro", "SubRubroD", "SubRubroH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 0
                .Columns(2).Width = 6600
                .Columns(2).Caption = "Descripcion de SubRubro"
            End With
            
        Case "Fabricante", "FabricanteD", "FabricanteH", "CodigoFabricante", "CodigoFabricanteD", "CodigoFabricanteH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Fabricante"
            
                For i = 2 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
                
                                    
        Case "PorcentajeIva", "PorcentajeIvaD", "PorcentajeIvaH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 5000
                .Columns(1).Caption = "Descripcion de Rubro"
                .Columns(2).Width = 1500
                .Columns(2).Caption = "%"
            End With
    
        Case "CodigoCliente", "CodigoClienteD", "CodigoClienteH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                For i = 0 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
                .Columns(1).Width = 1000
                .Columns(1).Caption = "Codigo"
                .Columns(3).Width = 6600
                .Columns(3).Caption = "Nombre de Cliente"
            End With
        
        Case "Articulos", "CodigoArticuloD", "CodigoArticuloH", "CodigoArticulo"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 0
                .Columns(1).Width = 1000
                .Columns(1).Caption = "Codigo"
                .Columns(2).Width = 0
                .Columns(3).Width = 0
                .Columns(4).Width = 6600
                .Columns(4).Caption = "Descripcion"
                
                For i = 5 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
            
        Case "Proveedor", "ProveedorD", "ProveedorH", "CodigoProveedor", "CodigoProveedorD", "CodigoProveedorH", "ProveedorPrecio"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 0
                .Columns(1).Width = 1000
                .Columns(1).Caption = "Codigo"
                .Columns(2).Width = 0
                .Columns(3).Width = 6600
                .Columns(3).Caption = "Proveedor"
                
                For i = 4 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
            
        Case "CodigoCuenta", "CodigoCuentaD", "CodigoCuentaH"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 0
                .Columns(1).Width = 2000
                .Columns(1).Caption = "Codigo"
                .Columns(2).Width = 6600
                .Columns(2).Caption = "Cuenta"
                
                For i = 3 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
    
        Case "Caja", "CajaD", "CajaH", "Banco", "BancoD", "BancoH", "CajaBanco", "CajaBancoD", "CajaBancoH", "caja-importe-cobro", "compra-caja"
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6000
                .Columns(1).Caption = "Caja/Banco"
                .Columns(2).Width = 600
                .Columns(2).Caption = "Es Caja"
                
                For i = 3 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
            
        Case "TipoCuentaBanco", "TipoCuentaBancoD", "TipoCuentaBancoH"
    
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Banco"
                
                For i = 2 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
    
        Case "BancoCuenta", "BancoCuentaD", "BancoCuentaH", "BancoCuentaDeposito", "BancoCuentaDepositoD", "BancoCuentaDepositoH"
    
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 500
                .Columns(0).Caption = "ID"
                .Columns(1).Width = 0
                .Columns(2).Width = 1500
                .Columns(2).Caption = "Codigo"
                .Columns(3).Width = 5500
                .Columns(3).Caption = "Banco"
                
                For i = 4 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
    
        Case "Mozo", "MozoD", "MozoH"
    
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 1000
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Nombre"
                
                For i = 2 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
    
        Case "TipoMovimientos", "TipoMovimientosD", "TipoMovimientosH", "TipoMovimientosBanco"
    
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 0
                .Columns(1).Width = 750
                .Columns(1).Caption = "Codigo"
                .Columns(2).Width = 6600
                .Columns(2).Caption = "Descripcion"
                
                For i = 3 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
        
        Case "TipoValor", "TipoValorD", "TipoValorH"
    
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 750
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Tipo de Valor"
                
                For i = 3 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
    
        Case "EstadoCheque", "EstadoChequeD", "EstadoChequeH"
    
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 750
                .Columns(0).Caption = "ID Estado Cheque"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Descripcion"
                
                For i = 2 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
        
        Case "BancoCheque", "BancoChequeD", "BancoChequeH", "BancoDeposito", "BancoDepositoD", "BancoDepositoH"
    
            With dgBusqueda
                Set .DataSource = rsBusqueda
                .HeadLines = 2
                
                .Columns(0).Width = 750
                .Columns(0).Caption = "Codigo"
                .Columns(1).Width = 6600
                .Columns(1).Caption = "Nombre"
                
                For i = 2 To .Columns.Count - 1
                    .Columns(i).Width = 0
                Next
            End With
    End Select

If Err Then GrabarLog "CargarRegistros", Err.Number & " " & Err.Description, Me.Caption
End Function
Public Sub txtBusqueda_Change()
On Error Resume Next

    With rsBusqueda
        If Trim(txtBusqueda.Text) = "" Then
            SeleccionarModelo
            Exit Sub
        End If
    
        Select Case vVieneBusqueda
    
            Case "CodigoPostal"
                .Filter = "Localidad LIKE '%" & Trim(txtBusqueda.Text) & "%' or CodigoPostal LIKE '%" & Trim(txtBusqueda.Text) & "%'"
                
            Case "Vendedor", "Proveedor", "Fabricante", "Cliente"
                .Filter = "Nombre LIKE '%" & Trim(txtBusqueda.Text) & "%'"
                
            Case "Reparto"
                .Filter = "Descrip LIKE '%" & Trim(txtBusqueda.Text) & "%'"
            
            Case "TipoIva"
                .Filter = "TipoIva LIKE '%" & Trim(txtBusqueda.Text) & "%'"
            
            Case "TipoCliente", "Actividad", "PorcentajeIva"
                .Filter = "Descripcion LIKE '%" & Trim(txtBusqueda.Text) & "%'"
            
            Case "Lista"
                .Filter = "Lista = " & Val(txtBusqueda.Text) & ""
            
            Case "EstadoCliente"
                .Filter = "Estado LIKE '%" & Trim(txtBusqueda.Text) & "%'"
            
            Case "Rubro"
                .Filter = "Rubro LIKE '%" & Trim(txtBusqueda.Text) & "%'"
                
            Case "SubRubro"
                .Filter = "SubRubro LIKE '%" & Trim(txtBusqueda.Text) & "%'"
                                
            Case "TipoMovimientos", "TipoMovimientosBanco"
                .Filter = "TipoMovimiento LIKE '%" & Trim(txtBusqueda.Text) & "%' or Codigo Like '%" & Trim(txtBusqueda.Text) & "%'"
                
            Case "Cotizacion"
                '.filter = "Descrip LIKE '%" & Trim(txtBusqueda.Text) & "%'"
            Case "TipoCuentaBanco"
                .Filter = "TipoCuentaBanco LIKE '%" & Trim(txtBusqueda.Text) & "%'"
        
            Case "CodigoCuenta"
                .Filter = "Cuenta LIKE '%" & Trim(txtBusqueda.Text) & "%' or CodigoCuenta LIKE '%" & Trim(txtBusqueda.Text) & "%'"
            
            Case "Mozo"
                .Filter = "Mozo LIKE '%" & Trim(txtBusqueda.Text) & "%'"
                
            Case "CodigoCliente"
                .Filter = "Nombre LIKE '%" & Trim(txtBusqueda.Text) & "%'"
                    
            Case "TipoValor"
                .Filter = "TipoValor LIKE '%" & Trim(txtBusqueda.Text) & "%'"
            
            Case "TipoDocumento"
                .Filter = "TipoDocumento LIKE '%" & Trim(txtBusqueda.Text) & "%'"
            
            Case "Bancos", "CajaBanco", "Banco", "BancoD", "BancoH", "BancoDeposito", "BancoDepositoD", "BancoDepositoH"
                '.Filter = "Descripcion LIKE '%" & Trim(txtBusqueda.Text) & "%' or idbancos like '%" & (txtBusqueda.Text) & "%'"
               
               If Val(Me.txtBusqueda.Text) > 0 Then
               
                    .Filter = " idbancos like '" & txtBusqueda.Text + "%'"
                
               Else
               
                    .Filter = "Descripcion LIKE '%" & Trim(txtBusqueda.Text) & "%'"
                End If
            
            Case "EstadoCheque"
                .Filter = "Descripcion LIKE '%" & Trim(txtBusqueda.Text) & "%'"
        
            Case "BancoCheque"
                .Filter = "Nombre LIKE '%" & Trim(txtBusqueda.Text) & "%'"
        
        End Select

    End With
    
If Err Then GrabarLog "txtBusqueda_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub txtBusqueda_KeyPress(Keyascii As Integer)
On Error Resume Next

    If Keyascii = 13 Then
        dgBusqueda_DblClick
    End If

If Err Then GrabarLog "txtBusqueda_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
