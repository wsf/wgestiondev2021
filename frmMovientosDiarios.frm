VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmMovientosDiarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Actividades Diarios"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6765
   Begin VB.Frame fraGeneral 
      Caption         =   "Generar Listado"
      Height          =   1785
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6735
      Begin VB.TextBox txtNInterno 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   1950
         TabIndex        =   11
         Top             =   1350
         Width           =   1215
      End
      Begin VB.TextBox txtNInterno 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   1950
         TabIndex        =   10
         Top             =   990
         Width           =   1215
      End
      Begin VB.CheckBox chkFechas 
         Caption         =   "Todas las Fechas"
         Height          =   255
         Left            =   3570
         TabIndex        =   5
         Top             =   330
         Value           =   1  'Checked
         Width           =   2835
      End
      Begin VB.TextBox txtNAsiento 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   0
         Top             =   210
         Width           =   1215
      End
      Begin VB.TextBox txtNAsiento 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   1
         Top             =   570
         Width           =   1215
      End
      Begin MSComCtl2.DTPicker dtpCuentas 
         Height          =   315
         Index           =   0
         Left            =   4725
         TabIndex        =   6
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   25821185
         CurrentDate     =   39455
      End
      Begin MSComCtl2.DTPicker dtpCuentas 
         Height          =   315
         Index           =   1
         Left            =   4725
         TabIndex        =   7
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   25821185
         CurrentDate     =   39455
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Interno Desde :"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   13
         Top             =   1020
         Width           =   1710
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Interno Hasta :"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   1380
         Width           =   1680
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "> Fecha Desde:"
         Height          =   195
         Index           =   2
         Left            =   3255
         TabIndex        =   9
         Top             =   765
         Width           =   1380
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "> Fecha Hasta:"
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   8
         Top             =   1125
         Width           =   1410
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Asiento Hasta :"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   615
         Width           =   1620
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Asiento Desde :"
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   3
         Top             =   255
         Width           =   1800
      End
   End
   Begin XtremeSuiteControls.PushButton cmdEjecutar 
      Height          =   375
      Left            =   30
      TabIndex        =   14
      Top             =   1800
      Width           =   6705
      _Version        =   851968
      _ExtentX        =   11827
      _ExtentY        =   661
      _StockProps     =   79
      Caption         =   "Filtrar Datos"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmMovientosDiarios.frx":0000
   End
End
Attribute VB_Name = "frmMovientosDiarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vsql As String
Public vGidasientosdesde, vGidasientoshasta As Long

Private Sub chkFechas_Click()
On Error Resume Next

    dtpCuentas(0).Enabled = Not CBool(chkFechas.Value)
    dtpCuentas(1).Enabled = Not CBool(chkFechas.Value)

If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub cmdEjecutar_Click()
On Error Resume Next
    
    vsql = ""
    If Not (Val(txtNAsiento(0).Text) = 0) And Not (Val(txtNAsiento(1).Text) = 0) Then
        vsql = vsql + " AND (Numero >= " & txtNAsiento(0).Text & " AND Numero <= " & txtNAsiento(1).Text & ")"
    End If
    If Not (Val(txtNInterno(0).Text) = 0) And Not (Val(txtNInterno(1).Text) = 0) Then
        vsql = vsql + " AND (NroInterno >= " & txtNInterno(0).Text & " AND NroInterno <= " & txtNInterno(1).Text & ")"
    End If
    If (chkFechas.Value = 0) Then
        vsql = vsql + " AND (Fecha >= '" & strfechaMySQL(dtpCuentas(0).Value) & "' AND Fecha <= '" & strfechaMySQL(dtpCuentas(1).Value) & "')"
    End If
    
    If Me.vGidasientosdesde > 0 Or vGidasientoshasta > 0 Then
        vsql = " AND idAsientos > " & Str(vGidasientosdesde) + " and idAsientos <= " & Str(vGidasientoshasta)
    End If
    
    MostrarReporte
    
If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub MostrarReporte()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento

    MsgBox "Prepare la Impresora!!!", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsDiario
        If .State = 1 Then .Close
        
        .Source = "SHAPE {SELECT * FROM Asientos WHERE 1=1 " + vsql + " order by fecha,1} AS Diario APPEND ({SELECT * FROM AsientosDetalle INNER JOIN Cuentas ON AsientosDetalle.CodigoCuenta = Cuentas.CodigoCuenta where debe+haber > 0}  AS AsientosDetalle RELATE 'Numero' TO 'Numero') AS AsientosDetalle"
        
       '  .Source = "SHAPE {SELECT * FROM Asientos }  AS Diario APPEND ({SELECT * FROM AsientosDetalle INNER JOIN Cuentas ON AsientosDetalle.CodigoCuenta = Cuentas.CodigoCuenta }  AS AsientosDetalle RELATE 'Numero' TO 'Numero') AS AsientosDetalle"
        
        If .State = 0 Then .Open
        .Close
        .Open
        
        If .RecordCount = 0 Then Exit Sub
    End With
    
    With drLDiario
        .Sections(2).Controls("lblTitulo").Caption = "[ Movimientos diarios desde el número:  " & txtNInterno(0).Text & " Hasta : " & txtNInterno(1).Text & " ]"
        .Sections(2).Controls("snombre").Caption = vDatosEmpresa.Nombre
        .Sections(2).Controls("sdirtel").Caption = vDatosEmpresa.Direccion & "  /  " & vDatosEmpresa.Telefono
        .Sections(2).Controls("slocalidad").Caption = vDatosEmpresa.Localidad
        .Sections(2).Controls("semail").Caption = vDatosEmpresa.Email
        .Show
    End With
    
If Err Then GrabarLog "MostrarReporte", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next
    
'    With Me
'        .Left = 0
'        .Top = 0
'        .Width = 6855
'        .Height = 2970
'        .KeyPreview = True
'    End With



Me.dtpCuentas(0) = Date
Me.dtpCuentas(1) = Date




    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000


If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub

