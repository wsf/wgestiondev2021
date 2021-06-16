VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmBancosAlta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Bancos"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   7395
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   225
      Left            =   -30
      TabIndex        =   13
      Top             =   480
      Width           =   7605
      _Version        =   851968
      _ExtentX        =   13414
      _ExtentY        =   397
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   555
      Left            =   -30
      TabIndex        =   11
      Top             =   -60
      Width           =   7425
      _Version        =   851968
      _ExtentX        =   13097
      _ExtentY        =   979
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   2
      Begin XtremeSuiteControls.PushButton PbAcciones 
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   150
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Grabar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmBancosAlta.frx":0000
      End
   End
   Begin XtremeSuiteControls.TabControl TabAlta 
      Height          =   3015
      Left            =   30
      TabIndex        =   6
      Top             =   780
      Width           =   7335
      _Version        =   851968
      _ExtentX        =   12938
      _ExtentY        =   5318
      _StockProps     =   68
      Color           =   8
      ItemCount       =   1
      Item(0).Caption =   "Ficha"
      Item(0).ControlCount=   14
      Item(0).Control(0)=   "txtAlta(0)"
      Item(0).Control(1)=   "txtAlta(1)"
      Item(0).Control(2)=   "lblAlta(2)"
      Item(0).Control(3)=   "lblAlta(1)"
      Item(0).Control(4)=   "lblAlta(0)"
      Item(0).Control(5)=   "txtAlta(3)"
      Item(0).Control(6)=   "txtAlta(4)"
      Item(0).Control(7)=   "lblAlta(3)"
      Item(0).Control(8)=   "pbCarga(0)"
      Item(0).Control(9)=   "txtAlta(2)"
      Item(0).Control(10)=   "vtipocaja"
      Item(0).Control(11)=   "vlabel"
      Item(0).Control(12)=   "lblAlta(4)"
      Item(0).Control(13)=   "vtipodisponibilidad"
      Begin XtremeSuiteControls.ComboBox vtipocaja 
         Height          =   315
         Left            =   2760
         TabIndex        =   14
         Top             =   1320
         Width           =   3075
         _Version        =   851968
         _ExtentX        =   5424
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   0
         Left            =   2760
         TabIndex        =   0
         Top             =   600
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   1
         Left            =   2760
         TabIndex        =   1
         Top             =   960
         Width           =   4455
         _Version        =   851968
         _ExtentX        =   7858
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   255
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   3
         Left            =   2700
         TabIndex        =   3
         Top             =   2460
         Width           =   1215
         _Version        =   851968
         _ExtentX        =   2143
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   10
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   4
         Left            =   4500
         TabIndex        =   4
         Top             =   2460
         Width           =   2655
         _Version        =   851968
         _ExtentX        =   4683
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton pbCarga 
         Height          =   315
         Index           =   0
         Left            =   4020
         TabIndex        =   5
         Tag             =   "CodigoCuenta"
         Top             =   2460
         Width           =   315
         _Version        =   851968
         _ExtentX        =   556
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "..."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   2
         Left            =   5910
         TabIndex        =   2
         Top             =   1320
         Width           =   1275
         _Version        =   851968
         _ExtentX        =   2249
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   2
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.ComboBox vtipodisponibilidad 
         Height          =   315
         Left            =   2760
         TabIndex        =   17
         Top             =   1680
         Width           =   3075
         _Version        =   851968
         _ExtentX        =   5424
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de Disponibilidad para uso:"
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   16
         Top             =   1770
         Width           =   2325
      End
      Begin XtremeSuiteControls.Label vlabel 
         Height          =   285
         Left            =   2760
         TabIndex        =   15
         Top             =   2070
         Width           =   4425
         _Version        =   851968
         _ExtentX        =   7805
         _ExtentY        =   503
         _StockProps     =   79
         ForeColor       =   255
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Es Caja? 'S' o 'N':"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   10
         Top             =   1365
         Width           =   2055
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre del Banco:"
         Height          =   195
         Index           =   1
         Left            =   450
         TabIndex        =   9
         Top             =   1005
         Width           =   2175
      End
      Begin VB.Label lblAlta 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo Banco:"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   645
         Width           =   2200
      End
      Begin VB.Label lblAlta 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Contable Asociada :"
         Height          =   195
         Index           =   3
         Left            =   420
         TabIndex        =   7
         Top             =   2505
         Width           =   2205
      End
   End
End
Attribute VB_Name = "frmBancosAlta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vAccion As String
Private Sub Form_Load()
On Error Resume Next
    
    With Me
        .Show
    End With
    
    
    Limpiar
    
    init
    
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub init()

vtipocaja.Clear

vtipocaja.AddItem "Caja"
vtipocaja.AddItem "Banco Propio"
vtipocaja.AddItem "Nombre de Bancos "


Me.vtipocaja.Clear

Me.vtipodisponibilidad.AddItem ("Disponible")
Me.vtipodisponibilidad.AddItem ("Interno")

Me.vtipodisponibilidad.Text = "Disponible"

vtipocaja.Text = "Caja"

End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            GuardarBanco

            
        Case 1
            Limpiar
            Unload Me
                        
    End Select
    
If Err Then GrabarLog "PBAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Limpiar()
On Error Resume Next

    Dim i As Integer
        
    vAccion = "Nuevo"
        
    For i = 0 To txtAlta.Count - 1
        txtAlta(i).Text = ""
        txtAlta(i).Tag = ""
    Next
    
    
    txtAlta(0).Locked = Not True
    txtAlta(0).Text = Val(GenerarDato("SELECT MAX(idBancos) AS UltimoCodigo FROM Bancos", "UltimoCodigo")) + 1
    txtAlta(0).Text = FormatoUltimoCodigo(3, txtAlta(0).Text)

    txtAlta(0).SetFocus
    
    
    
If Err Then GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub GuardarBanco()
On Error Resume Next

    Dim vsql, vcampos, vvalores As String

    If ValidarCampos = Not True Then
        MsgBox "Uno de los campos No es Valido o no esta bien Cargado!!", vbExclamation, "Mensaje ..."
        Exit Sub
    End If
    
    Dim rsBancos As New ADODB.Recordset, sqlBancos As String
    
    Select Case vAccion

        Case "Nuevo"
            sqlBancos = "SELECT * FROM Bancos WHERE 1=2"
        
        Case "Modificar"
            sqlBancos = "SELECT * FROM Bancos WHERE (idBancos = '" & Trim$(txtAlta(0).Text) & "')"
            vsql = "delete from bancos WHERE (idBancos = '" & Trim$(txtAlta(0).Text) & "')"
            Call EjecutarScript(vsql, pathDBMySQL)
        
        Case "Duplicar"
            
    End Select
    
    With rsBancos
        
        
        
        vvalores = "'" + Trim(txtAlta(0).Text) + "','" + Left(txtAlta(1).Text, 255) + "','" + Trim(UCase(Left(txtAlta(2).Text, 1))) + "','" + Trim(txtAlta(3).Text) + "','" + Me.vtipodisponibilidad + "'"
        vcampos = "idBancos,Descripcion,EsCaja,CuentaContableAsociada,tipodisponibilidad"
        
        vsql = "insert into Bancos (" + vcampos + ") values (" + vvalores + ")"
        
        Call EjecutarScript(vsql, pathDBMySQL)
        
        
        'Call .Open(sqlBancos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
       ' If .State = 1 Then If .EOF = True Then .AddNew
            
       ' .Fields("idBancos").Value = Left(txtAlta(0).Text, 3)
       ' .Fields("Descripcion").Value = Left(txtAlta(1).Text, 255)
       ' .Fields("EsCaja").Value = UCase(Left(txtAlta(2).Text, 1))
       ' .Fields("CuentaContableAsociada").Value = Trim(txtAlta(3).Text)
    
       ' .Update
    End With

    sqlBancos = ""
    
    If rsBancos.State = 1 Then
        rsBancos.Close
        Set rsBancos = Nothing
    End If

If Err Then
    GrabarLog "Limpiar", Err.Number & " " & Err.Description, Me.Caption
Else
    Limpiar
    frmBancos.Buscar
    Unload Me
End If
End Sub
Private Function ValidarCampos() As Boolean
On Error Resume Next

    Dim i As Integer
    ValidarCampos = True
    
    
    'If Not traerDatos2("select * from bancos where idbancos ='" + Trim(Me.txtAlta(0)) + "'", "idBancos", pathDBMySQL) = "" Then
    '    MsgBox "Código de banco/caja existente", vbInformation, "Cuidado..."
    '    ValidarCampos = False
    '    Exit Function
    'End If
    
    
    For i = 0 To txtAlta.Count - 4
        If txtAlta(i).Text = "" Then
            ValidarCampos = Not True
            Exit Function
        End If
    Next

    If Not UCase$(txtAlta(2).Text) = "S" And Not UCase$(txtAlta(2).Text) = "N" And Not UCase$(txtAlta(2).Text) = "B" Then
        ValidarCampos = Not True
        Exit Function
    End If
    
    If vAccion = "Nuevo" Then
        If Not Trim(TraerDato("Bancos", "idBancos = '" & Trim$(txtAlta(0).Text) & "'", "idBancos")) = "" Then
            MsgBox "Este código fue ingresado en otra caja/banco", vbCritical
           ValidarCampos = Not True
            Exit Function
        End If
    End If
    
If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub pbCarga_Click(Index As Integer)
    On Error Resume Next

    vVuelveBusqueda = Me.Name
    vVieneBusqueda = pbCarga(Index).Tag

    Select Case Index
    
        Case 0 To pbCarga.Count - 1
            frmBusqueda.Show

    End Select

If Err Then GrabarLog "pbCarga_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Public Sub ModificarBanco(vidBancos As String)
    On Error Resume Next
    
    Dim rsBancos As New ADODB.Recordset, sqlBancos As String
    
    sqlBancos = "SELECT * FROM Bancos WHERE (idBancos = '" & vidBancos & "')"
    
    With rsBancos
        Call .Open(sqlBancos, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
        
            'No Opcionales
            txtAlta(0).Text = .Fields("idBancos").Value
            txtAlta(0).Locked = True
        
            txtAlta(1).Text = .Fields("Descripcion").Value
            txtAlta(2).Text = .Fields("EsCaja").Value
            txtAlta(3).Text = .Fields("CuentaContableAsociada").Value
            txtAlta(4).Text = TraerDato("Cuentas", "CodigoCuenta = '" & .Fields("CuentaContableAsociada").Value & "'", "Cuenta")
            Me.vtipodisponibilidad.Text = .Fields("tipodisponibilidad").Value
        
        End If

    End With
        
    sqlBancos = ""
    
    If rsBancos.State = 1 Then
        rsBancos.Close
        Set rsBancos = Nothing
    End If
    
    If Err Then GrabarLog "ModificarCliente", Err.Number & " " & Err.Description, Me.Name
End Sub


Private Sub txtAlta_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        If Not (Index = txtAlta.Count - 1) Then
            If txtAlta(Index + 1).Visible = True Then
                txtAlta(Index + 1).SetFocus
            End If
        Else
            PbAcciones(0).SetFocus
        End If
    End If

If Err Then GrabarLog "txtAlta_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtAlta_LostFocus(Index As Integer)
'If Index = 0 Then txtAlta(0).Text = Format(txtAlta(0).Text, "00000")
End Sub

Private Sub vtipocaja_Click()
Select Case vtipocaja

Case "Caja"
    Me.txtAlta(2).Text = "S"
    vlabel.Caption = "Cajas propias para cheques, documentos, vale, documentos"
Case "Banco Propio"
    Me.txtAlta(2).Text = "N"
    vlabel.Caption = "Nombre de las cuentas bancarias"

Case "Nombre de Bancos "
    Me.txtAlta(2).Text = "B"
    vlabel.Caption = "Nombre de bancos argentinos"
    
End Select

End Sub

Private Sub vtipocaja_Change()
Select Case vtipocaja

Case "Caja"
    Me.txtAlta(2).Text = "S"
    vlabel.Caption = "Cajas propias para cheques, documentos, vale, documentos"
Case "Banco Propio"
    Me.txtAlta(2).Text = "N"
    vlabel.Caption = "Nombre de las cuentas bancarias"

Case "Nombre de Bancos "
    Me.txtAlta(2).Text = "B"
    vlabel.Caption = "Nombre de bancos argentinos"
    
End Select


End Sub

