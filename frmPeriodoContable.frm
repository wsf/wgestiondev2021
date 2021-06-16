VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Object = "{9746E3DA-06E1-4D26-9CE4-D9F6411A9C70}#1.0#0"; "SMGA_OcxTxt2009.ocx"
Begin VB.Form frmPeriodoContable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Periodo Contable"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   9450
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   135
      Left            =   0
      TabIndex        =   12
      Top             =   2340
      Width           =   9435
      _Version        =   851968
      _ExtentX        =   16642
      _ExtentY        =   238
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.TabControl TabBalance 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   9435
      _Version        =   851968
      _ExtentX        =   16642
      _ExtentY        =   4048
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Datos Generales"
      Item(0).ImageIndex=   0
      Item(0).ControlCount=   8
      Item(0).Control(0)=   "lblDatos(0)"
      Item(0).Control(1)=   "lblDatos(1)"
      Item(0).Control(2)=   "lblDatos(2)"
      Item(0).Control(3)=   "txtAlta(0)"
      Item(0).Control(4)=   "dtpFecha(1)"
      Item(0).Control(5)=   "dtpFecha(0)"
      Item(0).Control(6)=   "lblDatos(3)"
      Item(0).Control(7)=   "vcodigo"
      Item(1).Caption =   "Datos"
      Item(1).ControlCount=   3
      Item(1).Control(0)=   "grilla"
      Item(1).Control(1)=   "PushButton1"
      Item(1).Control(2)=   "PushButton2"
      Begin XtremeSuiteControls.FlatEdit vcodigo 
         Height          =   285
         Left            =   3210
         TabIndex        =   11
         Top             =   1650
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   315
         Left            =   -61840
         TabIndex        =   8
         Top             =   420
         Visible         =   0   'False
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Activar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPeriodoContable.frx":0000
      End
      Begin Grid.KlexGrid grilla 
         Height          =   1785
         Left            =   -69910
         TabIndex        =   7
         Top             =   390
         Visible         =   0   'False
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   3149
         EnterKeyBehaviour=   0
         BackColorAlternate=   0
         GridLines       =   0
         GridLinesFixed  =   2
         BackColorFixed  =   -2147483626
         Cols            =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GridColorFixed  =   8421504
         MouseIcon       =   "frmPeriodoContable.frx":059A
         Rows            =   10
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   0
         Left            =   3195
         TabIndex        =   4
         Top             =   570
         Width           =   3735
         _Version        =   851968
         _ExtentX        =   6588
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         Alignment       =   1
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   300
         Index           =   0
         Left            =   3210
         TabIndex        =   5
         Top             =   960
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
      End
      Begin Aplisoft_CajasDeTexto.TxF dtpFecha 
         Height          =   300
         Index           =   1
         Left            =   3210
         TabIndex        =   6
         Top             =   1290
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackStyle       =   0
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   345
         Left            =   -61840
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1185
         _Version        =   851968
         _ExtentX        =   2090
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Desactivar"
         UseVisualStyle  =   -1  'True
         Picture         =   "frmPeriodoContable.frx":05B6
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "> Codigo del balance:"
         Height          =   195
         Index           =   3
         Left            =   1380
         TabIndex        =   10
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "> Fecha Fin :"
         Height          =   195
         Index           =   2
         Left            =   1410
         TabIndex        =   3
         Top             =   1350
         Width           =   1305
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "> Fecha Inicio :"
         Height          =   195
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label lblDatos 
         BackStyle       =   0  'Transparent
         Caption         =   "> Nro de Balance :"
         Height          =   195
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   630
         Width           =   1425
      End
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   345
      Index           =   0
      Left            =   60
      TabIndex        =   13
      Top             =   2550
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Grabar"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmPeriodoContable.frx":0B50
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   345
      Index           =   1
      Left            =   8280
      TabIndex        =   14
      Top             =   2550
      Visible         =   0   'False
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmPeriodoContable.frx":0F57
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   345
      Index           =   2
      Left            =   1170
      TabIndex        =   15
      Top             =   2550
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Borrar"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmPeriodoContable.frx":1357
   End
End
Attribute VB_Name = "frmPeriodoContable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vaccion As String
Dim rsPeriodoContable As New ADODB.Recordset, sqlPeriodoContable As String

Private Sub dtpFecha_KeyPress(Index As Integer, Keyascii As Integer)
On Error Resume Next
    
    If Keyascii = 13 Then
    
        Select Case Index
    
            Case 0
                dtpFecha(1).SetFocus
            
            Case 1
                Me.vcodigo.SetFocus
        
        End Select
    
    End If
    
If Err Then GrabarLog "dtpFecha_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub actualizarGrilla()
On Error Resume Next
    If rsPeriodoContable.State = 1 Then rsPeriodoContable.Close
    Call rsPeriodoContable.Open("select * from balances", ConnDDBB, adOpenStatic, adLockPessimistic)
    Set grilla.Recordset = rsPeriodoContable
If Err Then Exit Sub
End Sub
Private Sub Form_Load()
On Error Resume Next

Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2 - 1000

    'With Me
    '    .Show
    'End With

    LimpiarCampos
    
    actualizarGrilla
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Grabar
            
        Case 1
            Unload Me
            'MigrarLocalidades
        Case 2
            borrarLinea
    End Select

If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub borrarLinea()
If MsgBox("Está seguro de borrar la linea ?", vbYesNo, "Borrando") = vbNo Then Exit Sub

Dim vsql As String
vsql = "delete from balances where idBalances=" + grilla.TextMatrix(grilla.Row, 1)
Call EjecutarScript(vsql, pathDBMySQL)
actualizarGrilla
End Sub
Private Sub limpiarcampo()
Me.txtAlta(0).Text = ""

Me.dtpFecha(0).Value = Date
Me.dtpFecha(1).Value = Date


End Sub
Private Sub Grabar()
    On Error Resume Next

    If Not ValidarCampos() = True Then
        Exit Sub
    End If
    
    
    Select Case vaccion

        Case "Nuevo"
            sqlPeriodoContable = "SELECT * FROM Balances WHERE 1=2"
        
        Case "Modificar"
            sqlPeriodoContable = "SELECT * FROM Balances WHERE (idBalances = " & Trim(txtAlta(0).Text) & ")"
        
        Case "Duplicar"
            
    End Select
        
    With rsPeriodoContable
        Call .Open(sqlPeriodoContable, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .State = 0 Then
        
            Select Case vaccion
            
                Case "Nuevo"
                    .AddNew
                    .Fields("NroBalance").Value = Val(txtAlta(0).Text)
                
                Case "Modificar"
                    'No hago nada
                    
                Case "Duplicar"
                    .AddNew

            End Select
            
            
            .Fields("FechaInicio").Value = strfechaMySQL(dtpFecha(0).Value)
            .Fields("FechaFin").Value = strfechaMySQL(dtpFecha(1).Value)
            .Fields("Activo").Value = "S"
            .Fields("codigo").Value = Me.vcodigo.Text
            

            .Update
        
        End If
        
    End With

    sqlPeriodoContable = ""
    
    actualizarGrilla
    
    
    If rsPeriodoContable.State = 1 Then
        rsPeriodoContable.Close
        Set rsPeriodoContable = Nothing
    End If
    
    If Err Then
        GrabarLog "Guardar", Err.Number & " " & Err.Description, Me.Name
    Else
        LimpiarCampos
        Unload Me
        'frmClientes.Buscar
    End If

End Sub
Private Function ValidarCampos() As Boolean
    On Error Resume Next

    Dim i As Integer
    
    ValidarCampos = True
    
    For i = 0 To Val(txtAlta.Count - 1)

        If Trim(txtAlta(i).Text) = "" Then
            MsgBox "Campos obligatorios vacios!", vbExclamation, "Mensaje ..."
            ValidarCampos = Not True
            Exit Function
        End If

    Next
    
    If Not Trim(TraerDato("Balances", "NroBalance = " & Trim(txtAlta(0).Text) & "", "idBalances")) = "" Then
        MsgBox "Existe un registro con ese codigo!", vbExclamation, "Mensaje ..."
        ValidarCampos = Not True
        Exit Function
    End If
    
    If Not Trim(TraerDato("Balances", "(FechaInicio >= '" & strfechaMySQL(dtpFecha(0).Value) & "' AND FechaFin <= '" & strfechaMySQL(dtpFecha(0).Value) & "')", "idBalances")) = "" Then
        MsgBox "La fecha ingresada existe en algun periodo ya guardado!", vbExclamation, "Mensaje ..."
        ValidarCampos = Not True
        Exit Function
    End If
    
    If Not Trim(TraerDato("Balances", "(FechaInicio >= '" & strfechaMySQL(dtpFecha(1).Value) & "' AND FechaFin <= '" & strfechaMySQL(dtpFecha(1).Value) & "')", "idBalances")) = "" Then
        MsgBox "La fecha ingresada existe en algun periodo ya guardado!", vbExclamation, "Mensaje ..."
        ValidarCampos = Not True
        Exit Function
    End If
    
    
    If dtpFecha(0).Value >= dtpFecha(1).Value Then
        MsgBox "El rango de Fecha Ingresado No es Válido!!", vbExclamation, "Mensaje ..."
        ValidarCampos = Not True
        Exit Function
    End If
    
    If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub LimpiarCampos()
On Error Resume Next

    
    vaccion = "Nuevo"
    
    txtAlta(0).Text = ""
    dtpFecha(0).Value = Date
    dtpFecha(1).Value = Date
    
    
    txtAlta(0).SetFocus
    
If Err Then GrabarLog "LimpiarCampos", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub PushButton1_Click()
On Error Resume Next

Dim vsql As String
Dim vid, vrow As Long
vrow = grilla.Row
vid = grilla.TextMatrix(vrow, 1)

vsql = "update balances set Activo='S' where idBalances=" + Str(vid)
Call EjecutarScript(vsql, pathDBMySQL)

actualizarGrilla

If Err Then Exit Sub
End Sub

Private Sub PushButton2_Click()
Dim vsql As String
Dim vid, vrow As Long
vrow = grilla.Row
vid = grilla.TextMatrix(vrow, 1)

vsql = "update balances set Activo='N' where idBalances=" + Str(vid)
Call EjecutarScript(vsql, pathDBMySQL)

actualizarGrilla

End Sub

Private Sub TabBalance_BeforeItemClick(ByVal Item As XtremeSuiteControls.ITabControlItem, Cancel As Variant)
actualizarGrilla
End Sub

Private Sub txtAlta_KeyPress(Index As Integer, Keyascii As Integer)
On Error Resume Next

    If Keyascii = 13 Then
        Select Case Index
        
            Case 0
                dtpFecha(0).SetFocus
            Case 1
            
            Case 2
        
        End Select
        
    End If

If Err Then GrabarLog "txtAlta_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub vcodigo_KeyPress(Keyascii As Integer)
If Keyascii = 13 Then PbAcciones(0).SetFocus
End Sub
