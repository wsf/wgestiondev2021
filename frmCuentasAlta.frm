VERSION 5.00
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmCuentasAltas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   11490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   11490
   Begin XtremeSuiteControls.GroupBox GroupBox1 
      Height          =   225
      Left            =   0
      TabIndex        =   8
      Top             =   390
      Width           =   11445
      _Version        =   851968
      _ExtentX        =   20188
      _ExtentY        =   397
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      BorderStyle     =   1
   End
   Begin XtremeSuiteControls.GroupBox GBCuentas 
      Height          =   2325
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   11445
      _Version        =   851968
      _ExtentX        =   20188
      _ExtentY        =   4101
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.FlatEdit vnivel 
         Height          =   315
         Left            =   2280
         TabIndex        =   11
         Top             =   1200
         Width           =   2925
         _Version        =   851968
         _ExtentX        =   5159
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   0
         Left            =   2280
         TabIndex        =   1
         Top             =   420
         Width           =   2895
         _Version        =   851968
         _ExtentX        =   5106
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.FlatEdit txtAlta 
         Height          =   315
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Top             =   810
         Width           =   8985
         _Version        =   851968
         _ExtentX        =   15849
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.CheckBox chkImputable 
         Height          =   315
         Left            =   2250
         TabIndex        =   3
         Top             =   1740
         Width           =   4815
         _Version        =   851968
         _ExtentX        =   8493
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "La Cuenta NO es Imputable"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label lblCuentas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ing. Nivel de Jerarquía: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   12
         Top             =   1260
         Width           =   1710
      End
      Begin VB.Label lblFormatoContable 
         Caption         =   "> Formato contable:"
         ForeColor       =   &H00000000&
         Height          =   165
         Left            =   6450
         TabIndex        =   7
         Top             =   480
         Width           =   1515
      End
      Begin XtremeSuiteControls.Label lblCodigoFormateo 
         Height          =   345
         Left            =   8010
         TabIndex        =   6
         Top             =   390
         Width           =   3255
         _Version        =   851968
         _ExtentX        =   5741
         _ExtentY        =   609
         _StockProps     =   79
         ForeColor       =   14737632
         BackColor       =   -2147483636
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
      End
      Begin VB.Label lblCuentas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese del Codigo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   1410
      End
      Begin VB.Label lblCuentas 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ingrese del Nombre:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   420
         TabIndex        =   4
         Top             =   840
         Width           =   1425
      End
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   345
      Index           =   0
      Left            =   30
      TabIndex        =   9
      Top             =   30
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Grabar"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCuentasAlta.frx":0000
   End
   Begin XtremeSuiteControls.PushButton PbAcciones 
      Height          =   345
      Index           =   1
      Left            =   10380
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   1095
      _Version        =   851968
      _ExtentX        =   1931
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Cerrar"
      UseVisualStyle  =   -1  'True
      Picture         =   "frmCuentasAlta.frx":0407
   End
End
Attribute VB_Name = "frmCuentasAltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vaccion As String
Private Sub Form_Load()
    On Error Resume Next

    With Me
        .Show
        '.Top = 0
        '.Left = 0
        '.Width = 7500
        '.Height = 2725
    End With
    
    Call LimpiarCampos

    CentrarFormulario Me
    
    
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub PbAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            GrabarCuenta
        
        Case 1
            
            Unload Me
    
    End Select

If Err Then GrabarLog "PbAcciones_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtAlta_Change(Index As Integer)
On Error Resume Next
    
    If Index = 0 Then
        lblCodigoFormateo.Caption = MostrarCodigoCuenta(txtAlta(Index).Text)
    End If
    
    Me.vnivel.Text = Replace(Replace(Replace(txtAlta(0), "-", ""), ".", ""), "0", "")
    
    
If Err Then GrabarLog "txtAlta_Change", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub txtAlta_KeyPress(Index As Integer, KeyAscii As Integer)
On Error Resume Next

    If KeyAscii = 13 Then
        Select Case Index
    
            Case 0
                txtAlta(1).SetFocus
        
            Case 1
                PbAcciones(0).SetFocus
    
        End Select

    End If

If Err Then GrabarLog "txtCodigo_KeyPress", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub GrabarCuenta()
    On Error Resume Next

    If Not ValidarCampos() = True Then
        Exit Sub
    End If
    
    Dim rsCuentas As New ADODB.Recordset, sqlCuentas As String
    
    Select Case vaccion

        Case "Nuevo"
            sqlCuentas = "SELECT * FROM Cuentas WHERE 1=2"
        
        Case "Modificar"
            sqlCuentas = "SELECT * FROM Cuentas WHERE (CodigoCuenta = '" & Trim(txtAlta(0).Text) & "')"
        
        Case "Duplicar"
            
    End Select
        
    With rsCuentas
        Call .Open(sqlCuentas, ConnDDBB, adOpenStatic, adLockPessimistic)
        
       ' validar()
       ' If Len(rsCuentas.Fields("CodigoCuenta")) = Len(txtAlta(0).Text) Then
       '    MsgBox "No c"
       ' End If
        
        If Not .State = 0 Then
        
            Select Case vaccion
            
                Case "Nuevo"
                    .AddNew
                
                Case "Modificar"
                    'No hago nada
                    
                Case "Duplicar"
                    '.AddNew
                    '.Fields("NCheque").Value = Trim(txtAlta(0).Text)
                    '.Fields("Fecha").Value = strfechaMySQL(dtpFecha(0).Value)
                    '.Fields("FechaDeposito").Value = strfechaMySQL(dtpFecha(0).Value)
            End Select



            .Fields("CodigoCuenta").Value = Trim(txtAlta(0).Text) 'Left(txtAlta(0).Text, 10)
            .Fields("Cuenta").Value = Left(txtAlta(1).Text, 255)
            .Fields("niveles").Value = Trim(Me.vnivel)
        
            If chkImputable.Value = xtpChecked Then
                .Fields("Imputable").Value = "N"
            Else
                 .Fields("Imputable").Value = "S"
            End If
            
            .Update
        
        End If
        
    End With

    sqlCuentas = ""
    
    If rsCuentas.State = 1 Then
        rsCuentas.Close
        Set rsCuentas = Nothing
    End If
    
    If Err Then
        GrabarLog "GrabarCuenta", Err.Number & " " & Err.Description, Me.Name
    Else
        LimpiarCampos
        Unload Me
        frmCuentas.Buscar
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
        
    If vaccion = "Nuevo" Then
        If Not Trim(TraerDato("Cuentas", "CodigoCuenta = '" & Trim(txtAlta(0).Text) & "'", "idCuentas")) = "" Then
            MsgBox "Existe un registro con ese Numero!", vbExclamation, "Mensaje ..."
            ValidarCampos = Not True
            Exit Function
        End If
    End If
    
    If Err Then GrabarLog "ValidarCampos", Err.Number & " " & Err.Description, Me.Caption
End Function
Private Sub LimpiarCampos()
    On Error Resume Next
    
    Dim i As Integer
    
    For i = 0 To txtAlta.Count - 1
        txtAlta(i).Text = ""
        txtAlta(i).Tag = ""
    Next
    
    vaccion = "Nuevo"
    KeyPreview = True

    If Err Then GrabarLog "LimpiarCampos", Err.Number & "-" & Err.Description, Me.Name
End Sub
Public Sub ModificarCuenta(vIDCuenta As Long)
    On Error Resume Next
    
    Dim rsCuenta As New ADODB.Recordset, sqlCuenta As String
    
    sqlCuenta = "SELECT * FROM Cuentas WHERE (idCuentas = " & vIDCuenta & ")"
    
    With rsCuenta
        Call .Open(sqlCuenta, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not (.EOF = True) And Not (.BOF = True) Then
        
            txtAlta(0).Text = EsNulo(.Fields("CodigoCuenta").Value)
            txtAlta(0).Locked = True
        
            txtAlta(1).Tag = EsNulo(.Fields("idCuentas").Value)
            txtAlta(1).Text = EsNulo(.Fields("Cuenta").Value)
        
            If EsNulo(.Fields("Imputable").Value) = "S" Then
                chkImputable.Value = xtpUnchecked
            Else
                chkImputable.Value = xtpChecked
            End If

        End If

    End With
    
    sqlCuenta = ""

    If rsCuenta.State = 1 Then
        rsCuenta.Close
        Set rsCuenta = Nothing
    End If
    
    If Err Then GrabarLog "ModificarCuenta", Err.Number & " " & Err.Description, Me.Name
End Sub
