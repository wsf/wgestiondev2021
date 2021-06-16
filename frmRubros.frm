VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmRubros 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Rubros & SubRubros"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5940
   Begin XtremeSuiteControls.TabControl TabRubros 
      Height          =   3975
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   5895
      _Version        =   851968
      _ExtentX        =   10398
      _ExtentY        =   7011
      _StockProps     =   68
      ItemCount       =   2
      Item(0).Caption =   "Rubros"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "dgRubros"
      Item(1).Caption =   "Sub-Rubros"
      Item(1).ControlCount=   1
      Item(1).Control(0)=   "dgSubRubros"
      Begin MSDataGridLib.DataGrid dgRubros 
         Bindings        =   "frmRubros.frx":0000
         Height          =   3285
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   5794
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   2
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
      Begin MSDataGridLib.DataGrid dgSubRubros 
         Bindings        =   "frmRubros.frx":0016
         Height          =   3285
         Left            =   -69880
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   5794
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         HeadLines       =   2
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
   Begin VB.PictureBox PicInferior 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      Picture         =   "frmRubros.frx":002C
      ScaleHeight     =   555
      ScaleWidth      =   5850
      TabIndex        =   1
      Top             =   5160
      Width           =   5850
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   0
         Left            =   2520
         TabIndex        =   2
         Top             =   90
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   3
         Top             =   90
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Borrar"
         Appearance      =   6
      End
      Begin XtremeSuiteControls.PushButton cmdAcciones 
         Height          =   375
         Index           =   2
         Left            =   4680
         TabIndex        =   4
         Top             =   90
         Width           =   1095
         _Version        =   851968
         _ExtentX        =   1931
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   6
      End
      Begin VB.Label lblWGestion 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
         TabIndex        =   5
         Top             =   150
         Width           =   1770
      End
      Begin VB.Label lblWGestion 
         BackStyle       =   0  'Transparent
         Caption         =   "WGESTION 2010"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
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
         TabIndex        =   6
         Top             =   170
         Width           =   1770
      End
   End
   Begin VB.Frame fraAcciones 
      Caption         =   "Acciones:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   90
      TabIndex        =   0
      Top             =   4020
      Width           =   5745
      Begin XtremeSuiteControls.FlatEdit txtRubro 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   240
         Width           =   3720
         _Version        =   851968
         _ExtentX        =   6562
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   3
      End
      Begin XtremeSuiteControls.FlatEdit txtRubro 
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   8
         Top             =   600
         Width           =   3720
         _Version        =   851968
         _ExtentX        =   6562
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
         MaxLength       =   255
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1725
         _Version        =   851968
         _ExtentX        =   3043
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Codigo (ID) :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeSuiteControls.Label lblActualizar 
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1710
         _Version        =   851968
         _ExtentX        =   3016
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Descripcion :"
         Alignment       =   1
         Transparent     =   -1  'True
      End
   End
End
Attribute VB_Name = "frmRubros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vModificando As Boolean
Dim rsRubros As ADODB.Recordset
Dim rsSubRubros As ADODB.Recordset
Private Sub cmdAcciones_Click(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            Agregar

        
        Case 1
            Select Case TabRubros.Item(Index).Selected
                
                Case 0
                    With rsRubros
                        If Not (.EOF = True) And Not (.BOF = True) Then
                            Call BorrarBase("Rubros WHERE (idRubros = '" & Trim(.Fields("idRubros").Value) & "')", pathDBMySQL)
                            Nuevo
                        End If
                    End With
            
                Case 1
                    With rsSubRubros
                        If Not (.EOF = True) And Not (.BOF = True) Then
                            Call BorrarBase("SubRubros WHERE (idSubRubros = '" & Trim(.Fields("idSubRubros").Value) & "')", pathDBMySQL)
                            Nuevo
                        End If
                    End With
                
            End Select
        Case 2
            Unload Me
    
    End Select
    
If Err Then GrabarLog "cmdAcciones_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Modificar()
On Error Resume Next

    Select Case TabRubros.SelectedItem
    
        Case 0
            txtRubro(0).Text = EsNulo(rsRubros.Fields("idRubros").Value)
            txtRubro(1).Text = EsNulo(rsRubros.Fields("Rubro").Value)
            dgRubros.Enabled = False
        
        Case 1
            txtRubro(0).Text = EsNulo(rsSubRubros.Fields("idSubRubros").Value)
            txtRubro(1).Text = EsNulo(rsSubRubros.Fields("SubRubro").Value)
            dgSubRubros.Enabled = False
    End Select
    
    fraAcciones.Caption = "Acciones - Modificando"
    
    cmdAcciones(0).Caption = "Actualizar"

    vModificando = True
    
If Err Then GrabarLog "Modificar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgRubros_DblClick()
On Error Resume Next

    Modificar
    
If Err Then GrabarLog "dgRubros_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub dgSubRubros_DblClick()
On Error Resume Next

    Modificar
    
If Err Then GrabarLog "dgSubRubros_DblClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    With Me
        .Show
        .Width = 6000
        .Height = 6250
        .Top = 300
        .Left = 1000
    End With
    
    Nuevo
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarRubros()
On Error Resume Next

    Set rsRubros = New ADODB.Recordset
    
    With rsRubros
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        Call .Open("SELECT * FROM Rubros", ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then
        
        Else
        
        End If
        
        Set dgRubros.DataSource = rsRubros
        
        FormatoGrilla (0)
    End With
    
If Err Then GrabarLog "CargarRubros", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CargarSubRubros()
On Error Resume Next

    Set rsSubRubros = New ADODB.Recordset
    
    With rsSubRubros
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        Call .Open("SELECT * FROM SubRubros", ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 Then
        
        Else
        
        End If
        
        Set dgSubRubros.DataSource = rsSubRubros
        
        FormatoGrilla (1)
    End With
    
If Err Then GrabarLog "CargarRubros", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            With dgRubros
                .HeadLines = 1.2
        
                .Columns(0).Width = 1000
                .Columns(1).Width = 4000
    
            End With
        Case 1
            With dgSubRubros
                .HeadLines = 1.2
        
                .Columns(0).Width = 1000
                .Columns(1).Width = 0
                .Columns(2).Width = 4000
    
            End With
    
    End Select

If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Agregar()
On Error Resume Next
    
    If Trim(txtRubro(0).Text) = "" Then
        MsgBox "Debe ingresar un codigo de SubRubro/Rubro", vbInformation, "Mensaje ..."
        Exit Sub
    End If

    If Trim(txtRubro(1).Text) = "" Then
        MsgBox "Debe ingresar un SubRubro/Rubro", vbInformation, "Mensaje ..."
        Exit Sub
    End If
    
    Select Case TabRubros.SelectedItem
    
         Case 0
            If Not (TraerDato("Rubros", "idRubros = '" & Trim(txtRubro(0).Text) & "'", "idRubros") = "") And Not (vModificando = True) Then
                MsgBox "El Codigo/Rubro Ingresado ya existe en el sistema", vbInformation, "Mensaje ..."
                Exit Sub
            End If
            
            If Not vModificando = True Then
                Call EjecutarScript("INSERT INTO Rubros (idRubros, Rubro) VALUES ('" & Trim(txtRubro(0).Text) & "', '" & Trim(txtRubro(1).Text) & "')")
            Else
                Call EjecutarScript("UPDATE Rubros SET idRubros = '" & Trim(txtRubro(0).Text) & "', Rubro = '" & Trim(txtRubro(1).Text) & "' WHERE idRubros =  '" & Trim(txtRubro(0).Text) & "'")
            End If
    
        Case 1
            If Not (TraerDato("SubRubros", "idSubRubros = '" & Trim(txtRubro(0).Text) & "'", "idSubRubros") = "") And Not (vModificando = True) Then
                MsgBox "El Codigo/Rubro Ingresado ya existe en el sistema", vbInformation, "Mensaje ..."
                Exit Sub
            End If
            
            If Not vModificando = True Then
                Call EjecutarScript("INSERT INTO SubRubros (idSubRubros, SubRubro) VALUES ('" & Trim(txtRubro(0).Text) & "', '" & Trim(txtRubro(1).Text) & "')")
            Else
                Call EjecutarScript("UPDATE SubRubros SET idSubRubros = '" & Trim(txtRubro(0).Text) & "', SubRubro = '" & Trim(txtRubro(1).Text) & "' WHERE idSubRubros =  '" & Trim(txtRubro(0).Text) & "'")
            End If
        
    End Select
        
    Nuevo

If Err Then GrabarLog "Agregar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Nuevo()
On Error Resume Next

    Select Case TabRubros.SelectedItem
    
        Case 0
            fraAcciones.Caption = "Acciones - Nuevo Rubro"
            dgRubros.Enabled = True
        
        Case 1
            fraAcciones.Caption = "Acciones - Nuevo Sub-Rubro"
            dgSubRubros.Enabled = True
    
    End Select
    
    txtRubro(0).Text = ""
    txtRubro(1).Text = ""
    
    cmdAcciones(0).Caption = "Agregar"
    
    vModificando = False
    
    CargarRubros
    CargarSubRubros
    
If Err Then GrabarLog "Nuevo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub TabRubros_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
On Error Resume Next

    Nuevo

If Err Then GrabarLog "TabRubros_SelectedChanged", Err.Number & " " & Err.Description, Me.Caption
End Sub
