VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmConfigRemito 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Configuracion del Remito & Migrador"
   ClientHeight    =   4920
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6150
   Icon            =   "frmConfigRemito.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab TabGeneral 
      Height          =   4335
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   7646
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "General"
      TabPicture(0)   =   "frmConfigRemito.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraGeneral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Alarmas"
      TabPicture(1)   =   "frmConfigRemito.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraAlarmas"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Teclado Abreviado"
      TabPicture(2)   =   "frmConfigRemito.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   5895
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Caption         =   "Formulario de Remito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            TabIndex        =   21
            Top             =   120
            Width           =   5805
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Caption         =   "Formulario de Migrador"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            TabIndex        =   20
            Top             =   2100
            Width           =   5805
         End
      End
      Begin VB.Frame fraAlarmas 
         Height          =   3735
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   5895
         Begin VB.CheckBox chkalarma 
            Caption         =   "Facturar sin Nro de Lista"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   600
            Width           =   5700
         End
         Begin VB.CheckBox chkalarma 
            Caption         =   "Guardar igual sin Nro de Lista"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   5700
         End
         Begin VB.CheckBox chkalarma 
            Caption         =   "No Guardar sin un Repartidor"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   240
            Width           =   5700
         End
      End
      Begin VB.Frame fraGeneral 
         Height          =   3735
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   5895
         Begin VB.ComboBox cboGral 
            Height          =   315
            Index           =   1
            Left            =   2640
            TabIndex        =   23
            Top             =   2160
            Width           =   3015
         End
         Begin VB.ComboBox cboGral 
            Height          =   315
            Index           =   0
            Left            =   2640
            TabIndex        =   22
            Top             =   1800
            Width           =   3015
         End
         Begin VB.CheckBox chkgral 
            Caption         =   "Cerrar Listados de Documentos despues de Modificación"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   1320
            Width           =   5700
         End
         Begin VB.CheckBox chkgral 
            Caption         =   "Cargar los números de comprobante en cada Inicio"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   960
            Width           =   5490
         End
         Begin VB.CheckBox chkgral 
            Caption         =   "Ver la ventana de Información del cliente"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   600
            Width           =   5640
         End
         Begin VB.Label lblGral 
            AutoSize        =   -1  'True
            Caption         =   "> Lista de Precios por defecto:"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   25
            Top             =   2220
            Width           =   2160
         End
         Begin VB.Label lblGral 
            Caption         =   "> Documento por Defecto :"
            Height          =   210
            Index           =   0
            Left            =   150
            TabIndex        =   24
            Top             =   1860
            Width           =   2025
         End
         Begin VB.Label lblremitogral 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Caption         =   "Formulario de Remito"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   30
            TabIndex        =   14
            Top             =   120
            Width           =   5805
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Ejemplo 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   8
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Ejemplo 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   7
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample2 
         Caption         =   "Ejemplo 2"
         Height          =   1785
         Left            =   645
         TabIndex        =   6
         Top             =   300
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Aplicar"
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4455
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2490
      TabIndex        =   0
      Top             =   4455
      Width           =   1095
   End
End
Attribute VB_Name = "frmConfigRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cboGral_GotFocus(Index As Integer)
On Error Resume Next

    Select Case Index
    
        Case 0
            With cboGral(Index)
                .Clear
                Call .AddItem("Factura", 0)
                Call .AddItem("Remito", 1)
                Call .AddItem("Presupuesto", 2)
                Call .AddItem("Documento", 3)
                Call .AddItem("nota de Crédito", 4)
                Call .AddItem("nota de Débito", 5)
            End With
        
        Case 1
            Call CargarCombo("Listas", "Lista", cboGral(1), True)
    
    End Select
    
    
If Err Then GrabarLog "cboGral_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdApply_Click()
Dim i As Integer
On Error Resume Next

    For i = 0 To 6
        
        If i <= 3 Then
            ConfigRemito(i) = CBool(chkgral(i).Value)
        Else
            ConfigRemito(i) = CBool(chkalarma(i - 4).Value)
        End If
    
    Next
    
    Call SaveConfigRemito

If Err Then GrabarLog "cmdApply_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdCancel_Click()
On Error Resume Next

    Unload Me

If Err Then GrabarLog "cmdCancel_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdOK_Click()
On Error Resume Next

    cmdApply_Click
    Unload Me

If Err Then GrabarLog "cmdApply_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
Dim i As Integer
On Error Resume Next

    'Cargar Informacion de la Matriz a los Checkbox
    For i = 0 To 6
        
        If i <= 3 Then
            chkgral(i).Value = ConfigRemito(i) * -1
        Else
            chkalarma(i - 4).Value = ConfigRemito(i) * -1
        End If
    
    Next
    
    'centrar el formulario
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
