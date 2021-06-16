VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmBorrarBases 
   Caption         =   "Borrado de Datos Almacenados"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   4620
   Begin VB.Frame Frame9 
      Caption         =   "Resultados"
      Height          =   1905
      Left            =   4680
      TabIndex        =   27
      Top             =   3900
      Width           =   3975
      Begin VB.ListBox l 
         Height          =   1425
         Left            =   180
         TabIndex        =   28
         Top             =   300
         Width           =   3645
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Depuración de datos"
      ForeColor       =   &H00800000&
      Height          =   3645
      Left            =   4680
      TabIndex        =   18
      Top             =   180
      Width           =   4005
      Begin ComctlLib.ProgressBar b1 
         Height          =   225
         Left            =   390
         TabIndex        =   23
         Top             =   2160
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.CommandButton cmdBorraFdetalles 
         Caption         =   "CtaCte - No Imputar = -1"
         Height          =   375
         Left            =   420
         TabIndex        =   22
         Top             =   2910
         Width           =   3225
      End
      Begin VB.CommandButton cmdCommand2 
         Caption         =   "Borra Fdetalles y Facturas Pagas"
         Height          =   405
         Left            =   390
         TabIndex        =   21
         Top             =   1740
         Width           =   3225
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   345
         Left            =   1770
         TabIndex        =   19
         Top             =   330
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65077249
         CurrentDate     =   39822
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   345
         Left            =   1770
         TabIndex        =   20
         Top             =   750
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   609
         _Version        =   393216
         Format          =   65077249
         CurrentDate     =   39822
      End
      Begin ComctlLib.ProgressBar b2 
         Height          =   225
         Left            =   420
         TabIndex        =   24
         Top             =   3300
         Width           =   3225
         _ExtentX        =   5689
         _ExtentY        =   397
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblFechaDesde 
         Caption         =   ":: Fecha hasta:"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   810
         Width           =   1245
      End
      Begin VB.Label lblLabel1 
         Caption         =   ":: Fecha desde:"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   390
         Width           =   1245
      End
   End
   Begin VB.Frame Frame4 
      Height          =   645
      Left            =   240
      TabIndex        =   7
      Top             =   4560
      Width           =   4005
      Begin VB.OptionButton o 
         Caption         =   "Borrar Todos los datos del programa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   180
         TabIndex        =   8
         Top             =   180
         Width           =   3555
      End
   End
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "Efectual la operación de BORRADO !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   210
      TabIndex        =   6
      Top             =   5370
      Width           =   4095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Información que desea eliminar :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4245
      Left            =   90
      TabIndex        =   0
      Top             =   180
      Width           =   4515
      Begin VB.Frame Frame7 
         Height          =   30
         Left            =   60
         TabIndex        =   17
         Top             =   3630
         Width           =   4395
      End
      Begin VB.Frame Frame6 
         Height          =   30
         Left            =   60
         TabIndex        =   16
         Top             =   3120
         Width           =   4395
      End
      Begin VB.Frame Frame5 
         Height          =   30
         Left            =   60
         TabIndex        =   15
         Top             =   2310
         Width           =   4395
      End
      Begin VB.Frame Frame3 
         Height          =   30
         Left            =   60
         TabIndex        =   14
         Top             =   750
         Width           =   4395
      End
      Begin VB.Frame Frame2 
         Height          =   30
         Left            =   60
         TabIndex        =   13
         Top             =   1500
         Width           =   4395
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los movimientos de Caja."
         Height          =   255
         Index           =   8
         Left            =   300
         TabIndex        =   12
         Top             =   3750
         Width           =   3495
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los cheques."
         Height          =   255
         Index           =   7
         Left            =   300
         TabIndex        =   11
         Top             =   3240
         Width           =   3495
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar Cuentas Corrientes Proveedores."
         Height          =   255
         Index           =   6
         Left            =   300
         TabIndex        =   10
         Top             =   2730
         Width           =   3495
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los documentos de Ventas."
         Height          =   255
         Index           =   5
         Left            =   300
         TabIndex        =   9
         Top             =   1680
         Width           =   4065
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar Cuentas Corrientes Clientes."
         Height          =   255
         Index           =   4
         Left            =   300
         TabIndex        =   5
         Top             =   2430
         Width           =   3135
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los documentos de  Compras."
         Height          =   255
         Index           =   3
         Left            =   300
         TabIndex        =   4
         Top             =   1980
         Width           =   3825
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los Proveedores. "
         Height          =   255
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los Clientes."
         Height          =   255
         Index           =   1
         Left            =   300
         TabIndex        =   2
         Top             =   900
         Width           =   4035
      End
      Begin VB.CheckBox c 
         Caption         =   "Eliminar todos los Artículos."
         Height          =   255
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   480
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmBorrarBases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub barti()
    On Error Resume Next
    
    BorrarBase "articulos", pathDBMySQL

    If Err Then GrabarLog "barti", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub bcaja()
    On Error Resume Next
    
    BorrarBase "caja", pathDBMySQL

    If Err Then GrabarLog "bcaja", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub bcheques()
    On Error Resume Next
    
    BorrarBase "cheques", pathDBMySQL

    If Err Then GrabarLog "bcheques", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub bcliente()
    On Error Resume Next
    
    BorrarBase "Clientes", pathDBMySQL

    If Err Then GrabarLog "bcliente", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub bcompras()
    On Error Resume Next
    
    BorrarBase "PFactura", pathDBMySQL
    BorrarBase "PFdetalle", pathDBMySQL
    BorrarBase "IvaFacturaCompra", pathDBMySQL

    If Err Then GrabarLog "bcompras", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub bctasclientes()
    On Error Resume Next
    
    BorrarBase "cuentascorrientes", pathDBMySQL

    If Err Then GrabarLog "bctasclientes", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub bctasproveedor()
    On Error Resume Next
    
    BorrarBase "pcuentascorrientes", pathDBMySQL

    If Err Then GrabarLog "bctasproveedor", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub bProveedor()
    On Error Resume Next
    
    BorrarBase "proveedores", pathDBMySQL

    If Err Then GrabarLog "bProveedor", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub bventas()
    On Error Resume Next
    
    BorrarBase "Factura", pathDBMySQL
    BorrarBase "FDetalle", pathDBMySQL
    BorrarBase "IvaFacturaVenta", pathDBMySQL

    If Err Then GrabarLog "bventas", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdEjecutar_Click()
On Error Resume Next

    MousePointer = vbHourglass

    If MsgBox("Está realmente seguro de borrar los datos seleccionados ?", vbYesNo, "Consulta ...") = vbYes Then
    
        If c(0).Value = 1 Then barti
    
        If c(1).Value = 1 Then bcliente
        If c(2).Value = 1 Then bProveedor
    
        If c(5).Value = 1 Then bventas
        If c(3).Value = 1 Then bcompras
    
        If c(4).Value = 1 Then bctasclientes
        If c(6).Value = 1 Then bctasproveedor
    
        If c(7).Value = 1 Then bcheques
    
        If c(8).Value = 1 Then bcaja
   
    End If

    MousePointer = vbDefault

    MsgBox "Los datos fueron depurados correctamente.", vbInformation, "Mensaje..."

If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
On Error Resume Next
    
    With Me
        .Show
        .Top = 300
        .Left = 2000
        .Height = 6240
        .Width = 4750
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub o_Click()
On Error Resume Next

    Dim i As Integer

    If o.Value Then

        For i = 0 To 8
            c(i).Value = 1
        Next

    Else

        For i = 0 To 8
            c(i).Value = 0
        Next

    End If

If Err Then GrabarLog "o_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
