VERSION 5.00
Object = "{50BF2256-701F-46F2-8ADB-2202CE6922BC}#1.0#0"; "KlexGrid.ocx"
Begin VB.Form frmConsultas 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Control de inconsistencia de datos"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5835
   ScaleWidth      =   12720
   Begin Grid.KlexGrid grid 
      Height          =   5325
      Left            =   30
      TabIndex        =   0
      Top             =   480
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   9393
      EnterKeyBehaviour=   0
      BackColorAlternate=   8421504
      GridLinesFixed  =   2
      BackColorBkg    =   16777215
      BackColorFixed  =   255
      BorderStyle     =   0
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
      MouseIcon       =   "frmControlGral.frx":0000
      Rows            =   10
   End
   Begin VB.Label etiqueta 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12675
   End
End
Attribute VB_Name = "frmConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim bcontrol As ADODB.Recordset

Public Sub buscar(vsql As String)

Set bcontrol = New ADODB.Recordset
  With bcontrol
        If .State = 1 Then .Close

        .CursorLocation = adUseServer
        Call .Open(vsql, ConnDDBB, adOpenDynamic, adLockPessimistic)
  End With


    Set Me.grid.Recordset = bcontrol
End Sub
   
  
