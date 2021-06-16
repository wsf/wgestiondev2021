VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmAyuda 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda de WGestion"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   8595
      Left            =   60
      TabIndex        =   0
      Top             =   300
      Width           =   11175
      ExtentX         =   19711
      ExtentY         =   15161
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblVolverHacia 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Volver hacia atrás con la tecla <Backspace> o clic en <botón derecho> y seleccionar atrás. "
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   30
      Width           =   9795
   End
End
Attribute VB_Name = "frmAyuda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public v As String

Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2 - 1000

wb.Navigate v
End Sub

