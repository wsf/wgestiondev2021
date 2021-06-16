VERSION 5.00
Begin VB.Form frmWebCam 
   Caption         =   "Form1"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12975
   LinkTopic       =   "Form1"
   ScaleHeight     =   8235
   ScaleWidth      =   12975
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   420
      Left            =   4545
      TabIndex        =   4
      Top             =   45
      Width           =   1410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   420
      Left            =   3060
      TabIndex        =   3
      Top             =   45
      Width           =   1410
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command1"
      Height          =   420
      Left            =   1575
      TabIndex        =   2
      Top             =   45
      Width           =   1410
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   420
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      Height          =   7620
      Left            =   90
      ScaleHeight     =   7560
      ScaleWidth      =   12780
      TabIndex        =   0
      Top             =   495
      Width           =   12840
   End
End
Attribute VB_Name = "frmWebCam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temp As Long

' botón que inicia la captura
'''''''''''''''''''''''''''''''''''''''
Private Sub Command1_Click()
Dim temp As Long

  hwdc = capCreateCaptureWindow("CapWindow", ws_child Or ws_visible, _
                                    0, 0, 320, 240, Picture1.hwnd, 0)
  If (hwdc <> 0) Then
    temp = SendMessage2(hwdc, wm_cap_driver_connect, 0, 0)
    temp = SendMessage2(hwdc, wm_cap_set_preview, 1, 0)
    temp = SendMessage2(hwdc, WM_CAP_SET_PREVIEWRATE, 30, 0)
    temp = SendMessage2(hwdc, WM_CAP_SET_SCALE, True, 0)
    'esto hace que la imagen recibida por el dispositivo se ajuste
    'al tamaño de la ventana de captura (justo lo que yo buscaba)
    DoEvents
    startcap = True
    Else
    MsgBox "No hay Camara Web", 48, "Error"
  End If

End Sub

' botón para detener la captura
'''''''''''''''''''''''''''''''''''''''
Private Sub Command2_Click()
    
    temp = DestroyWindow(hwdc)
    If startcap = True Then
        temp = SendMessage(hwdc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
        DoEvents
        startcap = False
    End If

End Sub

' Botón que abre el dialogo de formato
''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command3_Click()
        If startcap = True Then
            
            temp = SendMessage(hwdc, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
            DoEvents
        End If
End Sub
' Mostrar dialogo de Configuracion de la WebCam
''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Command4_Click()
 Dim temp As Long
    If startcap = True Then
        temp = SendMessage(hwdc, WM_CAP_DLG_VIDEOCONFIG, 0&, 0&)
        DoEvents
    End If
End Sub

Private Sub Form_Load()
    Command1.Caption = "Iniciar"
    Command2.Caption = "Detener"
    Command3.Caption = "Formato"
    Command4.Caption = "Configurar"
    Me.Caption = "Capturador de Web Cam"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29
End Sub

Private Sub Form_Unload(Cancel As Integer)

    temp = DestroyWindow(hwdc)
    If startcap = True Then
        temp = SendMessage(hwdc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
        DoEvents
        startcap = False
    End If
End Sub
 
 



