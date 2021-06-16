VERSION 5.00
Begin VB.Form frmAlert 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3720
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   600
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2160
      Top             =   1080
   End
   Begin VB.PictureBox picBackground 
      AutoRedraw      =   -1  'True
      BackColor       =   &H0000FFFF&
      Height          =   1965
      Left            =   0
      ScaleHeight     =   1905
      ScaleWidth      =   3645
      TabIndex        =   0
      Top             =   0
      Width           =   3705
      Begin VB.Label lblAlert 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Alert Message"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1155
         Left            =   60
         TabIndex        =   1
         Top             =   180
         Width           =   3435
      End
   End
   Begin VB.Timer tmrAlert 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2160
      Top             =   120
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' API Declarations
Private Declare Function GetSystemMetrics& _
                Lib "user32" (ByVal nIndex As Long)

Private Declare Function sndPlaySound _
                Lib "WINMM.DLL" _
                Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
                                       ByVal uFlags As Long) As Long

' Constants
Const SM_CXFULLSCREEN = 16   ' Width of window client area
Const SM_CYFULLSCREEN = 20   ' Height of window client area
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_NODEFAULT = &H2
Const SND_LOOP = &H8
Const SND_NOSTOP = &H10

' Declarations
Private ClsGradient As New CGradient

Private fX As Long

Private fY As Long

Private lngScaleX As Long

Private lngScaleY As Long

Private AlertIndex As Long
Dim vtop As Long
Dim vleft As Long
Public Sub DisplayAlert(MessageText As String, _
                        Duration As Long)

    Dim wflags As Long, X As Long

    ' Increase the alert count
    AlertCount = AlertCount + 1
    AlertIndex = AlertCount

    ' Set the message
    lblAlert.Caption = MessageText

    ' Set the duration
    tmrAlert.Interval = Duration

    ' Get the system metrics we need
    fX = GetSystemMetrics(SM_CXFULLSCREEN)
    fY = GetSystemMetrics(SM_CYFULLSCREEN)
    lngScaleX = Me.width - Me.ScaleWidth
    lngScaleY = Me.height - Me.ScaleHeight
    
    ' Size the form
    Me.height = 1965
    Me.width = picBackground.width + lngScaleX
    Me.Left = fX * Screen.TwipsPerPixelX - Me.width - vleft
    Me.Top = (fY * Screen.TwipsPerPixelY) - ((picBackground.height + lngScaleY) * (AlertCount - 1)) + 160
    Me.Show
    
    ' Play sound
    wflags = SND_ASYNC Or SND_NODEFAULT
    If MessageText = ("No tiene Tareas/Notas para hoy") Then
        X = sndPlaySound(App.Path & "\newalert.wav", wflags)
    Else
        X = sndPlaySound(App.Path & "\notify.wav", wflags)
    End If
    ' Draw the gradient background
    With ClsGradient
        .Angle = 150 'Angulo del efecto entre color 1 y color 2
        .Color1 = RGB(125, 125, 125) 'RGB(255, 255, 255) 'Blanco
        .Color2 = RGB(255, 255, 255) 'RGB(199, 199, 199) 'Gris
        .Draw picBackground
    End With

    picBackground.Refresh

    ' Open the alert box
    tmrOpen.Enabled = True

End Sub

Private Sub Form_Load()
    vtop = 5840
    vleft = 40
End Sub

Private Sub lblAlert_Click()
   frmAgenda.tab_agenda.tab = 2
End Sub

Private Sub lblAlert_MouseMove(Button As Integer, _
                               Shift As Integer, _
                               X As Single, _
                               Y As Single)

    ' Show as hyperlink
    If lblAlert.FontUnderline = False Then
        lblAlert.FontUnderline = True
        lblAlert.ForeColor = RGB(0, 0, 255)
    End If

End Sub

Private Sub picBackground_MouseMove(Button As Integer, _
                                    Shift As Integer, _
                                    X As Single, _
                                    Y As Single)

    ' Show text
    If lblAlert.FontUnderline = True Then
        lblAlert.FontUnderline = False
        lblAlert.ForeColor = &H0
    End If

End Sub

Private Sub tmrAlert_Timer()
    ' Alert was displayed, now close it
    tmrAlert.Enabled = False
    tmrClose.Enabled = True
End Sub

Private Sub tmrClose_Timer()
    Dim curHeight As Long
    curHeight = Me.height

    If curHeight > 120 Then
        Me.height = curHeight - 30
        Me.Top = vtop + 30   'Me.Top + 30
    Else

        ' Close form
        If AlertCount = AlertIndex Then AlertCount = 0
        Unload Me
    End If

End Sub
Private Sub tmrOpen_Timer()
    Dim curHeight As Long
    Dim newHeight As Long
    curHeight = Me.height

    If curHeight < picBackground.height + lngScaleY Then
        newHeight = curHeight + 30

        If newHeight > picBackground.height + lngScaleY Then newHeight = picBackground.height + lngScaleY
        Me.height = Me.height + (newHeight - curHeight)
        Me.Top = vtop - (newHeight - curHeight) 'Me.Top - (newHeight - curHeight)
    Else
        tmrOpen.Enabled = False
        tmrAlert.Enabled = True
    End If

End Sub


