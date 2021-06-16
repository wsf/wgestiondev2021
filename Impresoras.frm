VERSION 5.00
Begin VB.Form frmChangePrinter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambiar Impresora Predeterminada"
   ClientHeight    =   3150
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   5490
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdcambiar 
      Caption         =   "Cambiar Impresora"
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   2640
      Width           =   2175
   End
   Begin VB.ListBox lstImpresoras 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label lblnuevaimpresora 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "> Impresora por Defecto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   2
      Top             =   2160
      Width           =   2205
   End
End
Attribute VB_Name = "frmChangePrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const HWND_BROADCAST As Long = &HFFFF&
Private Const WM_WININICHANGE As Long = &H1A

Private Declare Function GetProfileString Lib "kernel32" _
   Alias "GetProfileStringA" _
  (ByVal lpAppName As String, _
   ByVal lpKeyName As String, _
   ByVal lpDefault As String, _
   ByVal lpReturnedString As String, _
   ByVal nSize As Long) As Long

Private Declare Function WriteProfileString Lib "kernel32" _
   Alias "WriteProfileStringA" _
  (ByVal lpszSection As String, _
   ByVal lpszKeyName As String, _
   ByVal lpszString As String) As Long
   
Private Declare Function SendNotifyMessage Lib "user32" _
   Alias "SendNotifyMessageA" _
  (ByVal hwnd As Long, _
   ByVal msg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
Private Sub cmdcambiar_Click()
On Error Resume Next

   Call SetDefaultPrinterWinNT
   
   lblnuevaimpresora = "> Impresora por Defecto: " & lstImpresoras.Text
    
    If Err.Number = 0 Then
       Unload Me
   Else
        MsgBox lblnuevaimpresora.Caption & " no puede ser seteada como predeterminada."
        Exit Sub
   End If

If Err Then GrabarLog "cmdcambiar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Load()
On Error Resume Next

    ProfileLoadWinIniList lstImpresoras, "PrinterPorts"
    cmdcambiar.Enabled = False
    frmSaldosClientes.cmdchangeprinter.Enabled = False
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)
    frmSaldosClientes.cmdchangeprinter.Enabled = True

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GetDriverAndPort(ByVal Buffer As String, _
                             DriverName As String, _
                             PrinterPort As String)

   Dim posDriver As Long
   Dim posPort As Long
   DriverName = ""
   PrinterPort = ""

  'The driver name is first in the string
  'terminated by a comma
   posDriver = InStr(Buffer, ",")
   
   If posDriver > 0 Then

     'Strip out the driver name
      DriverName = Left(Buffer, posDriver - 1)

     'The port name is the second entry after
     'the driver name separated by commas.
      posPort = InStr(posDriver + 1, Buffer, ",")

      If posPort > 0 Then
      
        'Strip out the port name
         PrinterPort = Mid(Buffer, posDriver + 1, posPort - posDriver - 1)
         
       End If
   End If
   
End Sub
Private Sub lstImpresoras_Click()
On Error Resume Next

   cmdcambiar.Enabled = lstImpresoras.ListIndex > -1
   
If Err Then GrabarLog "lstImpresoras_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function ProfileLoadWinIniList(lst As ListBox, _
                                       lpSectionName As String) As Long

'Load the listbox data from win.ini.

   Dim success As Long
   Dim nSize As Long
   Dim lpKeyName As String
   Dim ret As String
  
  'call the API passing null as the parameter
  'for the lpKeyName parameter. This causes
  'the API to return a list of all keys under
  'that section. Pad the passed string large
  'enough to hold the data. Adjust to suit.
   ret = Space$(8102)
   nSize = Len(ret)
   success = GetProfileString(lpSectionName, _
                              vbNullString, _
                              "", _
                              ret, _
                              nSize)
   
  'The returned string is a null-separated
  'list terminated by a pair of null characters.
   If success Then
    
     'trim terminating null and trailing spaces
      ret = Left$(ret, success)
      
        'with the resulting string,
        'extract each element
         Do Until ret = ""
      
           'strip off an item and
           'add the item to the listbox
            lpKeyName = StripNulls(ret)
            lst.AddItem lpKeyName
      
         Loop
  
   End If
  
  'return the number of items as an
  'indicator of success
   ProfileLoadWinIniList = lst.ListCount

End Function
Private Sub SetDefaultPrinterWinNT()

   Dim Buffer As String
   Dim DeviceName As String
   Dim DriverName As String
   Dim PrinterPort As String
   Dim PrinterName As String
   Dim r As Long
   
   If lstImpresoras.ListIndex > -1 Then
        
     'Get the printer information for the currently selected
     'printer in the list. The information is taken from the
     'WIN.INI file.
      Buffer = Space(1024)
      PrinterName = lstImpresoras.List(lstImpresoras.ListIndex)
      
      Call GetProfileString("PrinterPorts", _
                             PrinterName, "", _
                             Buffer, Len(Buffer))
 
     'Parse the driver name and port name out of the buffer
      GetDriverAndPort Buffer, DriverName, PrinterPort

      If (Len(DriverName) > 0) And (Len(PrinterPort) > 0) Then
         SetDefPrinter PrinterName, DriverName, PrinterPort
      End If
      
    End If
    
End Sub
Private Sub SetDefPrinter(ByVal PrinterName As String, _
                          ByVal DriverName As String, _
                          ByVal PrinterPort As String)
   Dim DeviceLine As String
    
  'rebuild a valid device line string
   DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
   
  'Store the new printer information in the
  '[WINDOWS] section of the WIN.INI file for
  'the DEVICE= item
   Call WriteProfileString("windows", "Device", DeviceLine)
    
  'Cause all applications to reload the INI file
   Call SendNotifyMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
    
End Sub
Private Function StripNulls(startstr As String) As String

 'Take a string separated by chr$(0)
 'and split off 1 item, shortening the
 'string so next item is ready for removal.
  Dim pos As Long

  pos = InStr(startstr$, Chr$(0))
  
  If pos Then
      
      StripNulls = Mid$(startstr, 1, pos - 1)
      startstr = Mid$(startstr, pos + 1, Len(startstr))
    
  End If

End Function
