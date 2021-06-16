VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Acerca de MiApl"
   ClientHeight    =   3555
   ClientLeft      =   2340
   ClientTop       =   1815
   ClientWidth     =   5940
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5577.967
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   2640
      Width           =   1500
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Info. del sistema..."
      Height          =   345
      Left            =   4320
      TabIndex        =   1
      Top             =   3075
      Width           =   1485
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "Título de la aplicación"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   4
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1697.936
      Y2              =   1697.936
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión"
      Height          =   225
      Left            =   1050
      TabIndex        =   5
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "Advertencia: W.S.F. No se responsabiliza por el uso incorrecto de esta aplicación."
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   240
      TabIndex        =   3
      Top             =   2625
      Width           =   3870
   End
   Begin WGestion.AlphaIcon AlphaIcon1 
      Height          =   720
      Left            =   180
      Top             =   120
      Width           =   720
      _ExtentX        =   1270
      _ExtentY        =   1270
      bvData          =   "frmAbout.frx":0000
      bData           =   -1  'True
      IconHeight      =   48
      IconWidth       =   48
      Stretch         =   0   'False
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Opciones de seguridad de clave del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos ROOT de clave del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en valor nulo
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx _
                Lib "advapi32" _
                Alias "RegOpenKeyExA" (ByVal hKey As Long, _
                                       ByVal lpSubKey As String, _
                                       ByVal ulOptions As Long, _
                                       ByVal samDesired As Long, _
                                       ByRef phkResult As Long) As Long

Private Declare Function RegQueryValueEx _
                Lib "advapi32" _
                Alias "RegQueryValueExA" (ByVal hKey As Long, _
                                          ByVal lpValueName As String, _
                                          ByVal lpReserved As Long, _
                                          ByRef lpType As Long, _
                                          ByVal lpData As String, _
                                          ByRef lpcbData As Long) As Long

Private Declare Function RegCloseKey _
                Lib "advapi32" (ByVal hKey As Long) As Long
Private Sub cmdOK_Click()
    Unload Me
End Sub
Private Sub cmdSysInfo_Click()
    Call StartSysInfo
End Sub
Private Sub Form_Load()
On Error Resume Next
    

    With Me
        .Caption = "Acerca de WGestion"
        .lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
        .lblTitle.Caption = "WGestion" & " - " & "Sistema de Gestión Comercial "
        .lblDescription.Caption = vParametro
    End With
    
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Function GetKeyValue(KeyRoot As Long, _
                            KeyName As String, _
                            SubKeyRef As String, _
                            ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim RC As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador de una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de datos de una clave de Registro
    Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de variable de clave de Registro
    '------------------------------------------------------------
    ' Abrir clave de registro bajo KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir clave de Registro
    
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Error de controlador...
    
    tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
    KeyValSize = 1024                                       ' Marcar tamaño de variable
    
    '------------------------------------------------------------
    ' Obtener valor de clave de Registro...
    '------------------------------------------------------------
    RC = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Obtener o crear valor de clave
                        
    If (RC <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agregar cadena terminada en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Encontrado valor nulo, se va a quitar de la cadena
    Else                                                    ' En WinNT las cadenas no terminan en valor nulo...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se ha encontrado valor nulo, sólo se va a extraer la cadena
    End If

    '------------------------------------------------------------
    ' Determinar tipo de valor de clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Buscar tipos de datos...

        Case REG_SZ                                             ' Tipo de datos String de clave de Registro
            KeyVal = tmpVal                                     ' Copiar valor de cadena

        Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro

            For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
                KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor carácter a carácter
            Next

            KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a cadena
    End Select
    
    GetKeyValue = True                                      ' Se ha devuelto correctamente
    RC = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:                              ' Borrar después de que se produzca un error...
    KeyVal = ""                                             ' Establecer valor a cadena vacía
    GetKeyValue = False                                     ' Fallo de retorno
    RC = RegCloseKey(hKey)                                  ' Cerrar clave de Registro
End Function
Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim RC As Long
    Dim SysInfoPath As String
    
    ' Intentar obtener ruta de acceso y nombre del programa de Info. del sistema a partir del Registro...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Intentar obtener sólo ruta del programa de Info. del sistema a partir del Registro...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then

        ' Validar la existencia de versión conocida de 32 bits del archivo
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
            ' Error: no se puede encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If

        ' Error: no se puede encontrar la entrada del Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "La información del sistema no está disponible en este momento", vbOKOnly
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

    If Err Then GrabarLog "Form_unload", Err.Number & " " & Err.Description, Me.Name
End Sub
