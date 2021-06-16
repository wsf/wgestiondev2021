Attribute VB_Name = "m_odbc"
Option Explicit

'Constantes
Private Const ODBC_ADD_DSN = 1          'Nuevo DSN
Private Const ODBC_CONFIG_DSN = 2       'Modificar DSN
Private Const ODBC_REMOVE_DSN = 3       'Eliminar DSN
Private Const ODBC_ADD_SYS_DSN = 4      'Nuevo DSN de sistema
Private Const ODBC_CONFIG_SYS_DSN = 5   'Modificar DSN de sistema
Private Const ODBC_REMOVE_SYS_DSN = 6   'Eliminar DSN de sistema
Private Const vbAPINull As Long = 0     'Null Pointer
Private Const SQL_SUCCESS As Long = 0
Private Const SQL_FETCH_NEXT As Long = 1

'Declaración de funciones de API
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Private Declare Function SQLDataSources Lib "ODBC32.DLL" (ByVal henv As Long, ByVal fDirection As Integer, ByVal szDSN As String, ByVal cbDSNMax As Integer, pcbDSN As Integer, ByVal szDescription As String, ByVal cbDescriptionMax As Integer, pcbDescription As Integer) As Integer
Private Declare Function SQLAllocEnv Lib "ODBC32.DLL" (Env As Long) As Integer

'Constantes
'ruta hasta el servidor (ip/nombre/ruta)
Private Const C_Server = ""
'usuario
Private Const C_User = ""
'contraseña
Private Const C_Pass = ""
'base de datos
Private Const C_BD = ""
'puerto
Private Const C_Port = 3306
'Nombre ODBC de MySql
'(si no tienes ninguno instalado http://dev.mysql.com/downloads/connector/odbc/5.1.html)
'Public Const C_MYSQL_ODBC = "MySQL ODBC 5.1 Driver"
Public Const C_MYSQL_ODBC = "MySQL ODBC 3.51 Driver"
Public Function BorrarDSN(sDSN As String, sDriver As String, Optional sHwnd As Long = vbAPINull) As Boolean
    Dim sAtributos As String
    
    If ExisteDSN(sDSN) Then
        sAtributos = "DSN=" & sDSN & Chr(0)
        BorrarDSN = CBool(SQLConfigDataSource(sHwnd, ODBC_REMOVE_SYS_DSN, sDriver, sAtributos))
    Else
    
        BorrarDSN = False
    End If
End Function
Public Function CrearDSN(sDSN As String, sDriver As String, sAtributos As String, Optional sHwnd As Long = vbAPINull) As Boolean
    CrearDSN = CBool(SQLConfigDataSource(sHwnd, ODBC_ADD_SYS_DSN, sDriver, sAtributos))
End Function
Public Function ExisteDSN(sDSN As String) As Boolean
    Dim i As Integer, j As Integer
    Dim sDSNItem As String * 1024
    Dim sDRVItem As String * 1024
    Dim sDSNActual As String
    Dim sDRV As String
    Dim iDSNLen As Integer
    Dim iDRVLen As Integer
    Dim lHenv As Long
    Dim DSNLISTA(100)
    ExisteDSN = False
    For j = 1 To 52
        DSNLISTA(j) = ""
    Next j
    
    j = 1
    If SQLAllocEnv(lHenv) <> -1 Then
        Do Until i <> SQL_SUCCESS
            sDSNItem = Space(1024)
            sDRVItem = Space(1024)
            i = SQLDataSources(lHenv, SQL_FETCH_NEXT, sDSNItem, 1024, iDSNLen, sDRVItem, 1024, iDRVLen)
            sDSNActual = VBA.Left(sDSNItem, iDSNLen)
            sDRV = VBA.Left(sDRVItem, iDRVLen)
            If sDSN <> Space(iDSNLen) Then
                DSNLISTA(j) = sDSN
                If UCase(sDSN) = UCase(sDSNActual) Then
                    ExisteDSN = True
                    Exit Do
                End If
            End If
        Loop
    End If
End Function
Public Function IniciaDSN(sDSNname As String) As Boolean
    If ExisteDSN(sDSNname) = True Then
        If BorrarDSN(sDSNname, C_MYSQL_ODBC) = False Then
            IniciaDSN = False
            Exit Function
        End If
    End If
    IniciaDSN = MySQLCrearDSN(sDSNname)
End Function
Public Function MySQLCrearDSN(sDSN As String, _
 Optional sServer As String = C_Server, Optional sBD As String = C_BD, _
 Optional sUser As String = C_User, Optional sPass As String = C_Pass, _
 Optional sPort As Integer = C_Port) As Boolean
 
    Dim sDriver As String
    Dim sAtributos As String

    sDriver = C_MYSQL_ODBC
    sAtributos = "DSN=" & sDSN & Chr(0)
    sAtributos = sAtributos & "SERVER=" & vConfigGral.vDireccionDB & Chr(0)
    sAtributos = sAtributos & "PORT=" & sPort & Chr(0)
    sAtributos = sAtributos & "DATABASE=" & vConfigGral.vEmpresa & Chr(0)
    sAtributos = sAtributos & "USER=" & vConfigGral.vUserDB & Chr(0)
    sAtributos = sAtributos & "PASSWORD=" & vConfigGral.vPassDB & Chr(0)
    sAtributos = sAtributos & "OPTION=3" & Chr(0)
    
    If ExisteDSN(sDSN) Then
        Call BorrarDSN(sDSN, sDriver)
    End If
    
    MySQLCrearDSN = CrearDSN(sDSN, sDriver, sAtributos)

End Function
