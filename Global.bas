Attribute VB_Name = "Global"
Option Explicit
Public errexit As Boolean

Public m_SortColumn As Integer

Public m_SortOrder As SortSettings

Public vrollbk_nrointerno As Long
Public vrollbk_nroasiento As Long
Public vrollbk  As Boolean

Public validado As Boolean

Public Const vmaximizar = 0 ' tamaño de las ventanas

Public Const vCodigoChequesEntregados = "098"

Public vAsientoAutomatico As Boolean


Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccesstype As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet" Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, ByVal lpszUrl As String, ByVal lpszHeaders As String, ByVal dwHeadersLength As Long, ByVal dwFlags As Long, ByVal dwContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Public Declare Function EbExecuteLine Lib "vba6.dll" (ByVal pStringToExec As Long, ByVal Foo1 As Long, ByVal Foo2 As Long, ByVal fCheckOnly As Long) As Long
'Global Cn As ADODB.Connection

Public Const ID_TAB_ICON = 1040

Public Const ID_Registros_1 = 701
Public Const ID_Registros_2 = 702
Public Const ID_Registros_3 = 703
Public Const ID_Registros_4 = 704
Public Const ID_Registros_5 = 705
Public Const ID_Registros_6 = 706

Public Const ID_Configuracion_1 = 801
Public Const ID_Configuracion_2 = 802
Public Const ID_Configuracion_3 = 803
Public Const ID_Configuracion_4 = 804
Public Const ID_Configuracion_5 = 805
Public Const ID_Configuracion_6 = 806

Public Const ID_Tab_Importar = 0
Public Const ID_Tab_Exportar = 1
Public Const ID_Tab_Registros = 2
Public Const ID_Tab_Configuracion = 3

    

Public Const ID_INDICATOR_CAPS = 59137
Public Const ID_INDICATOR_NUM = 59138
Public Const ID_INDICATOR_SCRL = 59139

Private Const HWND_BROADCAST As Long = &HFFFF
Private Const WM_WININICHANGE As Long = &H1A

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As String) As Long

Declare Function PathFileExists _
    Lib "shlwapi.dll" _
    Alias "PathFileExistsA" (ByVal pszPath As String) As Long

Declare Function WriteProfileString _
        Lib "kernel32" _
        Alias "WriteProfileStringA" (ByVal lpszSection As String, _
                                     ByVal lpszKeyName As String, _
                                     ByVal lpszString As String) As Long

Private Declare Function CopyFile _
                Lib "kernel32" _
                Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                                   ByVal lpNewFileName As String, _
                                   ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile _
                Lib "kernel32" _
                Alias "DeleteFileA" (ByVal lpFileName As String) As Long
                
 Public valertaModulo As String
 
Public Declare Function GetTickCount _
               Lib "kernel32.dll" () As Long

'Public vImportar As Boolean
'Public vPcServer As Boolean 'Defino si la pc es Maestro o algun puesto en red


Public Type totalesDoc

    iva105 As Double
    iva21 As Double
    iva27 As Double
    total As Double
    
End Type



Public Type dsCheques ' data set de cheques
  idCheques          As Long
  idEstadoCheque     As Integer
  fecha              As Date
  Codigo             As String ' codigo del cliente
  Nombre             As String  'nombre del cliente
  
  idBancos           As String
  idBancosCuentas    As Integer
  
  Ncheque            As String
  Firmante           As String
  CP                 As String
  FechaDeposito      As Date
  monto              As Double
  Endoso             As String
  remito             As Integer
  NroInterno         As Long
  Observaciones      As String
  FechaAcreditacion  As Date
  'Foto               As Long
  TipoMovimiento     As String
  vid                As Long
  'TimeStamp          timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP,
  idCustodia         As String  ' custodio del cheque
  
  marcainterna       As Integer
  
  sucursal           As String
  
End Type

Public Type vConfigGral
        vIdEmpresa As Long
        vempresa As String
        vIdServidor As Long
        vServidor As String
        vDireccionDB As String
        vUserDB As String
        vPassDB As String
        vIdUsuario As Long
        vUser As String
        vPass As String
        vIncluyeContabilidad As Boolean
        vEmpresaPrincipal As Boolean
        vUsarEstaEmpresa As Boolean
        vIncluyeStac As Boolean
        vIncluyeResto As Boolean
        vIncluyeTicket As Boolean
        vIncluyeCobros As Boolean
        vImpresoraSeleccionada As String
        vImprimirReciboCliente As Boolean
        vImprimirReciboProveedor As Boolean
        vClientes As String
        vremito As String
        vTipoCliente As String
        vComunaCnn As String
End Type

'Public Type EstadoDocumento
'     Pagado="Pagado" As Constants
'     Pendiente = "Pendiente"
'     Quebranto = "Quebranto"
'     Adeudado = "Adeudado"
'End Type



Public Type vImpresoras
    vNombreImpresora As String
    vModelo As String
    vModeloInterno As String
    vNroPuerto As Integer
    vEsFiscal As String
    vPorDefecto As String
End Type

Public Type vParametrosSistema
    vFormatoArticulos As Byte
    vFormatoClientes As Byte
    vFormatoEmpleados As Byte
    vFormatoProveedores As Byte
    vFechaInicio As Date
    vFechaFin As Date
End Type


Public Type vfactura
    vcodigo As String
    vnombre As String
    vdireccion As String
    vlocalidad As String
    vcuit As String
    vtotal As Double
    vSubTotal As Double
    vIva210 As Double
    viva150 As Double
    viva270 As Double
    vfecha As Date
    vsaldo As Double
    vcae As String
    vcaeVto As String
    vIva As String
End Type


Public Type vDatosEmpresa
        Nombre As String
        Alias As String
        CondicionIva As String '(cod:CondiIva-Empresa)
        cuit As String
        Direccion As String
        Localidad As String
        Telefono As String
        Email As String
        WebSite As String
        Responsable As String
        UsarNroInterno As String
End Type


'----------------------------------
Public vgidCtaCte As Long
Public vgTablaCatcte As String
'----------------------------------
Public totalesd As totalesDoc


'---------------------- variable globales para accesos a los data set  ------------
Public gbldsCheques As dsCheques
'---------------------------------------------------------------------------------

Public vConfigGral As vConfigGral
Public vDatosEmpresa As vDatosEmpresa
Public vImpresoras As vImpresoras
Public vParametrosSistema As vParametrosSistema
    
Public vParametro As String
Public ConfigRemito() As Boolean
Public vUsuarioSistema As String

'----------------- variables para acceder a las bases de datos -----------
Global ConnDDBB As ADODB.Connection
Global ConnComunaDB As ADODB.Connection
Global ConnComunaDB2 As ADODB.Connection
Global ConnComunaDB3 As ADODB.Connection



'Global Rec As ADODB.Recordset
'Global sql As String
'Global Rec2 As ADODB.Recordset
'Global sql2 As String
'---------------------------------------------------------------------------

'Password
Public sININame As String
Public key As String
Public Number As Long
Public OldOnline As Boolean
Public file_name As String
Public msgTitle As String
Public msgSubject As String
Public msgDetail As String
Public msgStandard As Boolean
Public Section As String
Public Password As String

'Variable para cambiar la impresora predeterminada con una API
Public di

' variable para mandar los totales de las factura la fondo
Public margenfactura As Integer



Public vFormulario As String

'Public gnombre, gtelefono, gdireccion As String

Public gdolar, giva As Double

'Public Arollback(100, 30)

Public gupago, guventa As String

Public vVieneBusqueda As String
Public vVuelveBusqueda As String
Public vVieneImpresion As String
Public vGrabarTabla As String
Public vVieneConcepto As String

Public vDSN As Boolean
Public vPFDetalle As Boolean
Public gsaldo, gcredito As Double
Declare Function SetLocaleInfo _
        Lib "kernel32" _
        Alias "SetLocaleInfoA" (ByVal Locale As Long, _
                                ByVal LCType As Long, _
                                ByVal lpLCData As String) As Long
Declare Function GetLocaleInfo _
        Lib "kernel32" _
        Alias "GetLocaleInfoA" (ByVal Locale As Long, _
                                ByVal LCType As Long, _
                                ByVal lpLCData As String, _
                                ByVal cchData As Long) As Long

Type NUMBERFMT
    NumDigits As Long ' número de dígitos decimales
    LeadingZero As Long ' si hay ceros iniciales en los campos decimales
    Grouping As Long ' tamaño del grupo a la izquierda del decimal
    lpDecimalSep As String ' puntero a la cadena del separador de decimales
    lpThousandSep As String ' puntero a la cadena del separador de miles
    NegativeOrder As Long ' orden de números negativos
End Type

Declare Function GetNumberFormat _
        Lib "kernel32" _
        Alias "GetNumberFormatA" (ByVal Locale As Long, _
                                  ByVal dwFlags As Long, _
                                  ByVal lpValue As String, _
                                  lpFormat As NUMBERFMT, _
                                  ByVal lpNumberStr As String, _
                                  ByVal cchNumber As Long) As Long

Public Const LOCAL_DEFAULT = &H2C0A

Public Const LOCALE_SDECIMAL = &HE

Public Const LOCALE_STHOUSAND = &HF

Public Const LOCALE_IDIGITS = &H11

Public Const LOCALE_STIMEFORMAT = &H1003

Public Const LOCALE_SSHORTDATE = &H1F

Public Const LOCALE_SLONGDATE = &H20

Public Const LOCALE_SCURRENCY = &H14

Public Const LOCALE_SMONDECIMALSEP = &H16

Public Const LOCALE_SMONTHOUSANDSEP = &H17

Public Const FMT_FECHA_CORTA As String = "dd/MM/yyyy"

Public Const FMT_FECHA_LARGA As String = "dddd, d' de 'MMMM' de 'yyyy"

Public Const FMT_HORA As String = "HH:mm:ss"

Public Const SIMB_MONEDA As String = "$"

Public Const SEP_DEC As String = "."

Public Const SEP_MILES As String = ","

' Public oAccess As Access.Application ' sacado


Public Const vCampoMovimientosCaja = "Fecha,NroInterno,Tipo,Partner,Debito,Credito,Comentario,NroCheque,Saldo,cp"


Public Sub BorrarBase(vtabla As String, _
                      cbdd As String)
    On Error Resume Next
    
    Dim oConn As New ADODB.Connection, strSQL As String
 
    With oConn
        .ConnectionString = cbdd
        .Open
        
        If .State = 0 Then
            
            MsgBox Err.Description
            Exit Sub
        
        Else
            strSQL = "DELETE FROM " & Trim(vtabla)
            
            .Execute strSQL
            .Close
        
        End If
    
    End With
    
    If Err Then
        MsgBox "Debe consultar al servicio técnico por este error:" + Chr(13) + ">Tipo de Error:" + Err.Description + Chr(13) + "> Operación: " + strSQL
        GrabarLog "BorrarBase", Err.Number & " " & Err.Description, "Global"
        'alerta "Error en " + Chr(13) + strSQL, 5000
    Else
        'alerta strSQL, 2000
    End If
End Sub
Function Borrar(AdoName As Object, mensaje As Boolean) As Boolean
    On Error Resume Next

    'Nombre del Ado
    With AdoName
        
        If Not (.Recordset.EOF = True) And Not (.Recordset.BOF = True) Then
            
            If mensaje = True Then
                If MsgBox("¿ Esta seguro que desea borrar este registro ? ", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                
                    .Recordset.Delete
                    Borrar = True
                
                Else
                
                    Borrar = False
                End If
            Else
            
                 .Recordset.Delete
                 Borrar = True
                 
            End If
        
        End If
    
    End With

    If Err Then GrabarLog "Borrar", Err.Number & " " & Err.Description, "Global"
End Function
Function CentrarFormulario(ByRef frm As Form)

    frm.Left = (Screen.Width - frm.Width) / 2
    frm.Top = (Screen.Height - frm.Height) / 2 - 1000

End Function

Function BorrarRecordset(rsRecordset As ADODB.Recordset, mensaje As Boolean) As Boolean
    On Error Resume Next

    'Nombre del Ado
    With rsRecordset
        
        If Not (.EOF = True) And Not (.BOF = True) Then
            
            If mensaje = True Then
                If MsgBox("¿ Esta seguro que desea borrar este registro ? ", vbInformation + vbYesNo, "Mensaje ...") = vbYes Then
                
                    .Delete
                    BorrarRecordset = True
                
                Else
                
                    BorrarRecordset = False
                End If
            Else
            
                 .Delete
                 BorrarRecordset = True
                 
            End If
        
        End If
    
    End With

    If Err Then GrabarLog "Borrar", Err.Number & " " & Err.Description, "Global"
End Function
Public Function CambiarCR(Optional strError As String) As Boolean
    Dim lngResu As Long
    Dim Buffer As String * 255
    
    On Error GoTo errores
    
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SSHORTDATE, FMT_FECHA_CORTA)

    If lngResu = 0 Then strError = "Error al setear fecha corta."
    
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SLONGDATE, FMT_FECHA_LARGA)

    If lngResu = 0 Then strError = "Error al setear fecha larga."
    
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SDECIMAL, SEP_DEC)

    If lngResu = 0 Then strError = "Error al setear separador de decimales."
    
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_STHOUSAND, SEP_MILES)

    If lngResu = 0 Then strError = "Error al setear separador de miles."
    
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_STIMEFORMAT, FMT_HORA)

    If lngResu = 0 Then strError = "Error al setear formato de hora."
    
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SMONDECIMALSEP, SEP_DEC)

    If lngResu = 0 Then strError = "Error al setear separador de decimales de moneda."
    
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SMONTHOUSANDSEP, SEP_MILES)
    
    If lngResu = 0 Then strError = "Error al setear separador de miles de moneda."
    lngResu = SetLocaleInfo(LOCAL_DEFAULT, LOCALE_SCURRENCY, SIMB_MONEDA)

    If lngResu = 0 Then strError = "Error al setear símbolo de moneda."
    lngResu = GetLocaleInfo(LOCAL_DEFAULT, LOCALE_SDECIMAL, Buffer, Len(Buffer))

    If Left$(Buffer, 1) = SEP_DEC Then
        lngResu = GetLocaleInfo(LOCAL_DEFAULT, LOCALE_SMONDECIMALSEP, Buffer, Len(Buffer))

        If Left$(Buffer, 1) = SEP_DEC Then
            lngResu = GetLocaleInfo(LOCAL_DEFAULT, LOCALE_STHOUSAND, Buffer, Len(Buffer))

            If Left$(Buffer, 1) = SEP_MILES Then
                lngResu = GetLocaleInfo(LOCAL_DEFAULT, LOCALE_SMONTHOUSANDSEP, Buffer, Len(Buffer))

                If Left$(Buffer, 1) = SEP_MILES Then
                    CambiarCR = (strError = vbNullString)
                End If
            End If
        End If
    End If
    
    Exit Function
errores:
    CambiarCR = False
End Function
Public Sub GuardarUDato(vtipo As Integer, vUReparto, vURepartidor, vUFecha)
On Error Resume Next
    
    Dim rsUReparto As New ADODB.Recordset
    Dim sqlUReparto As String
    
    sqlUReparto = "SELECT * FROM configura"
    
    With rsUReparto
        Call .Open(sqlUReparto, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        Select Case vtipo
        
            Case 0
                .Fields("ufreparto").Value = Trim(vUReparto)
                .Fields("urepartidor").Value = Trim(vURepartidor)
                .Fields("ufecha_reparto").Value = vUFecha
        
            Case 1
                .Fields("ultimafecha").Value = vUFecha
            
        End Select
        
        
        .Update
    End With
    
    sqlUReparto = ""
    
    rsUReparto.Close
    Set rsUReparto = Nothing
    

If Err Then GrabarLog "GuardarUReparto", Err.Number & " " & Err.Description, "Global"
End Sub
Public Sub CargarCombo(tabla As String, campo As String, combo As ComboBox, distinto As Boolean, Optional Indice As String, Optional vPathDB As String)
On Error Resume Next

    Dim rsCombo As New ADODB.Recordset, sqlCombo As String

    If distinto = True Then
        sqlCombo = "SELECT DISTINCT " & campo & " FROM " & tabla
    Else
        sqlCombo = "SELECT * FROM " & tabla
    End If
   
    With rsCombo
        If vPathDB = "" Then
            Call .Open(sqlCombo, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(sqlCombo, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        combo.Clear
        
        If Not .EOF = True Then .MoveFirst
        
        Do Until .EOF = True
            
            If Trim(Indice) = "" Then
                combo.AddItem Trim(.Fields(campo).Value)
                combo.Tag = Trim(.Fields(0).Value)
            Else
                combo.AddItem Trim(.Fields(campo).Value)
                combo.ItemData(combo.NewIndex) = Val(Trim(.Fields(Trim(Indice)).Value))
                combo.Tag = Trim(.Fields(0).Value)
            End If
            
            .MoveNext
    
        Loop
    
    End With
    
    sqlCombo = ""
    
    rsCombo.Close
    Set rsCombo = Nothing
    

If Err Then GrabarLog "CargarCombo", Err.Number & " " & Err.Description, "Global"
End Sub
Public Sub CargarComboNew(tabla As String, campo As String, combo As XtremeSuiteControls.ComboBox, distinto As Boolean, Optional Indice As String, Optional vPathDB As String)
On Error Resume Next

    Dim rsCombo As New ADODB.Recordset, sqlCombo As String

    If distinto = True Then
        sqlCombo = "SELECT DISTINCT " & campo & " FROM " & tabla
    Else
        sqlCombo = "SELECT * FROM " & tabla
    End If
   
    With rsCombo
        If vPathDB = "" Then
            Call .Open(sqlCombo, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(sqlCombo, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        combo.Clear
        
        If Not .EOF = True Then .MoveFirst
        
        Do Until .EOF = True
            
            If Trim(Indice) = "" Then
                combo.AddItem Trim(.Fields(campo).Value)
            Else
                combo.AddItem Trim(.Fields(campo).Value)
                combo.ItemData(combo.NewIndex) = Val(Trim(.Fields(Indice).Value))
            End If
            
            .MoveNext
    
        Loop
    
    End With
    
    sqlCombo = ""
    
    rsCombo.Close
    Set rsCombo = Nothing

If Err Then GrabarLog "CargarCombo", Err.Number & " " & Err.Description, "Global"
End Sub
Public Sub CargarComboNew2(tabla As String, campo1 As String, campo2 As String, combo As XtremeSuiteControls.ComboBox, distinto As Boolean, Optional Indice As String, Optional vPathDB As String)
On Error Resume Next

    Dim rsCombo As New ADODB.Recordset, sqlCombo As String


    If distinto = True Then
        sqlCombo = "SELECT DISTINCT " & campo1 & "," & campo2 & " FROM " & tabla
    Else
        sqlCombo = "SELECT " & campo1 & "," & campo2 & " FROM " & tabla
    End If
   
    With rsCombo
        If vPathDB = "" Then
            Call .Open(sqlCombo, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(sqlCombo, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        
        ReDim MyArray(.RecordCount + 1, 2)
        
             
        combo.Clear
        
        If Not .EOF = True Then .MoveFirst
   

        Do Until .EOF = True
                combo.AddItem (Trim(.Fields(campo2).Value))
                combo.ItemData(combo.NewIndex) = (Trim(.Fields(campo1).Value))
                .MoveNext
        Loop
    
    End With
    
       
    
    sqlCombo = ""
    
    rsCombo.Close
    Set rsCombo = Nothing

If Err Then GrabarLog "CargarCombo", Err.Number & " " & Err.Description, "Global"
End Sub



Public Sub CargarComboTarjetaPorBanco(tabla As String, campo As String, combo As ComboBox, distinto As Boolean, descFiltro As String, filtro As String, Optional Indice As String, Optional vPathDB As String)
On Error Resume Next

    Dim rsCombo As New ADODB.Recordset, sqlCombo As String

    If distinto = True Then
        sqlCombo = "SELECT DISTINCT " & campo & " FROM " & tabla
    Else
        sqlCombo = "SELECT * FROM Tarjeta t inner join bancos b on b.idBancos = t.idBancos " & " where descripcion = '" & filtro & "'"
    End If
   
    With rsCombo
        If vPathDB = "" Then
            Call .Open(sqlCombo, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(sqlCombo, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        combo.Clear
        
        If Not .EOF = True Then .MoveFirst
        
        Do Until .EOF = True
            
            If Trim(Indice) = "" Then
                combo.AddItem Trim(.Fields(campo).Value)
                combo.Tag = Trim(.Fields(0).Value)
            Else
                combo.AddItem Trim(.Fields(campo).Value)
                combo.ItemData(combo.NewIndex) = Val(Trim(.Fields(Trim(Indice)).Value))
                combo.Tag = Trim(.Fields(0).Value)
            End If
            
            .MoveNext
    
        Loop
    
    End With
    
    sqlCombo = ""
    
    rsCombo.Close
    Set rsCombo = Nothing

If Err Then GrabarLog "CargarCombo", Err.Number & " " & Err.Description, "Global"
End Sub
Public Function char_i(v As Variant, _
                       l As Integer) As String
    Dim i, j As Integer

    i = Len(v)

    If IsNull(i) = True Then Exit Function

    If i > l Then v = Left(v, l)

    For j = i To l - 1
        v = v & " "
    Next

    char_i = v
End Function
Public Function CopiarArchivo(ByVal sOLDFILE As String, _
                          ByVal sNEWFILE As String, _
                          bOverWrite As Boolean) As Boolean
    Dim lTMP As Long

    If Trim(sOLDFILE) <> vbNullString And Trim(sNEWFILE) <> vbNullString Then
        lTMP = CopyFile(sOLDFILE, sNEWFILE, Not bOverWrite)

        If lTMP <> 0 Then CopiarArchivo = True Else CopiarArchivo = False
    End If

End Function

Public Sub LastKlexRow(ByRef v As KlexGrid)
       With v ' posicionarse en el ultimo registro
            If .Rows > 1 Then
            .Row = v.Rows - 1
            .TopRow = v.Row
            .RowSel = v.Row
            .Col = 0
            .ColSel = .Cols - 1
            End If
       End With
End Sub


Public Function EnLetras2(ByVal numero2 As String) As String

Dim pnumero As Double


pnumero = Val(numero2)


Dim xcen(9) 'centenas
Dim xdec(9) 'decenas
Dim xuni(9) 'unidades
Dim xexc(6) 'except
Dim ceros(9)

Dim letras
Dim i
Dim c
Dim j
Dim xnumero
Dim xnum
Dim Num
Dim digito
Dim numero_ent
Dim entero
Dim decimales
Dim temp
  
  xcen(2) = "Dosc"
  xcen(3) = "Tresc"
  xcen(4) = "Cuatrosc"
  xcen(5) = "Quin"
  xcen(6) = "Seisc"
  xcen(7) = "Setec"
  xcen(8) = "Ochoc"
  xcen(9) = "Novec"
  xdec(2) = "Veinti"
  xdec(3) = "Trei"
  xdec(4) = "Cuare"
  xdec(5) = "Cincue"
  xdec(6) = "Sese"
  xdec(7) = "Sete"
  xdec(8) = "Oche"
  xdec(9) = "Nove"
  xuni(1) = "Uno"
  xuni(2) = "Dos"
  xuni(3) = "Tres"
  xuni(4) = "Cuatro"
  xuni(5) = "Cinco"
  xuni(6) = "Seis"
  xuni(7) = "Siete"
  xuni(8) = "Ocho"
  xuni(9) = "Nueve"
  xexc(1) = "Diez"
  xexc(2) = "Once"
  xexc(3) = "Doce"
  xexc(4) = "Trece"
  xexc(5) = "Catorce"
  xexc(6) = "Quince"
  ceros(1) = "0"
  ceros(2) = "00"
  ceros(3) = "000"
  ceros(4) = "0000"
  ceros(5) = "00000"
  ceros(6) = "000000"
  ceros(7) = "0000000"
  ceros(8) = "00000000"
  
  c = 1
  i = 1
  j = 0
  
  xnumero = CStr(pnumero)
If CDbl(LTrim(RTrim(pnumero))) < 999999999.99 Then
    numero_ent = CDbl(Int(pnumero))
    If Len(numero_ent) < 9 Then
        numero_ent = ceros(9 - Len(numero_ent)) & numero_ent
    End If
    entero = CDbl(Int(numero_ent))
    decimales = (CDbl(xnumero) - entero) * 100
    
    Do While i < 8
        temp = 0
        Num = CDbl(Mid(numero_ent, i, 3))
        xnum = Mid(numero_ent, i, 3)
        digito = CDbl(Mid(xnum, 1, 1))
        
        '/* analizo el numero entero de a 3 */
        If xnum = "000" Then
            j = 0
        Else
            j = 1
            If digito > 1 Then
                letras = letras & xcen(digito) & "ientos "
            End If
            If Mid(xnum, 1, 1) = "1" And Mid(xnum, 2, 2) <> "00" Then
                letras = letras & "ciento "
            ElseIf Mid(xnum, 1, 1) = "1" Then
                letras = letras & "cien "
            End If
  
            '/* analisis de las decenas */
            digito = CDbl(Mid(xnum, 2, 1))
            If digito > 2 And Mid(xnum, 3, 1) = "0" Then
                letras = letras & xdec(digito) & "nta "
                temp = 1
            End If
            
            If digito > 2 And Mid(xnum, 3, 1) <> "0" Then
                letras = letras & xdec(digito) & "nta y "
                
            End If
            
            If digito = 2 And Mid(xnum, 3, 1) = "0" Then
                letras = letras & "veinte "
                temp = 1
            ElseIf digito = 2 And Mid(xnum, 3, 1) <> "0" Then
                letras = letras & "veinti"
                
            End If
            
            If digito = 1 And Mid(xnum, 3, 1) >= "6" Then
                letras = letras & "dieci"
            ElseIf digito = 1 And Mid(xnum, 3, 1) < "6" Then
                letras = letras & xexc(CDbl(Mid(xnum, 3, 1) + 1))
                temp = 1
            End If
        End If
   
        If temp = 0 Then
    '/* analisis del ultimo digito */
        digito = CDbl(Mid(xnum, 3, 1))
                If ((c = 1) Or (c = 2)) And xnum = "001" Then
                    letras = letras & "un"
                Else
                    If ((c = 1) Or (c = 2)) And xnum >= "020" And Mid(xnum, 3, 1) = "1" Then
                        letras = letras & "un"
                    Else
                        If digito <> 0 Then
                            letras = letras & xuni(digito)
                        End If
                    End If
                End If
        End If
  
  If j = 1 And i = 1 And xnum = "001" And c = 1 Then
    letras = letras & " millon "
  ElseIf j = 1 And i = 1 And xnum <> "001" And c = 1 Then
    letras = letras & " millones "
  ElseIf j = 1 And i = 4 And c = 2 Then
    letras = letras & " mil "
  End If
  i = i + 3
  c = c + 1
  Loop
  If letras = "" Then
  letras = "cero "
  End If
  If decimales <> 0 Then
    decimales = Round(decimales)
    
    letras = "SON: " & letras & " con " & CStr(decimales) & "/100 pesos"
  Else
    letras = "SON: " & letras & " pesos"
  End If
  
End If

EnLetras2 = UCase(letras)


End Function



Public Function EnLetras(ByVal numero As String) As String

    Dim b, paso As Integer

    Dim expresion, entero, deci, flag As String

    flag = "N"

    For paso = 1 To Len(numero)

        If Mid(numero, paso, 1) = "." Then

            flag = "S"

        Else

            If flag = "N" Then

                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero

            Else

                deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero

            End If

        End If

    Next paso

    If Len(deci) = 1 Then

        deci = deci & "0"

    End If

    flag = "N"

    If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999

        For paso = Len(entero) To 1 Step -1

            b = Len(entero) - (paso - 1)

            Select Case paso

                Case 3, 6, 9

                    Select Case Mid(entero, b, 1)

                        Case "1"

                            If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then

                                expresion = expresion & "cien "

                            Else

                                expresion = expresion & "ciento "

                            End If

                        Case "2"

                            expresion = expresion & "doscientos "

                        Case "3"

                            expresion = expresion & "trescientos "

                        Case "4"

                            expresion = expresion & "cuatrocientos "

                        Case "5"

                            expresion = expresion & "quinientos "

                        Case "6"

                            expresion = expresion & "seiscientos "

                        Case "7"

                            expresion = expresion & "setecientos "

                        Case "8"

                            expresion = expresion & "ochocientos "

                        Case "9"

                            expresion = expresion & "novecientos "

                    End Select

                Case 2, 5, 8

                    Select Case Mid(entero, b, 1)

                        Case "1"

                            If Mid(entero, b + 1, 1) = "0" Then

                                flag = "S"

                                expresion = expresion & "diez "

                            End If

                            If Mid(entero, b + 1, 1) = "1" Then

                                flag = "S"

                                expresion = expresion & "once "

                            End If

                            If Mid(entero, b + 1, 1) = "2" Then

                                flag = "S"

                                expresion = expresion & "doce "

                            End If

                            If Mid(entero, b + 1, 1) = "3" Then

                                flag = "S"

                                expresion = expresion & "trece "

                            End If

                            If Mid(entero, b + 1, 1) = "4" Then

                                flag = "S"

                                expresion = expresion & "catorce "

                            End If

                            If Mid(entero, b + 1, 1) = "5" Then

                                flag = "S"

                                expresion = expresion & "quince "

                            End If

                            If Mid(entero, b + 1, 1) > "5" Then

                                flag = "N"

                                expresion = expresion & "dieci"

                            End If

                        Case "2"

                            If Mid(entero, b + 1, 1) = "0" Then

                                expresion = expresion & "veinte "

                                flag = "S"

                            Else

                                expresion = expresion & "veinti"

                                flag = "N"

                            End If

                        Case "3"

                            If Mid(entero, b + 1, 1) = "0" Then

                                expresion = expresion & "treinta "

                                flag = "S"

                            Else

                                expresion = expresion & "treinta y "

                                flag = "N"

                            End If

                        Case "4"

                            If Mid(entero, b + 1, 1) = "0" Then

                                expresion = expresion & "cuarenta "

                                flag = "S"

                            Else

                                expresion = expresion & "cuarenta y "

                                flag = "N"

                            End If

                        Case "5"

                            If Mid(entero, b + 1, 1) = "0" Then

                                expresion = expresion & "cincuenta "

                                flag = "S"

                            Else

                                expresion = expresion & "cincuenta y "

                                flag = "N"

                            End If

                        Case "6"

                            If Mid(entero, b + 1, 1) = "0" Then

                                expresion = expresion & "sesenta "

                                flag = "S"

                            Else

                                expresion = expresion & "sesenta y "

                                flag = "N"

                            End If

                        Case "7"

                            If Mid(entero, b + 1, 1) = "0" Then

                                expresion = expresion & "setenta "

                                flag = "S"

                            Else

                                expresion = expresion & "setenta y "

                                flag = "N"

                            End If

                        Case "8"

                            If Mid(entero, b + 1, 1) = "0" Then

                                expresion = expresion & "ochenta "

                                flag = "S"

                            Else

                                expresion = expresion & "ochenta y "

                                flag = "N"

                            End If

                        Case "9"

                            If Mid(entero, b + 1, 1) = "0" Then

                                expresion = expresion & "noventa "

                                flag = "S"

                            Else

                                expresion = expresion & "noventa y "

                                flag = "N"

                            End If

                    End Select

                Case 1, 4, 7

                    Select Case Mid(entero, b, 1)

                        Case "1"

                            If flag = "N" Then

                                If paso = 1 Then

                                    expresion = expresion & "uno "

                                Else

                                    expresion = expresion & "un "

                                End If

                            End If

                        Case "2"

                            If flag = "N" Then

                                expresion = expresion & "dos "

                            End If

                        Case "3"

                            If flag = "N" Then

                                expresion = expresion & "tres "

                            End If

                        Case "4"

                            If flag = "N" Then

                                expresion = expresion & "cuatro "

                            End If

                        Case "5"

                            If flag = "N" Then

                                expresion = expresion & "cinco "

                            End If

                        Case "6"

                            If flag = "N" Then

                                expresion = expresion & "seis "

                            End If

                        Case "7"

                            If flag = "N" Then

                                expresion = expresion & "siete "

                            End If

                        Case "8"

                            If flag = "N" Then

                                expresion = expresion & "ocho "

                            End If

                        Case "9"

                            If flag = "N" Then

                                expresion = expresion & "nueve "

                            End If

                    End Select

            End Select

            If paso = 4 Then

                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And Len(entero) <= 6) Then
               ' If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And Len(entero) <= 5) Then

                If Val(entero) >= 1000 Then expresion = expresion & "mil "

                End If

            End If

            If paso = 7 Then

                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then

                    expresion = expresion & "millón "

                Else

                    expresion = expresion & "millones "

                End If

            End If

        Next paso

        If deci <> "" Then

            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo

                EnLetras = "menos " & expresion & "con " & deci ' & "/100"

            Else

                EnLetras = expresion & "con " & deci ' & "/100"

            End If

        Else

            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo

                EnLetras = "menos " & expresion

            Else

                EnLetras = expresion

            End If

        End If

    Else 'si el numero a convertir esta fuera del rango superior e inferior

        EnLetras = ""

    End If

End Function
Public Sub Espero(Segundos As Single)
    Dim ComienzoSeg As Single
    Dim FinSeg As Single
    ComienzoSeg = Timer
    FinSeg = ComienzoSeg + Segundos

    Do While FinSeg > Timer
        DoEvents

        If ComienzoSeg > Timer Then
            FinSeg = FinSeg - 24 * 60 * 60
        End If

    Loop

End Sub
Public Sub GrabarLog(vProceso As String, _
                     vestado As String, _
                     vName As String)

    Dim rsLog As New ADODB.Recordset, sqlLog As String
    
    sqlLog = "SELECT * FROM log WHERE 2 = 1"
   Exit Sub
    If Not Left(vestado, 1) = "-" Then Exit Sub
    
    With rsLog
        .CursorLocation = adUseClient
        Call .Open(sqlLog, PathDBConfig, adOpenDynamic, adLockOptimistic)
        
        If Not .State = 0 Then
        
            .AddNew
            
            .Fields("Hora") = Format(Time, "hh:mm:ss")
            .Fields("fecha") = Format(Date)
            .Fields("proceso") = vProceso
            .Fields("Formulario") = Trim(Left(vName, 49))
            .Fields("Comentario") = Trim(Left(vestado, 99))
            
            .Update
        End If
        
    End With

    sqlLog = ""
    
    If rsLog.State = 1 Then
        rsLog.Close
        Set rsLog = Nothing
    End If
    
End Sub
Function inulo(v) As String

    If Not Format(v, "########.##") = "" Then
        inulo = v
    Else
        inulo = 0
    End If

End Function
Public Function num_i(v As Variant) As String
    Dim i, j As Integer

    v = Format(Format(v, "#######0.00"), "@@@@@@@@@")
    i = Len(v)
    num_i = v
End Function

Public Function num_i2(v As Variant) As String
    Dim i, j As Integer

    v = Format(v, "@@@@")
    i = Len(v)

    num_i2 = v
End Function
Public Sub OrdenarDataGrid(ByVal ColIndex As Integer, _
                           rs As ADODB.Recordset, _
                           DataGrid As DataGrid)

    Dim strColName As String
    Static bSortAsc As Boolean
    Static strPrevCol As String
    
    strColName = DataGrid.Columns(ColIndex).DataField
 
    If strColName = strPrevCol Then

        If bSortAsc Then
            rs.Sort = strColName & " DESC"
            bSortAsc = False
        Else
            rs.Sort = strColName
            bSortAsc = True
        End If

    Else
        rs.Sort = strColName & " ASC"
        bSortAsc = True
    End If
 
    strPrevCol = strColName
 
End Sub



Function pathDBMigrar(vDBaMigrar As String) As String
On Error Resume Next
    
    If InStr(1, vDBaMigrar, ".DBF") = 0 Then
        
        Select Case vDBaMigrar
    
            Case "OscarValentini", "MTavani", "ARV"
                pathDBMigrar = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Datos\" & vDBaMigrar & ".mdb;Persist Security Info=False"
    
            Case "WGestionServicios"
                pathDBMigrar = "driver={MySQL ODBC 3.51 Driver};server=192.168.0.201;uid=root;pwd=root.2009;database=" & vDBaMigrar & ";OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384

        End Select
    
    Else
        pathDBMigrar = PathDBSisagro(vDBaMigrar)
    End If
    
If Err Then GrabarLog "pathDBMigrar", Err.Number & " " & Err.Description, "Global"
End Function
Function pathDBMySQL() As String
On Error Resume Next
    
    'ORIGINAL pathDBMySQL = "driver={MySQL ODBC 3.51 Driver};server=" & vConfigGral.vDireccionDB & ";uid=" & vConfigGral.vUserDB & ";pwd=" & vConfigGral.vPassDB & ";database=" & vConfigGral.vEmpresa & ";OPTION=2 + 8 + 32 + 2048 + 16384"
    pathDBMySQL = "driver={MySQL ODBC 3.51 Driver};server=" & vConfigGral.vDireccionDB & ";port=3306;uid=" & vConfigGral.vUserDB & ";pwd=" & vConfigGral.vPassDB & ";database=" & vConfigGral.vempresa & ";OPTION=8"
    
If Err Then GrabarLog "pathDBMySQL", Err.Number & " " & Err.Description, "Global"
End Function

Function pathDBMySQLComuna() As String
On Error Resume Next
    
    'ORIGINAL pathDBMySQL = "driver={MySQL ODBC 3.51 Driver};server=" & vConfigGral.vDireccionDB & ";uid=" & vConfigGral.vUserDB & ";pwd=" & vConfigGral.vPassDB & ";database=" & vConfigGral.vEmpresa & ";OPTION=2 + 8 + 32 + 2048 + 16384"
    pathDBMySQLComuna = vConfigGral.vComunaCnn
    
If Err < 0 Then GrabarLog "pathDBMySQL", Err.Number & " " & Err.Description, "Global"
End Function


Function PathDBConfig() As String
On Error Resume Next

    PathDBConfig = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Datos\Configuracion.mdb;Persist Security Info=False"

If Err Then GrabarLog "PathDBConfig", Err.Number & " " & Err.Description, "Global"
End Function

Function PathDBListados()
On Error Resume Next

    PathDBListados = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Datos\Listados.mdb;Persist Security Info=False"

If Err Then GrabarLog "PathDBConfig", Err.Number & " " & Err.Description, "Global"
End Function
Function PathDBSisagro(vtabla As String) As String
On Error Resume Next
    
    Dim vPathSisAgro As String
    
    vPathSisAgro = "C:\sisagro\exe\AMigrar\"

    PathDBSisagro = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vPathSisAgro & "\;Extended Properties=dBASE IV;" 'User ID=;Password=;"
    'PathDBSisagro = "Driver={Microsoft Visual FoxPro Driver (*.dbf)};DriverID=277;Dbq=" & vPathSisAgro & "\" & vTabla & ".DBF"
    'PathDBSisagro = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" & vPathSisAgro & "\" & vTabla & ".DBF" & ";Exclusive=no;"
    'PathDBSisagro = "Driver={Microsoft Visual FoxPro Driver};SourceType=DBF;SourceDB=" & vPathSisAgro & ";Exclusive=NO;"

If Err Then GrabarLog "PathDBSisagro", Err.Number & " " & Err.Description, "Global"
End Function

Function pathDBDEmySQL(vShape As Boolean)
    
    If vShape = False Then
    
        pathDBDEmySQL = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=" & vConfigGral.vempresa & ";User ID=" & vConfigGral.vUserDB & ";Password=" & vConfigGral.vPassDB & ";Initial Catalog=" & vConfigGral.vempresa & ""
    
    Else
        
        pathDBDEmySQL = "Provider=MSDataShape.1;Persist Security Info=False;Data Source=" & vConfigGral.vempresa & ";User ID=" & vConfigGral.vUserDB & ";Password=" & vConfigGral.vPassDB & ";Initial Catalog=" & vConfigGral.vempresa & ";Data Provider=MSDASQL.1"
    
    End If
    
End Function
Private Function PathDBSQL2K() As String
On Error Resume Next

    PathDBSQL2K = "Provider=sqloledb;Data Source=rg58sma;Initial Catalog=Produccion;User Id=sa;Password=root.2009;"

If Err Then GrabarLog "PathDBSQL2K", Err.Number & " " & Err.Description, "Global"
End Function
Function pathDBExcel(vArchivo) As String
On Error Resume Next

    pathDBExcel = "Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=" & vArchivo & ";DefaultDir=" & App.Path & "\Listas de Precio\;"

If Err Then GrabarLog "pathDBExcel", Err.Number & " " & Err.Description, "Global"
End Function
Function pathDBTemp() As String
    pathDBTemp = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & vConfigGral.vDireccionDB & "Temp.mdb;Persist Security Info=False"
End Function
Function pathDBLocalidad() As String
    pathDBLocalidad = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Datos\PeqProd.mdb;Persist Security Info=False"
End Function
Function pathDBDE() As String
    
    With vConfigGral
        pathDBDE = "Data Source =" & .vDireccionDB & .vempresa & ".mdb"
    End With

End Function
Function PonerPunto(v) As String
    On Error Resume Next

    Dim auxi, head, tail As String
    Dim i As Integer

    If Trim(Left(v, 3)) = "u$s" Then v = Right(v, Len(v) - 3)
    If Trim(Left(v, 1)) = "$" Then v = Right(v, Len(v) - 1)

    i = 1

    Do Until auxi = "," Or Len(v) <= i
        auxi = Mid(v, i, 1)
        i = i + 1
    Loop

    If auxi = "," Then
        i = i - 1
        head = Left(v, Len(v) - (Len(v) - i) - 1)
        tail = Right(v, Len(v) - i)
    
        v = head & "." & tail
    End If

    If Val(v) = 0 Then v = "0"

    PonerPunto = Trim(v)

    If Err Then Exit Function
End Function
Function EsNulo(vValor As Variant) As String
On Error Resume Next
    

    
    If IsNull(vValor) = True Then
        EsNulo = ""
    Else
     
            EsNulo = Replace(vValor, "'", "*")
   
    End If

If Err Then GrabarLog "EsNulo", Err.Number & " " & Err.Description, "Global"
End Function
Function EsNuloGuion(vValor As Variant) As String
On Error Resume Next
    
    If IsNull(vValor) = True Or Trim(vValor) = "" Then
        EsNuloGuion = "-"
    Else
        EsNuloGuion = vValor
    End If

If Err Then GrabarLog "EsNuloGuion", Err.Number & " " & Err.Description, "Global"
End Function
Function strFecha(vfecha As Date) As String
    strFecha = Format(vfecha, "mm/dd/yyyy")
End Function
Function strfecha2(vfecha As Date) As String
    strfecha2 = Format(vfecha, "dd/mm/yyyy")
End Function
Function strfechaMySQL(vfecha As Date) As String
    strfechaMySQL = Format(vfecha, "yyyy-mm-dd")
End Function
Public Function Wait(ByVal TimeToWait As Long) 'Time In seconds
    Dim EndTime As Long
    
    EndTime = GetTickCount + TimeToWait '1000 Cause u give seconds and GetTickCount uses Milliseconds

    Do Until GetTickCount > EndTime

        DoEvents
    
    Loop

End Function
Public Function UltimoRemito(vtabla As String, Optional vPathDB) As Long
On Error Resume Next
    
    Dim rsURemito As New ADODB.Recordset
    Dim sqlURemito As String
    
    sqlURemito = "SELECT * FROM " & Trim(vtabla) & " ORDER BY remito ASC"

    With rsURemito
        If IsMissing(vPathDB) = True Then
            Call .Open(sqlURemito, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(sqlURemito, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        If Not .EOF = True Then
            .MoveLast
            UltimoRemito = Val(.Fields("Remito").Value) + 1
        Else
            UltimoRemito = 1
        End If
    End With
    
    sqlURemito = ""
    
    rsURemito.Close
    Set rsURemito = Nothing

If Err Then GrabarLog "UltimoRemito", Err.Number & " " & Err.Description, "Global"
End Function
Public Sub CerrarForms(vtipo As Byte)
On Error Resume Next
    
    Unload frmArticulos
    Unload frmRemito
    Unload frmCtaCteC
    Unload frmBuscarCliente
    'Unload frmBuscarEmpleado
    'Unload frmBuscarContacto
    Unload frmBuscarProveedor
    Unload frmBuscarArticulo
    Unload frmLogin
    
    If vtipo = 1 Then End

    
If Err Then GrabarLog "CerrarForms", Err.Number & " " & Err.Description, "Global"
End Sub




Public Function UltimoNroOrdenPago(vn As Long) As Long
On Error Resume Next


Dim vvid As Long
Dim vsql As String
Dim i As Integer
vvid = 0


Dim v As Long


vvid = traerDatos2("select max(t.numero) as c from t_nro_orden_pago t", "c", pathDBMySQL)

If vn > 0 Then

    
    If vvid <= vn Then
    
        MsgBox "El nro de orden sugerido por el usuario ya fue utilizado "
        UltimoNroOrdenPago = 0
    
    Else
        v = vn
    End If
    
Else
    
    v = vvid + 1
End If

vsql = "insert into t_nro_orden_pago (numero) values (" + Str(v) + ")"
Call EjecutarScript(vsql, pathDBMySQL)
vvid = traerDatos2("select max(t.`id`) as c from t_nro_orden_pago t", "c", pathDBMySQL)
        

UltimoNroOrdenPago = v


If Err < 0 Then
    MsgBox "Ocurrió un error grave. El sistema necesita cerrarce." + Chr(13) + "Vualva a realizar la operación" _
    + Chr(13) + "Si el error persiste consulte al servicio técnico" + Chr(13) + _
    "Error: " + Err.Description + "  Nro:Interno: " + Str(vvid)
    End
End If
End Function



Public Function UltimoNroInterno2() As Long
On Error Resume Next


Dim vvid, vvid_inicial As Long
Dim vsql As String
Dim i As Integer

Dim error_nrointerno As Boolean

vvid = 0

error_nrointerno = False
vvid_inicial = traerDatos2("select max(t.`numero`) as c from t_nrointerno t", "c", pathDBMySQL)


vsql = "insert into t_nrointerno (auxiliar) values (1)"
Call EjecutarScript(vsql, pathDBMySQL)

vvid = traerDatos2("select max(t.`numero`) as c from t_nrointerno t", "c", pathDBMySQL)
        
                                   
                            
'                            Do While vvid = 1 Or i > 10
'                                    i = i + 1
'                                    Call EjecutarScript(vsql, pathDBMySQL)
'                                    vvid = traerDatos2("select max(t.`numero`) as c from t_nrointerno t", "c", pathDBMySQL)
'                            Loop



If vvid_inicial < Val(vvid) Then
    UltimoNroInterno2 = vvid
Else

    Dim iii As Integer
    iii = 0
    Do
        
        vvid_inicial = traerDatos2("select max(t.`numero`) as c from t_nrointerno t", "c", pathDBMySQL)
        vsql = "insert into t_nrointerno (auxiliar) values (1)"
        Call EjecutarScript(vsql, pathDBMySQL)
        vvid = traerDatos2("select max(t.`numero`) as c from t_nrointerno t", "c", pathDBMySQL)
        
        iii = iii + 1
        
    
    Loop Until vvid_inicial < Val(vvid) Or iii > 10
    
    UltimoNroInterno2 = vvid

End If




If Err < 0 Or vvid = 1 Or error_nrointerno Then
    MsgBox "Ocurrió un error grave. El sistema necesita cerrarce." + Chr(13) + "Vualva a realizar la operación" _
    + Chr(13) + "Si el error persiste consulte al servicio técnico" + Chr(13) + _
    "Error: " + Err.Description + "  Nro:Interno: " + Str(vvid)
    End
End If
End Function



Public Function UltimoNroInterno2Vieja() As Long
On Error Resume Next
Dim vmax, vTemp As Long
Dim vsql As String


vmax = 0


vTemp = traerDatos2("select max(t.`NroInterno`) as c from factura t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp

vTemp = traerDatos2("select max(t.`NroInterno`) as c from pfactura t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp


vTemp = traerDatos2("select max(t.`NroInterno`) as c from cuentascorrientes t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp


vTemp = traerDatos2("select max(t.`NroInterno`) as c from pcuentascorrientes t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp


vTemp = traerDatos2("select max(t.`NroInterno`) as c from bancosmovimientos t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp



vTemp = traerDatos2("select max(t.`NroInterno`) as c from ivafacturacompra t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp



vTemp = traerDatos2("select max(t.`NroInterno`) as c from ivafacturaventa t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp



vTemp = traerDatos2("select max(t.`NroInterno`) as c from asientos t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp

vTemp = traerDatos2("select max(t.`auxiliar`) as c from t_nrointerno t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp

vmax = vmax + 1

vsql = "select auxiliar from t_nrointerno t where t.auxiliar=" + Str(vmax)



If Not ((vmax - traerDatos2(vsql, "auxiliar", pathDBMySQL)) = 0) Then
     vsql = "insert into t_nrointerno (auxiliar) values (" + Str(vmax) + ")"
    Call EjecutarScript(vsql, pathDBMySQL)
Else


End If

UltimoNroInterno2Vieja = vmax

If Err Then
UltimoNroInterno2Vieja = 1
End If
End Function



Public Function UltimoNroInterno2ParaCambiodeNro() As Long
On Error Resume Next
Dim vmax, vTemp As Long
Dim vsql As String


vmax = 0


vTemp = traerDatos2("select  t.`NroInterno` as c , t.`idFactura`  from factura t order by t.`idFactura` desc", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp

vTemp = traerDatos2("select  t.`NroInterno` as c , t.`idpFactura`  from pfactura t order by t.`idpFactura` desc", "c", pathDBMySQL)

If vmax < vTemp Then vmax = vTemp

vTemp = traerDatos2("select  t.`NroInterno` as c , t.`id`  from cuentascorrientes t order by t.`id` desc", "c", pathDBMySQL)

If vmax < vTemp Then vmax = vTemp

vTemp = traerDatos2("select  t.`NroInterno` as c , t.`IdPcuentascorrientes`  from pcuentascorrientes t order by t.`IdPcuentascorrientes` desc", "c", pathDBMySQL)

If vmax < vTemp Then vmax = vTemp

vTemp = traerDatos2("select  t.`NroInterno` as c , t.`idBancosMovimientos`  from bancosmovimientos t order by t.`idBancosMovimientos` desc", "c", pathDBMySQL)

If vmax < vTemp Then vmax = vTemp


vTemp = traerDatos2("select  t.`NroInterno` as c , t.`idIvaFacturaCompra`  from ivafacturacompra t order by t.`idIvaFacturaCompra` desc", "c", pathDBMySQL)

If vmax < vTemp Then vmax = vTemp

vTemp = traerDatos2("select  t.`NroInterno` as c , t.`idIvaFacturaVenta`  from ivafacturaventa t order by t.`idIvaFacturaVenta` desc", "c", pathDBMySQL)

If vmax < vTemp Then vmax = vTemp


vTemp = traerDatos2("select  t.`NroInterno` as c , t.`idAsientos`  from asientos t order by t.`idAsientos` desc", "c", pathDBMySQL)

If vmax < vTemp Then vmax = vTemp


vTemp = traerDatos2("select  t.`auxiliar` as c , t.`numero`  from t_nrointerno t order by t.`numero` desc", "c", pathDBMySQL)
'vTemp = traerDatos2("select max(t.`auxiliar`) as c from t_nrointerno t", "c", pathDBMySQL)
If vmax < vTemp Then vmax = vTemp

vmax = vmax + 1

vsql = "select auxiliar from t_nrointerno t where t.auxiliar=" + Str(vmax)



If Not ((vmax - traerDatos2(vsql, "auxiliar", pathDBMySQL)) = 0) Then
     vsql = "insert into t_nrointerno (auxiliar) values (" + Str(vmax) + ")"
    Call EjecutarScript(vsql, pathDBMySQL)
Else


End If

UltimoNroInterno2ParaCambiodeNro = vmax

If Err Then
UltimoNroInterno2ParaCambiodeNro = 1
End If
End Function




Public Function VaciarTemporales() As Boolean
On Error Resume Next

    BorrarBase "cuentascorrientes_temp", pathDBMySQL
    BorrarBase "Temp_Documentos", pathDBMySQL
    BorrarBase "Temp_FacturaClientes", pathDBMySQL
    BorrarBase "Temp_FacturaDetalle", pathDBMySQL
    BorrarBase "Liqui_Temp", pathDBMySQL
    BorrarBase "Temp_LibretaDetalle", pathDBMySQL
    BorrarBase "Saldos", pathDBMySQL
    BorrarBase "Temclientes", pathDBMySQL
    BorrarBase "Vista", pathDBMySQL
    BorrarBase "Temp", pathDBMySQL
    BorrarBase "Temp2", pathDBMySQL
    
If Err Then
    GrabarLog "VaciarTemporales", Err.Number & " " & Err.Description, "Global"
    VaciarTemporales = False
Else
    VaciarTemporales = True
End If
End Function
Public Sub Compactar(vPath As String)
On Error Resume Next
        
   '  Set oAccess = New Access.Application ' sacado
        
   '  compactarBD vPath ' sacado
        
    ' Set oAccess = Nothing

If Err Then GrabarLog "Compactar", Err.Number & " " & Err.Description, "Global"
End Sub

'Public Function compactarBD(pathbase As String) As Boolean ' sacado
'Dim sDatabase, sBackup As String
'
'    On Local Error GoTo Handler
'
'    sDatabase = Trim$(pathbase)
'    oAccess.OpenCurrentDatabase sDatabase
'
'   ' Screen.MousePointer = vbHourglass
'
'    'Setup the Paths
'    sBackup = Replace(sDatabase, ".mdb", ".wsf")
'
'    'Close the active database so we can compact it
'    oAccess.CloseCurrentDatabase
'
'    'Compact Database to New Database File
'    oAccess.DBEngine.CompactDatabase sDatabase, sBackup
'
'    'Kill the old database file if it exists
'    BorrarArchivo sDatabase
'
'    'Copy the File back to its original name and kill the backup
'    CopiarArchivo sBackup, sDatabase, True
'    BorrarArchivo sBackup
'
'    'Reabre la base
'    oAccess.OpenCurrentDatabase sDatabase
'
'    Screen.MousePointer = vbNormal
'
'    Exit Function
'
'Handler:
'    Screen.MousePointer = vbNormal
'
'    Select Case Err.Number
'        Case 3356
'            'recopy
'            MsgBox Err.Number & " - " & Err.Description, vbExclamation, "Mensaje ..."
'        Case Is <> 0
'            MsgBox Err.Number & " - " & Err.Description, vbExclamation, "Mensaje ..."
'    End Select
'
'End Function
Public Function Comprimir() As String
On Error GoTo vbErrorHandler
    
    Dim oZip As CGZipFiles
    Set oZip = New CGZipFiles
    
    'Nombre del archivo a generarse
    With oZip
        Comprimir = App.Path & "\backup\" & Trim$(Right$(Date, 4) & Left$(Right$(Date, 7), 2) & Left$(Date, 2) & Left$(Time, 2) & Mid(Time, 4, 2)) & ".zip"
        .ZipFileName = App.Path & "\backup\" & Trim$(Right$(Date, 4) & Left$(Right$(Date, 7), 2) & Left$(Date, 2) & Left$(Time, 2) & Mid(Time, 4, 2)) & ".zip"
        
        'Actualiza el archivo ZIP (False)
        .UpdatingZip = False
    
        'Archivos a agregar en el zip
        .AddFile (App.Path & "\backup\" & vConfigGral.vempresa & "\*.*")
        If .MakeZipFile <> 0 Then
            MsgBox .GetLastMessage ' any errors
        End If
    
    End With
        
    'Queda en cero el Objeto
    Set oZip = Nothing

    Exit Function

vbErrorHandler:
    If Err.Number Then
        MsgBox Err.Number & " " & "Form1::cmdZip_Click" & " " & Err.Description
        GrabarLog "Comprimir", Err.Number & " " & Err.Description, "Global"
    End If
End Function
Public Sub CopiarBorrar(vtipo As Byte, Optional vArchivoZip As String)
On Error Resume Next

    Dim FSO As New FileSystemObject, vPathMySQL As String
    
    vPathMySQL = "C:\Archivos de programa\MySQL\MySQL Server 6.0\Data\"
    
    With vConfigGral
    
        Select Case vtipo
    
            Case 0
                'Call FSO.DeleteFolder(App.Path & "\Backup\" & .vEmpresa)
                
                'If Not FSO.FolderExists(App.Path & "\Backup\" & .vEmpresa) = True Then
                '    Call FSO.CreateFolder(App.Path & "\Backup\" & .vEmpresa)
                'End If
                'Err.Clear
                If FSO.FolderExists(vPathMySQL & .vempresa) = True Then
                    If FSO.FolderExists(App.Path & "\Backup\") = True Then
                        Call FSO.CopyFolder(vPathMySQL & .vempresa, App.Path & "\Backup\", False)
                    End If
                End If
                
                
            
            Case 1
                Call FSO.DeleteFolder(App.Path & "\Backup\" & .vempresa)

            Case 2
                Call FSO.CopyFile(vArchivoZip, LeerConfig(20) & ":\", False)
                
        End Select
    
    End With
    
If Err Then GrabarLog "CopiarBorrar", Err.Number & " " & Err.Description, "Global"
End Sub
Public Sub DatosGenerales(vOpen As Boolean, vcampo As Integer, vOpcion)
    On Error Resume Next
    
    Dim rsConfigura As New ADODB.Recordset, sqlConfigura As String

    sqlConfigura = "SELECT * FROM Configura"
    
    With rsConfigura
    
        If Not vOpen = True Then
            Call .Open(sqlConfigura, ConnDDBB, adOpenStatic, adLockReadOnly)
            
            'gnombre = Trim(.Fields("nombre").Value & " ")
            'gtelefono = Trim(.Fields("telefono").Value & " ")
            'gdireccion = Trim(.Fields("direccion").Value & " ")
            'giva = .Fields("Iva").Value
            'gdolar = .Fields("Dolar").Value
            
        
        Else
            Call .Open(sqlConfigura, ConnDDBB, adOpenDynamic, adLockPessimistic)
            
            .Fields(vcampo).Value = vOpcion
            .Update
        
        End If

    End With
    
    sqlConfigura = ""
    
    rsConfigura.Close
    Set rsConfigura = Nothing
    
    If Err Then GrabarLog "DatosEmpresa", Err.Number & " " & Err.Description, "Global"
End Sub
Public Function DepurarFDetalle(vsql As String) As Boolean
On Error Resume Next
    
    BorrarBase "FDetalle WHERE " & vsql, pathDBMySQL

If Err Then
    GrabarLog "DepurarFDetalle", Err.Number & " " & Err.Description, "Global"
Else
    DepurarFDetalle = True
End If
End Function
Public Function BorrarArchivo(ByVal sFile As String) As Boolean
    Dim lTMP As Long

    If Trim(sFile) <> vbNullString Then
        lTMP = DeleteFile(sFile)

        If lTMP = 0 Then BorrarArchivo = True Else BorrarArchivo = False
    Else
        BorrarArchivo = False
    End If

End Function

Public Sub conexionOk()
On Error Resume Next
Dim dato, vsql  As String

vsql = "select count(*) as c from t_nrointerno"

dato = ""
dato = traerDatos2(vsql, "c", pathDBMySQL)


If dato = "" Then

   ' MsgBox "Atención !" + Chr(13) + _
    " 1) El servidor del sistema de Servicio debe estar prendido" + Chr(13) + _
    " 2) El servidor del sistema de Caja debe estar prendido" + Chr(13) + _
    " 3) Debe estar conectado a la red WIFI llamada Wifi-Coop"

End If

If Err Then Exit Sub
End Sub

Public Sub LoadConfigRemito()
Dim i As Integer
On Error Resume Next

    Dim rsLCRemito As New ADODB.Recordset, sqlLCRemito As String
    
    sqlLCRemito = "SELECT * FROM Configura_remito"
    
    With rsLCRemito
        Call .Open(sqlLCRemito, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        ReDim ConfigRemito(6)

        For i = 0 To 6

            ConfigRemito(i) = .Fields(i).Value
    
        Next
    
    End With
    
    sqlLCRemito = ""
    
    rsLCRemito.Close
    Set rsLCRemito = Nothing

If Err Then GrabarLog "LoadConfigRemito", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Public Sub SaveConfigRemito()
Dim i As Integer
On Error Resume Next

    Dim rsSCRemito As New ADODB.Recordset, sqlSCRemito As String
    
    sqlSCRemito = "SELECT * FROM Configura_remito"
    
    With rsSCRemito
        Call .Open(sqlSCRemito, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .EOF = True Then .AddNew
        
        For i = 0 To 6

            .Fields(i).Value = ConfigRemito(i)
    
        Next
        
        .Update
    End With
    
    sqlSCRemito = ""
    
    rsSCRemito.Close
    Set rsSCRemito = Nothing

If Err Then GrabarLog "SaveConfigRemito", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Public Sub FTP(ByRef vArchivoZip As String)
On Error Resume Next
    
    ConexionFTP (vArchivoZip)

If Err Then GrabarLog "FTP", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Private Sub ConexionFTP(ByRef vArchivoZip As String)
On Error Resume Next

    'With frmPrincipal.FTPEngine

        'Establezco la conexion
        '.RemoteHost = LeerConfig(4)
        '.UserName = LeerConfig(5)
        '.Password = LeerConfig(6)
        '.RemotePort = LeerConfig(7)
        '.Connect
        
        
        'Intento copiar el archivo si pasó la conexión
        'Call .PutFile(LeerConfig(8), vArchivoZip, vArchivoZip, True)
    'End With

If Err Then GrabarLog "ConexionFTP", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Public Function LeerConfig(vCase As String, Optional vValor) As Variant

    Dim i As Integer
    Dim Doc As MSXML2.DOMDocument, Nod As MSXML2.IXMLDOMNode
    Set Doc = New MSXML2.DOMDocument
   
    If Doc.Load(App.Path & "\configuracion.xml") Then
      'Screen.MousePointer = vbHourglass
      
        For Each Nod In Doc.getElementsByTagName("Configuracion") 'Recorrer todos los nodos <Registro> del documento XML
            
            If IsMissing(vValor) = True Then
                For i = Val(vCase) To Val(vCase)
                    Select Case i
                        Case 0
                            LeerConfig = Nod.selectSingleNode("Sistema").Text
                        Case 1
                            LeerConfig = Nod.selectSingleNode("Titulo").Text
                        Case 2
                            LeerConfig = Nod.selectSingleNode("UServidor").Text
                        Case 3
                            LeerConfig = Nod.selectSingleNode("UUsuario").Text
                        Case 4
                            LeerConfig = Nod.selectSingleNode("UEmpresa").Text
                        Case 5
                            LeerConfig = EsNulo(Nod.selectSingleNode("Password").Text)
                        Case 6
                            LeerConfig = EsNulo(Nod.selectSingleNode("Login").Text)
                        Case 7
                            LeerConfig = Nod.selectSingleNode("Demo").Text
                        Case 8
                            LeerConfig = Nod.selectSingleNode("TipoVersion").Text
                        Case 9
                            LeerConfig = Nod.selectSingleNode("IncluyeContabilidad").Text
                        Case 10
                            LeerConfig = Nod.selectSingleNode("IncluyeStac").Text
                        Case 11
                            LeerConfig = Nod.selectSingleNode("IncluyeResto").Text
                        Case 12
                            LeerConfig = Nod.selectSingleNode("IncluyeTicket").Text
                        Case 13
                            LeerConfig = EsNulo(Nod.selectSingleNode("IncluyeCobros").Text)
                        Case 14
                            LeerConfig = EsNulo(Nod.selectSingleNode("Impresora").Text)
                        Case 15
                            LeerConfig = EsNulo(Nod.selectSingleNode("ImprimirReciboCliente").Text)
                        Case 16
                            LeerConfig = EsNulo(Nod.selectSingleNode("ImprimirReciboProveedor").Text)
                        Case 17
                            LeerConfig = EsNulo(Nod.selectSingleNode("TipoDeRemitoInicialVenta").Text)
                        Case 18
                            LeerConfig = EsNulo(Nod.selectSingleNode("TipoDeRemitoInicialCompra").Text)
                        Case 19
                            LeerConfig = EsNulo(Nod.selectSingleNode("TamanoIconosMenuPrincipal").Text)
                        Case 20
                            LeerConfig = EsNulo(Nod.selectSingleNode("UnidadExternaBackup").Text)
                        Case 21
                            LeerConfig = EsNulo(Nod.selectSingleNode("TipoListadoIvaVenta").Text)
                        Case 22
                            LeerConfig = EsNulo(Nod.selectSingleNode("TipoListadoIvaCompra").Text)
                        Case 23
                            LeerConfig = EsNulo(Nod.selectSingleNode("CargarDocumentosEnCobros").Text)
                        Case 24
                            LeerConfig = EsNulo(Nod.selectSingleNode("CargarDocumentosEnPagos").Text)
                        Case 25
                            LeerConfig = EsNulo(Nod.selectSingleNode("MantenerClienteEnVentas").Text)
                        Case 26
                            LeerConfig = EsNulo(Nod.selectSingleNode("MantenerProveedorEnCompras").Text)
                        Case 27
                            LeerConfig = EsNulo(Nod.selectSingleNode("Remito").Text)
                        Case 28
                            LeerConfig = EsNulo(Nod.selectSingleNode("AsientosAutomaticos").Text)
                        Case 30
                            LeerConfig = EsNulo(Nod.selectSingleNode("TipoCliente").Text)
                        Case 31
                            LeerConfig = EsNulo(Nod.selectSingleNode("ComunaCnn").Text)
                        
                    
                    End Select

                Next i
            
            Else
                Screen.MousePointer = vbDefault
                Nod.selectSingleNode(vCase).Text = vValor
                Doc.Save (App.Path & "\configuracion.xml")
            
            End If
        
        Next Nod
        Screen.MousePointer = vbDefault
    Else
        MsgBox "No se puede abrir el archivo " & App.Path & "\configuracion.xml", vbCritical, "Error"
    End If
   
End Function

Public Function LeerXml(vtag As String) As Variant
On Error Resume Next

    Dim i As Integer
    Dim Doc As MSXML2.DOMDocument, Nod As MSXML2.IXMLDOMNode
    Set Doc = New MSXML2.DOMDocument
   
    If Doc.Load(App.Path & "\configuracion.xml") Then
      'Screen.MousePointer = vbHourglass
      
    For Each Nod In Doc.getElementsByTagName("Configuracion")
        LeerXml = Nod.selectSingleNode(vtag).Text
      Next
 
        Screen.MousePointer = vbDefault
    Else
        MsgBox "No se puede abrir el archivo " & App.Path & "\configuracion.xml", vbCritical, "Error"
    End If
If Err Then
    LeerXml = ""
    Exit Function
End If
End Function


Public Function LeerXmlRecibido(vnomr As String) As Variant
On Error Resume Next

    Dim i As Integer
    Dim Doc As MSXML2.DOMDocument, Nod As MSXML2.IXMLDOMNode
    Set Doc = New MSXML2.DOMDocument
   
    If Doc.Load(App.Path & "\Log\" + vnomr) Then
      
        Dim v As String
        
        v = "FERecuperaLastCbteResponse xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' xmlns:xsd='http://www.w3.org/2001/XMLSchema'"
        LeerXmlRecibido = Doc.documentElement.childNodes(2).Text
    
    
    End If
If Err Then
    LeerXmlRecibido = ""
    Exit Function
End If
End Function



Public Function CotizacionDolar(vCase As String, Optional vValor) As Variant
    
    Dim i As Integer
    Dim Doc As MSXML2.DOMDocument, Nod As MSXML2.IXMLDOMNode
    Set Doc = New MSXML2.DOMDocument
   
    If Doc.Load(App.Path & "\dolar.xml") Then
     ' Screen.MousePointer = vbHourglass
      
        For Each Nod In Doc.getElementsByTagName("MIDOLAR")
            
            If IsMissing(vValor) = True Then
                For i = Val(vCase) To Val(vCase)
                    Select Case i
                        Case 0
                            CotizacionDolar = Nod.selectSingleNode("VALORCOMPRA").Text
                        Case 1
                            CotizacionDolar = Nod.selectSingleNode("VALORVENTA").Text
                        Case 2
                            CotizacionDolar = Nod.selectSingleNode("HORA").Text
                        Case 3
                            CotizacionDolar = Nod.selectSingleNode("FECHA").Text
                        Case 4
                            CotizacionDolar = Nod.selectSingleNode("HORAUNIX").Text
                    End Select

                Next i
            
            Else
                'Nod.selectSingleNode(vCase).Text = vValor
                'Doc.Save (App.Path & "\configuracion.xml")
            
            End If
        
        Next Nod
        Screen.MousePointer = vbDefault
    Else
        MsgBox "No se puede abrir el archivo " & App.Path & "\configuracion.xml", vbCritical, "Error"
    End If
   
End Function
Public Sub EjecutarScript(ByVal vsql As String, Optional vPathDB)
On Error Resume Next

    Dim connScripts As New ADODB.Connection
    
    With connScripts
        If IsMissing(vPathDB) = True Then
            .ConnectionString = pathDBMySQL
        Else
            .ConnectionString = vPathDB
        End If
        .Open
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
    
        Call .Execute(vsql)
    
    End With

    If connScripts.State = 1 Then
        connScripts.Close
        Set connScripts = Nothing
    End If
    
If Err Then
    
'    MsgBox "Ocurrio un error en la Tabla: " + Chr(13) + vsql + Chr(13) + " > Consulte con el soporte técnico de este sistema", vbCritical, "Error..."
   ' GrabarLog "EjecutarScript", Err.Number & " " & Err.Description, "Global"
   ' Call alerta("Error al " + vSQL, 900)
Else
'Call alerta(vSQL, 1000)
End If
End Sub



Public Function getSqlTotal(ByVal vsql As String, Optional vPathDB) As String
On Error Resume Next

    Dim connScripts As New ADODB.Connection
    Dim rec99 As New ADODB.Recordset
    Dim resultado As String
    
    resultado = ""
    
    
    Call rec99.Open(vsql, vPathDB, adOpenDynamic, adLockReadOnly)
    
    If rec99.EOF Then
        resultado = ""
        getSqlTotal = ""
        Exit Function
    End If
    
    
    Do Until rec99.EOF

            If Val(rec99.Fields(0)) > 0 And Not rec99.Fields(0).Name = "impuesto_inmobiliario" Then
                resultado = Str(Val(resultado) + Val(rec99.Fields(0)))
            Else
                resultado = resultado + " - " + rec99.Fields(0)
            End If
        
            rec99.MoveNext
    
    Loop
    
    
    getSqlTotal = resultado
    
If Err Then
    
'    MsgBox "Ocurrio un error en la Tabla: " + Chr(13) + vsql + Chr(13) + " > Consulte con el soporte técnico de este sistema", vbCritical, "Error..."
   ' GrabarLog "EjecutarScript", Err.Number & " " & Err.Description, "Global"
   ' Call alerta("Error al " + vSQL, 900)
Else
'Call alerta(vSQL, 1000)
End If
End Function


Public Sub BorrarEnTabla(vtabla As String, vCondicion As String, Optional vPathDB)
On Error Resume Next
    Dim connScripts As New ADODB.Connection
    
    With connScripts
        If IsMissing(vPathDB) = True Then
            .ConnectionString = pathDBMySQL
        Else
            .ConnectionString = vPathDB
        End If
        .Open
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
    
    
    Dim vsql As String
    
    
    vsql = "DELETE FROM " + Trim(vtabla) + " WHERE " + Trim(vCondicion)
    
    Call .Execute(vsql)
    
    End With

    If connScripts.State = 1 Then
        connScripts.Close
        Set connScripts = Nothing
    End If
    
If Err Then
    GrabarLog "EjecutarScript", Err.Number & " " & Err.Description, "Global"
End If
End Sub
Public Sub ActualizarEnTabla(vtabla As String, vcampos As String, vvalores As String, vCondicion As String, Optional vPathDB)
On Error Resume Next

' ----------- Formateo parametros -------------------------------------------------------------
If vCondicion = "" Then vCondicion = " 1=1 "
'----------------------------------------------------------------------------------------------

    Dim connScripts As New ADODB.Connection
    
    With connScripts
        If IsMissing(vPathDB) = True Then
            .ConnectionString = pathDBMySQL
        Else
            .ConnectionString = vPathDB
        End If
        .Open
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
  
' --------------------- split ----------------------------------------------------------------
            Dim vCampoArray() As String
            vCampoArray() = Split(vcampos, ",")
            
            Dim vValoresArray() As String
            vValoresArray() = Split(vcampos, ",")
                      
            Dim vParte As String
            
            vParte = ""
            Dim i As Integer
           'Ale: modificar
           ' For i = 1 To Len(vCampoArray())
           '         vParte = vParte + vCampoArray(i) + "=" + vValoresArray(i)
           ' Next
' ----------------------------------------------------------------------------------------------
    
    Dim vsql As String
    vsql = "UPDATE" + Trim(vtabla) + " SET (" + vParte + ") WHERE " + vCondicion
    Call .Execute(vsql)
    
    If connScripts.State = 1 Then
        connScripts.Close
        Set connScripts = Nothing
    End If
  End With
If Err Then
    GrabarLog "EjecutarScript", Err.Number & " " & Err.Description, "Global"
End If
End Sub


Public Function traerDatosLista(ByVal vsql As String, ByVal vPathDB As String, ByRef vlista() As String) As String
Dim vsql1 As String
Dim v As String


    Dim rsDato As New ADODB.Recordset
    'traerDatos2 = ""
    
    
    With rsDato
        DoEvents
        
        If IsMissing(vPathDB) = True Then
            Call .Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(vsql, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        Dim j, i As Integer
        j = 0
        
        Do Until .EOF
        
            For i = 0 To .Fields.Count - 1
                v = .Fields(i) + ","
            Next
            
            j = j + 1
            vlista(j) = v  ' escribe la linea
            
        Loop
        
        
    End With
    
  '  sqlDato = ""
    
    If rsDato.State = 1 Then
        rsDato.Close
        Set rsDato = Nothing
    End If

If Err Then traerDatosLista = "ERROR"

End Function



Public Function traerDatos2(ByVal vsql As String, ByVal vcampo As String, ByVal vPathDB As String)
On Error Resume Next

    Dim rsDato As New ADODB.Recordset
    traerDatos2 = ""
    With rsDato
        DoEvents
        
        If IsMissing(vPathDB) = True Then
            Call .Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(vsql, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        If .State = 1 Then
            If Not .EOF = True Then
                traerDatos2 = Trim(.Fields(vcampo).Value & " ")
            Else
                traerDatos2 = ""
            End If
        Else
            traerDatos2 = ""
            Err.Clear
        End If
        
    End With
    
  '  sqlDato = ""
    
    If rsDato.State = 1 Then
        rsDato.Close
        Set rsDato = Nothing
    End If
    
'If Err Then GrabarLog "TraerDato", Err.Number & " " & Err.Description, "Procedimientos"
If Err Then traerDatos2 = 0
End Function


Public Function TraerDato(vtabla, vfiltro, vcampo, Optional vPathDB) As String
On Error Resume Next

    Dim rsDato As New ADODB.Recordset
    Dim sqlDato As String
    
    sqlDato = "SELECT * FROM " & vtabla & " WHERE " & vfiltro
    
    With rsDato
        DoEvents
        
        If IsMissing(vPathDB) = True Then
            Call .Open(sqlDato, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(sqlDato, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        If .State = 1 Then
            If Not .EOF = True Then
                TraerDato = Trim(.Fields(vcampo).Value & " ")
            Else
                TraerDato = ""
            End If
        Else
            TraerDato = ""
            Err.Clear
        End If
        
    End With
    
    sqlDato = ""
    
    If rsDato.State = 1 Then
        rsDato.Close
        Set rsDato = Nothing
    End If
    
If Err Then
GrabarLog "TraerDato", Err.Number & " " & Err.Description, "Procedimientos"
Exit Function
End If
End Function
Public Sub LlenarGrilla(vtabla As String, vGrilla As KlexGrid, vfiltro As String, vTamanos As String, Optional vPathDB)
On Error Resume Next
Dim CnTemp As New ADODB.Connection
Dim RecTemp As New ADODB.Recordset

' ejecuto el script
'Dim connScripts As New ADODB.Connection
    

    
    With CnTemp
        If IsMissing(vPathDB) = True Then
            .ConnectionString = pathDBMySQL
        Else
            .ConnectionString = vPathDB
        End If
        .Open
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
    
        
        RecTemp.Open vfiltro, CnTemp, adOpenKeyset, adLockOptimistic
            
    End With

    Set vGrilla.Recordset = Nothing
    
    
    If RecTemp.EOF Then
         vGrilla.Visible = False
         vGrilla.Tag = ""
    Else
        vGrilla.Visible = True
        vGrilla.Tag = "visible"
        Set vGrilla.Recordset = RecTemp
         
    End If
    
If Err > 0 Then
    GrabarLog "EjecutarScript", Err.Number & " " & Err.Description, "Global"
    Exit Sub
End If

End Sub

Public Sub LlenarGrilla2(vGrilla As KlexGrid, vsql As String, vTamanos As String, Optional vPathDB)
On Error Resume Next
Dim CnTemp As New ADODB.Connection
Dim RecTemp As New ADODB.Recordset

' ejecuto el script
'Dim connScripts As New ADODB.Connection
    

    
    With CnTemp
        If IsMissing(vPathDB) = True Then
            .ConnectionString = pathDBMySQL
        Else
            .ConnectionString = vPathDB
        End If
        .Open
        If .State = 0 Then
            MsgBox Err.Description
            Exit Sub
        End If
    
        
        RecTemp.Open vsql, CnTemp, adOpenKeyset, adLockOptimistic
            
    End With

    Set vGrilla.Recordset = Nothing
    
    
    If RecTemp.EOF Then
         vGrilla.Visible = False
         vGrilla.Tag = ""
    Else
        vGrilla.Visible = True
        vGrilla.Tag = "visible"
        Set vGrilla.Recordset = RecTemp
         
    End If
    
If Err > 0 Then
    GrabarLog "EjecutarScript", Err.Number & " " & Err.Description, "Global"
    Exit Sub
End If

End Sub

Public Function formatNumero(vnumero) As String
    formatNumero = Format(vnumero, "###,###,##0.00")
End Function


Public Function GenerarDato(vsql, vcampo, Optional vPathDB) As String
On Error Resume Next

    Dim rsDato As New ADODB.Recordset, sqlDato As String

    sqlDato = vsql
    
    With rsDato
        .CursorLocation = adUseClient
        
        If IsMissing(vPathDB) = True Then
            Call .Open(sqlDato, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(sqlDato, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        If Not .EOF = True Then
            GenerarDato = Trim(.Fields(vcampo).Value & " ")
        Else
            GenerarDato = ""
        End If
        
    End With
    
    
    sqlDato = ""
    
    rsDato.Close
    Set rsDato = Nothing

If Err Then GrabarLog "GenerarDato", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function AjustarMes(vmes) As String
On Error Resume Next
    
    If vmes <= 9 Then
        AjustarMes = "0" & vmes
    Else
        AjustarMes = vmes
    End If
    
If Err Then GrabarLog "AjustarMes", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function DiasDelMes(Optional ByVal fecha As Variant) As Integer
On Error Resume Next

    Dim mes As Integer, y  As Integer

   If IsMissing(fecha) Then fecha = Date

    If IsDate(fecha) Then
        y = Year(fecha)
        mes = Month(fecha)
    ElseIf IsNumeric(fecha) Then
        y = Year(Date)
        mes = IIf(fecha > 0 And fecha < 13, CInt(fecha), 0)
    ElseIf VarType(fecha) = vbString Then
        y = Year(Date)
        Select Case UCase(Left$(fecha, 3))
            Case "FEB":                                             mes = 2
            Case "ENE", "MAR", "MAY", "JUL", "AGO", "OCT", "DIC":   mes = 1
            Case "ABR", "JUN", "SEP", "NOV":                        mes = 4
        End Select
    End If

    Select Case mes
        Case 2:                     DiasDelMes = IIf(saltarYear(fecha), 29, 28)
        Case 1, 3, 5, 7, 8, 10, 12: DiasDelMes = 31
        Case 4, 6, 9, 11:           DiasDelMes = 30
    End Select

If Err Then GrabarLog "DiaDelMes", Err.Number & " " & Err.Description, "Global"
End Function
Public Function UltimaHoja(vOpen As Boolean, vIva As String, Optional vValor As Integer) As Integer
    On Error Resume Next
    
    Dim rsHoja As New ADODB.Recordset, sqlHoja As String
    
    sqlHoja = "SELECT * FROM Configura WHERE 1= 1"
    
    With rsHoja
        
        If vOpen = True Then
            .Open sqlHoja, ConnDDBB, adOpenDynamic, adLockPessimistic
                        
            .Fields(vIva).Value = vValor
            .Update
        
        Else
            .Open sqlHoja, ConnDDBB, adOpenStatic, adLockReadOnly
            If Not .EOF Then
                UltimaHoja = .Fields(vIva).Value
            End If
        End If
        
    
    End With
    
    sqlHoja = ""
    
    rsHoja.Close
    Set rsHoja = Nothing

If Err Then GrabarLog "UltimaHoja", Err.Number & " " & Err.Description, "Procedimientos"
End Function
Public Function saltarYear(ByVal valor As Variant) As Boolean
    On Error Resume Next

    Dim iYear As Integer
    
    If IsDate(valor) Then iYear = Year(valor) Else iYear = CInt(valor)

    If TypeName(iYear) = "Integer" Then
        saltarYear = Day(DateSerial(iYear, 3, 0)) = 29
    End If

If Err Then GrabarLog "SaltarYear", Err.Number & " " & Err.Description, "Global"
End Function
Public Sub temp(Codigo, Nombre, saldo, fecha, NumHoja)
On Error Resume Next
    
    Dim rstemp As New ADODB.Recordset, sqlTemp As String

    sqlTemp = "SELECT * FROM Temp WHERE (NumHoja = " & NumHoja & ")"

    With rstemp
        .Open sqlTemp, ConnDDBB, adOpenDynamic, adLockPessimistic
        
        If .EOF = True Then
            .AddNew
            
            .Fields("Codigo").Value = Codigo
            .Fields("Nombre").Value = Nombre
            .Fields("Saldo").Value = saldo
            .Fields("Fecha").Value = fecha
            .Fields("NumHoja").Value = NumHoja
            
            .Update
        End If
    
    End With

    sqlTemp = ""
    
    rstemp.Close
    Set rstemp = Nothing

If Err Then GrabarLog "Temp", Err.Number & " " & Err.Description, "Procedimientos"
End Sub
Public Function Login(vServidor, vempresa, vUsuario, vPassword) As String
On Error Resume Next

    Login = ""
    
    With vConfigGral
               
     
        .vIdEmpresa = TraerDato("Empresas", "Alias = '" & Trim(vempresa) & "'", "idEmpresas", PathDBConfig)
        
        
        .vIdServidor = TraerDato("Servidor", "Servidor = '" & Trim(vServidor) & "'", "idServidor", PathDBConfig)
        
        .vIdUsuario = TraerDato("Usuarios", "Usuario = '" & Trim(vUsuario) & "'", "idUsuarios", PathDBConfig)
        
      

        'Controlo...... (Servidor/Empresa/Usuario)
       ' MsgBox .vIdEmpresa + .vIdServidor + .vIdUsuarioa
     
        If Val(TraerDato("EmpresasAsociadas", "(idServidor = " & .vIdServidor & ") AND (idEmpresas = " & .vIdEmpresa & ") AND (idUsuarios = " & .vIdUsuario & ")", "idEmpresasAsociadas", PathDBConfig)) > 0 Then
                
        mensaje "7"
                
            If Not .vIdEmpresa = 0 And Not .vIdServidor = 0 And Not .vIdUsuario = 0 Then
                .vempresa = Trim(vempresa)
                .vServidor = Trim(vServidor)
                .vUser = Trim(vUsuario)

                .vDireccionDB = TraerDato("Servidor", "idServidor = " & .vIdServidor & "", "Direccion", PathDBConfig)
                .vUserDB = TraerDato("Servidor", "idServidor = " & .vIdServidor & "", "User", PathDBConfig)
                .vPassDB = TraerDato("Servidor", "idServidor = " & .vIdServidor & "", "Pass", PathDBConfig)
        mensaje "8"
                'If Encriptar(Trim(vPassword), LeerConfig(0)) = TraerDato("Usuarios", "idUsuarios = " & (.vIdUsuario) & "", "Password", PathDBConfig) Then
                 If Trim(vPassword) = TraerDato("Usuarios", "idUsuarios = " & (.vIdUsuario) & "", "Password", PathDBConfig) Then
                    
                    .vPass = Trim(vPassword)
                    Login = "Correcto"
                
                Else
                    Login = "Password"
                End If
            
            Else
                'No Pudo comprobarse algun dato
            
            End If
        
        Else
            
            Login = "Empresa-Usuario"
        
        End If
        
        
        .vComunaCnn = LeerXml("ComunaCnn")
       
        .vIncluyeContabilidad = CBool(LeerConfig(9))
        .vIncluyeStac = CBool(LeerConfig(10))
        .vIncluyeResto = CBool(LeerConfig(11))
        
        .vIncluyeTicket = CBool(LeerConfig(12))
        
        .vIncluyeCobros = EsNulo(LeerConfig(13))
        .vImpresoraSeleccionada = LeerXml("Impresora")
        
    
        .vImprimirReciboCliente = EsNulo(LeerConfig(15))
        .vImprimirReciboProveedor = EsNulo(LeerConfig(16))
        
        .vTipoCliente = EsNulo(LeerConfig(30))
        
        mensaje "9"
        If TraerDato("Empresas", "idEmpresas = " & Trim(vConfigGral.vIdEmpresa) & "", "Principal", PathDBConfig) = "S" Then
            .vEmpresaPrincipal = True
        Else
            .vEmpresaPrincipal = False
        End If
        
        mensaje "10 - vTipoCliente: " + .vTipoCliente
        
    End With
    
    
   ' If Not Login = "Correcto" Then Exit Function
    
    Set ConnDDBB = New ADODB.Connection
    
    With ConnDDBB
       .ConnectionString = pathDBMySQL
       .CursorLocation = adUseClient
       .Open
       
       mensaje pathDBMySQL
       
       If .State = 0 Then
           ' MsgBox Err.Description
           ' End
        End If
    End With
    
    'No se usa hasta no empezar con el MySQL
    'vDSN = IniciaDSN(vConfigGral.vempresa)
    
    mensaje "11 - " + pathDBMySQL
    
    
    
    Set ConnComunaDB = New ADODB.Connection
    
    With ConnComunaDB
       .ConnectionString = vConfigGral.vComunaCnn
       .CursorLocation = adUseClient
       .Open
       
       mensaje "12- " + vConfigGral.vComunaCnn
       
       mensaje pathDBMySQL
       
       If .State = 0 Then
            mensaje "13- Error interno :" + Err.Description + " - " + Str(Err.Number)
            
           ' Exit Function
        End If
    End With
    
    
    
       mensaje "14- Última linea de login " + Chr(13) + pathDBMySQL
       
       
If Err Then

    mensaje "Error : " + Err.Description
    MsgBox Err.Number + " " + Err.Description
    
    GrabarLog "Login", Err.Number & " " & Err.Description, "Global"
End If
End Function
Public Sub DatosEmpresa()
On Error Resume Next

    'Cambiar y Dejar un solo recordset

    With vDatosEmpresa
        .Nombre = TraerDato("Empresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "Empresa", PathDBConfig)
        .Alias = TraerDato("Empresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "Alias", PathDBConfig)
        .CondicionIva = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "CondicionIva", PathDBConfig)
        .cuit = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "Cuit", PathDBConfig)
        .Direccion = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "Direccion", PathDBConfig)
        .Localidad = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "Localidad", PathDBConfig)
        .Telefono = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "Telefono", PathDBConfig)
        .Email = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "EMail", PathDBConfig)
        .WebSite = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "WebSite", PathDBConfig)
        .Responsable = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "Responsable", PathDBConfig)
        .UsarNroInterno = TraerDato("DatosEmpresas", "idEmpresas = " & vConfigGral.vIdEmpresa & "", "UsarNroInterno", PathDBConfig)
    End With
    
If Err Then
mensaje "Error al carga datos de la empresa. " + Chr(13) + Err.Description
GrabarLog "DatosEmpresa", Err.Number & " " & Err.Description, "Global"
End If
End Sub
Public Sub CargarImpresoras()
On Error Resume Next

    Dim rsImpresoras As New ADODB.Recordset, sqlImpresoras As String
    
    sqlImpresoras = "SELECT * FROM Impresoras WHERE (idImpresoras = 1)"
    
    With rsImpresoras
        .CursorLocation = adUseClient
        Call .Open(sqlImpresoras, PathDBConfig, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
        
            vImpresoras.vNombreImpresora = EsNulo(.Fields("Impresora").Value)
            vImpresoras.vModelo = EsNulo(.Fields("Modelo").Value)
            vImpresoras.vModeloInterno = EsNulo(.Fields("ModeloInterno").Value)
            vImpresoras.vNroPuerto = EsNulo(.Fields("Puerto").Value)
            vImpresoras.vEsFiscal = EsNulo(.Fields("EsFiscal").Value)
            vImpresoras.vPorDefecto = EsNulo(.Fields("PorDefecto").Value)
    
        End If
        
    End With
    
    sqlImpresoras = ""
    
    If rsImpresoras.State = 1 Then
        rsImpresoras.Close
        Set rsImpresoras = Nothing
    End If
    
If Err Then GrabarLog "CargarImpresoras", Err.Number & " " & Err.Description, "Global"
End Sub
Function ValidarNumero(Objeto As Object, KeyAscii As Integer)
On Error Resume Next

    'ConPunto nos dice si es un tipo de dato Integer o Double
    Select Case KeyAscii
        Case 48, 49, 50, 51, 52, 53, 54, 55, 56, 57
            '0=48; 1=49; 2=50; 3=51; 4=52; 5=53; 6=54; 7=55; 8=56; 9=57
        Case 8   'Delete
        Case Else
            KeyAscii = 13
    End Select

If Err Then GrabarLog "ValidarNumero", Err.Number & " " & Err.Description, "Global"
End Function
Public Sub ImprimirFormularios()
On Error Resume Next
    
    Dim rsFormularios As New ADODB.Recordset, sqlFormularios As String
    
    sqlFormularios = "SELECT * FROM Formularios"
    
    With rsFormularios
        Call .Open(sqlFormularios, PathDBConfig, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then .MoveFirst
            
        Do Until .EOF = True
        
            Debug.Print "Case " & Chr(34) & .Fields("Formulario").Value & Chr(34)
            Debug.Print .Fields("Formulario").Value & ".Show"
            Debug.Print .Fields("Formulario").Value & ".Tag = vIdFormularioActivo"
            Debug.Print " "
            
            .MoveNext
        Loop
    
    End With
    
    sqlFormularios = ""

    If rsFormularios.State = 1 Then
        rsFormularios.Close
        Set rsFormularios = Nothing
    End If
    
If Err Then GrabarLog "ImprimirFormularios", Err.Number & " " & Err.Description, "Global"
End Sub
Public Function TraerPropiedadExcel(vFile, vPropiedad As String) As Variant
On Error Resume Next

    Dim vObjExcel As Object, vObjHoja As Object

    Set vObjExcel = CreateObject("Excel.Application")
  ' Set vObjExcel = CreateObject("Word.Application")
  ' Set vObjExcel = CreateObject("Excel.Application.11")
  
    vObjExcel.Workbooks.Open FileName:=vFile
  
    If Val(vObjExcel.Application.Version) >= 8 Then
        Set vObjHoja = vObjExcel.ActiveSheet
    Else
        Set vObjHoja = vObjExcel
    End If
        
    TraerPropiedadExcel = vObjExcel.ActiveSheet.Name
        
    Set vObjHoja = Nothing
    Set vObjExcel = Nothing

If Err Then
MsgBox "Error: " + Str(Err)
End If
End Function
Public Function AddControl(Controls As CommandBarControls, ControlType As XTPControlType, id As Long, Caption As String, Optional BeginGroup As Boolean = False, Optional DescriptionText As String = "", Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional Category As String = "Controls") As CommandBarControl
    Dim control As CommandBarControl
    Set control = Controls.Add(ControlType, id, Caption)
    
    control.BeginGroup = BeginGroup
    control.DescriptionText = DescriptionText
    control.Style = ButtonStyle
    control.Category = Category
    
    Set AddControl = control
    
End Function

Public Function EjecutarConsulta(S As String) As Recordset
    
    Dim rs As New ADODB.Recordset
       
    Dim cmd As New ADODB.Command
    
    cmd.ActiveConnection = ConnDDBB
    Dim sql As String
    
    sql = S
  
    If rs.State = 0 Then
        rs.Open sql, ConnDDBB, adOpenKeyset, adLockOptimistic
    Else
        Set rs = ConnDDBB.Execute(sql)
    End If
    
    Set EjecutarConsulta = rs
    
End Function

Public Function CalcularSaldo(vnroremito As Long) As Double
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & vnroremito & ")"
     
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        Dim saldo, debito, credito As Double
        Do While Not .EOF
            If IsNull(.Fields("debito")) Then
                debito = 0
            Else
                debito = .Fields("debito")
            End If
            
            If IsNull(.Fields("credito")) Then
                credito = 0
            Else
                credito = .Fields("credito")
            End If
                        
            saldo = saldo + debito - credito
            
            .MoveNext
        Loop
        
        CalcularSaldo = saldo
        
    End With

If Err Then GrabarLog "wcorrientes", Err.Number & " " & Err.Description, "Global"
End Function

Public Function CalcularTotal(vnroremito As Long) As Double
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & vnroremito & ")"
     
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        Dim total, debito As Double
        Do While Not .EOF
            If IsNull(.Fields("debito")) Then
                debito = 0
            Else
                debito = .Fields("debito")
            End If
            
            'Hago la suma por si el monto de la factura se actualizo mediante alguna nota de debito
            total = total + debito
            
            .MoveNext
        Loop
        
        CalcularTotal = total
        
    End With

If Err Then GrabarLog "wcorrientes", Err.Number & " " & Err.Description, "Global"
End Function

Public Function CalcularPagado(vnroremito As Long) As Double
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & vnroremito & ")"
     
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        Dim pagado, total As Double
        Do While Not .EOF
            If IsNull(.Fields("credito")) Then
                pagado = 0
            Else
                pagado = .Fields("credito").Value
            End If
            
            'Hago la suma por si el monto de la factura se actualizo mediante alguna nota de debito
            total = total + pagado
            
            .MoveNext
        Loop
        
        CalcularPagado = total
        
    End With

If Err Then GrabarLog "CalcularPagado", Err.Number & " " & Err.Description, "Global"
End Function

Public Function ObtenerCodigoClienteDesdeCtaCte(remito As Long) As String
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & remito & ")"
     
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 And Not (.EOF = True) Then
            .MoveFirst
            ObtenerCodigoClienteDesdeCtaCte = .Fields("Codigo").Value
         End If
       
    End With

    sqlCtaCteC = ""

    If rsCtaCteC.State = 1 Then
        rsCtaCteC.Close
        Set rsCtaCteC = Nothing
    End If
    
If Err Then GrabarLog "ObtenerCodigoClienteDesdeCtaCte", Err.Number & " " & Err.Description, "Global"
End Function
Public Function ObtenerNombreClienteDesdeCtaCte(remito As Long) As String
    On Error Resume Next
    
    Dim rsCtaCteC As New ADODB.Recordset, sqlCtaCteC As String
    
    sqlCtaCteC = "SELECT * FROM CuentasCorrientes WHERE (remito = " & remito & ")"
     
    With rsCtaCteC
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteC, ConnDDBB, adOpenStatic, adLockPessimistic)
        .MoveFirst
        ObtenerNombreClienteDesdeCtaCte = .Fields("Nombre")
        
    End With
End Function
Public Function ObtenerCotizacionMoneda(tm As String, esVenta As Boolean) As Double
    On Error Resume Next
    
    Dim rs As New ADODB.Recordset, sql As String
    
    sql = "SELECT * FROM tipomoneda WHERE (idtipomoneda = '" & tm & "')"
         
    With rs
        .CursorLocation = adUseClient
               
        Call .Open(sql, ConnDDBB, adOpenStatic, adLockPessimistic)
        .MoveFirst
        
        If esVenta Then
            ObtenerCotizacionMoneda = .Fields("ValorVenta").Value
        Else
            ObtenerCotizacionMoneda = .Fields("ValorCompra").Value
        End If
        
    End With
End Function
Public Sub PagarCtaCteProveedor2(vnroremito As Long, importe As Double, fecha As Date, idMedioPago)
    On Error Resume Next
    
    Dim rsCtaCteP As New ADODB.Recordset, sqlCtaCteP As String
    Dim SaldoAnterior, debito, credito As Double
    Dim comentario, TipoDocumento As String
    
    sqlCtaCteP = "SELECT * FROM PCuentasCorrientes WHERE (remito = " & vnroremito & ")"
    
    With rsCtaCteP
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
               
        Call .Open(sqlCtaCteP, ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If .State = 1 And Not .EOF = True Then
                
            .MoveFirst
        
            comentario = EsNulo(.Fields("comentario").Value)
        
            'TipoDocumento = .Fields("tipodocumento")
        
            Do While Not .EOF
                If IsNull(.Fields("debito").Value) Then
                    debito = 0
                Else
                    debito = Val(Format(.Fields("debito"), "#####0.00"))
                End If
            
                If IsNull(.Fields("credito").Value) Or (.Fields("credito").Value = 0) Then
                    credito = 0
                Else
                    credito = Val(Format(.Fields("credito"), "#####0.00"))
                End If
                        
                SaldoAnterior = Val(Format(SaldoAnterior, "#####0.00")) + Val(Format(debito, "#####0.00")) - Val(Format(credito, "#####0.00"))
            
                .MoveNext
            Loop
        
            If .RecordCount > 0 Then .MoveLast
        
            'If .EOF = True Then
            .AddNew
            .Fields("remito").Value = Trim(vnroremito)
            .Fields("comentario").Value = "Pago : " & comentario
            'End If
        
            .Fields("Fecha").Value = fecha
            '.Fields("Fechainput").value = fecha
        
            'Buscar el cliente segun el remito
            .Fields("Codigo").Value = TraerDato("PCuentasCorrientes", "Remito = " & Val(vnroremito) & "", "Codigo")
            .Fields("Nombre").Value = TraerDato("PCuentasCorrientes", "Remito = " & Val(vnroremito) & "", "Nombre")
       
            '.Fields("anomes").value = Right(.Fields("Fecha").value, 4) & Mid(.Fields("Fecha").value, 4, 2)
    
            .Fields("debito") = 0 'Val(Format(.Fields("debito"), "#####0.00")) - Val(Format(importe, "#####0.00"))
            .Fields("credito") = Val(Format(importe, "#####0.00"))
            .Fields("saldo") = SaldoAnterior - Val(Format(.Fields("credito").Value, "#####0.00"))
                            
            .Fields("idMedioPago") = idMedioPago
                                        
            .Fields("TipoMovimiento") = "RC"
            
            .Update
        
        End If
    
    End With

    sqlCtaCteP = ""

    If rsCtaCteP.State = 1 Then
        rsCtaCteP.Close
        Set rsCtaCteP = Nothing
    End If

If Err Then GrabarLog "PagarCtaCte", Err.Number & " " & Err.Description, "Global"
End Sub
Public Sub CambiarImpresora(vTipoImpresora As String, vImpresora As String)
On Error Resume Next

    Dim di

    Select Case vTipoImpresora
    
        Case "Fiscal"
            'di = WriteProfileString("WINDOWS", "DEVICE", "\\SERVIDOR\Epson LX-810,winspool,Ne00:")
            Call SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
        
        Case "Matriz"
            'di = WriteProfileString("WINDOWS", "DEVICE", "\\SERVIDOR\Epson LX-810,winspool,Ne00:")
            Call SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")

        Case Else
            'di = WriteProfileString("WINDOWS", "DEVICE", "\\SERVIDOR\Epson LX-810,winspool,Ne00:")
            Call SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
                        
    End Select

If Err Then GrabarLog "CambiarImpresora", Err.Number & " " & Err.Description, "Global"
End Sub
Public Function ValidarCuit(vcuit As String) As Boolean
    On Error Resume Next
    
    Dim vLongitud As String, vLineas As String, vHayLineas As Boolean
    
    ValidarCuit = True
    
    vLongitud = Len(vcuit)
    
    If Val(vLongitud) = 11 Or Val(vLongitud) = 13 Then
        vHayLineas = CBool(InStr(1, vcuit, "-"))
    Else
        ValidarCuit = False
        Exit Function
    End If
    
    If (vHayLineas = False) And Not (vLongitud = 13) Then
        ValidarCuit = False
        Exit Function
    End If
    
    If Not VerificarCuit(vcuit) = True Then
        ValidarCuit = False
        Exit Function
    End If
    
    If 1 = 2 Then
        'Falta comprobar cuando modifica...
        If Not TraerDato("Clientes", "Cuit = '" & vcuit & "'", "idClientes") = "" Then
            ValidarCuit = False
            Exit Function
        End If
    End If
    
    
    If Err Then GrabarLog "ValidarCuit", Err.Number & " " & Err.Description, "Global"
End Function
Private Function VerificarCuit(ByRef vcuit As String) As Boolean
On Error Resume Next

    Dim Sen As Integer, S As Integer, i As Integer
    Dim Coef As String, cuit As String
    Dim r%

    Coef = "5432765432"  ' Destinado a la verificación del CUIT

    cuit = Left(vcuit, 2) + Mid(vcuit, 4, 8) + Right(vcuit, 1)

    S = 0
    
    'Efectuo suma en S de los productos de cada dígito del Coeficiente (COEF) * los del CUIT
    For i = 1 To 10
        S = S + Val(Mid(cuit, i, 1)) * Val(Mid(Coef, i, 1))
    Next
 
    r = S Mod 11 ' Averiguo el remanente de dividir S / 11
 
    If r > 1 Then   ' Si el remanente es > 1 divido en R 11/R
        r = 11 - r
    End If
 
    If r = Right(cuit, 1) Then  ' Si R=Código de verificación (dígito derecho del CUIT)
        VerificarCuit = True
    Else
        VerificarCuit = False
    End If
 
If Err Then GrabarLog "VerificarCuit", Err.Number & " " & Err.Description, "Global"
End Function
Public Function ContarCaracteres(ByVal cadena As String, ByVal caracter As String) As Integer

  Dim n As Integer
  Dim contador As Integer
  contador = 0
  For n = 1 To Len(cadena)
     If Mid(cadena, n, 1) <> caracter Then
        contador = contador + 1
      Else
        ContarCaracteres = contador
        Exit Function
     End If
  Next n
  If contador > 0 Then
    ContarCaracteres = contador
  Else
    ContarCaracteres = Len(cadena)
  End If
End Function
Public Function GenerarPass() As String
On Error Resume Next

    Randomize Timer
    
    Dim PasswordCreate As Integer
    Dim newPassword As String
    Dim newChar As Integer

    newPassword = ""
    
    For PasswordCreate = 1 To 6
        newChar = Int((Rnd * 255))
        
        While Not CheckCharacter(newChar)
            newChar = Int((Rnd * 255))
        Wend
        
        newPassword = newPassword & Chr(newChar)
    
    Next
    
    GenerarPass = newPassword

If Err Then GrabarLog "GenerarPass", Err.Number & " " & Err.Description, "Global"
End Function
Private Function CheckCharacter(newChar As Integer) As Boolean
Const StartSymbol1 As Integer = 33
Const EndSymbol1 As Integer = 47

Const StartNumeric As Integer = 48
Const EndNumeric As Integer = 57

Const StartSymbol2 As Integer = 58
Const EndSymbol2 As Integer = 64

Const StartUpper As Integer = 65
Const EndUpper As Integer = 90

Const StartLower As Integer = 97
Const EndLower As Integer = 122

Const StartExtended As Integer = 128
Const EndExtended As Integer = 255

    CheckCharacter = False
    
    Select Case newChar
            ' Case Statement for UpperCase Characters
            Case StartUpper To EndUpper
                CheckCharacter = False
                   
            ' Case Statement for LowerCase Characters
            Case StartLower To EndLower
                CheckCharacter = True
                   
            ' Case Statement for Numerics
            Case StartNumeric To EndNumeric
                CheckCharacter = True
                    
            ' Case Statement for symbols
            Case StartSymbol1 To EndSymbol1, StartSymbol2 To EndSymbol2
                CheckCharacter = False
            
            Case StartExtended To EndExtended
                CheckCharacter = False
    End Select

End Function
Public Function FormatoNroFactura(vNcomp As Long) As String
On Error Resume Next

    FormatoNroFactura = String(8 - Len(Trim(vNcomp)), "0") & vNcomp

If Err Then GrabarLog "FormatoNroFactura", Err.Number & " " & Err.Description, "Global"
End Function
Public Sub VerAyuda(vFormulario As String)
On Error Resume Next


If Err Then GrabarLog "VerAyuda", Err.Number & " " & Err.Description, "BasAyuda"
End Sub


Public Sub ValidadNroChe(vnrocheque As String)
If Val((traerDatos2("select nrocheque from cheques where nrocheque=" + Trim(vnrocheque), "nrocheque", pathDBMySQL))) > 0 Then
    MsgBox "El nro de cheque ingresado fue cargado anteriormente", vbInformation, "Nro de Cheques:"
End If
End Sub

Public Function consultarTarifaria(vidArticulos As Long) As String
Dim vsql, vmensaje As String
Dim pventa, pcosto, tarifa As Double
Dim i As Integer



vsql = "select * from articulos where idArticulos=" + Str(vidArticulos)
pcosto = Val(TraerDato2(vsql, "pcosto", pathDBMySQL))

vmensaje = vmensaje + "> Costo: " + Format(pcosto, "###,##0.00") + Chr(13) + Chr(13)

pventa = Val(TraerDato2(vsql, "pventa1", pathDBMySQL))

vsql = "select * from vistaarticulos where idArticulos=" + Str(vidArticulos)
pventa = pventa * (1 + (Val(TraerDato2(vsql, "Porcentaje", pathDBMySQL)) / 100))

vmensaje = vmensaje + "> Final (Lista1 + IVA): " + Format(pventa, "###,##0.00") + Chr(13) + Chr(13)

vmensaje = vmensaje + "____________________________________" + Chr(13) + Chr(13)

For i = 1 To 5
            pventa = Val(TraerDato2(vsql, "pventa" + Trim(Str(i)), pathDBMySQL))
            If pventa > 0 Then
                tarifa = ((pventa / pcosto) - 1) * 100
            Else
                tarifa = 0
            End If
            vmensaje = vmensaje + "> Lista " + Trim(Str(i)) + ": " + Format(pventa, "###,##0.00") + " - % Tarifa: " + Format(tarifa, "##0.00") + Chr(13) + Chr(13)
Next


vmensaje = vmensaje + "____________________________________"

consultarTarifaria = vmensaje

End Function


Public Function TraerDato2(vsql, vcampo, Optional vPathDB) As String
On Error Resume Next

    Dim rsDato As New ADODB.Recordset
    Dim sqlDato As String
    
    sqlDato = vsql
    
    With rsDato
        DoEvents
        
        If IsMissing(vPathDB) = True Then
            Call .Open(sqlDato, ConnDDBB, adOpenStatic, adLockReadOnly)
        Else
            Call .Open(sqlDato, vPathDB, adOpenStatic, adLockReadOnly)
        End If
        
        If .State = 1 Then
            If Not .EOF = True Then
                TraerDato2 = Trim(.Fields(vcampo).Value & " ")
            Else
                TraerDato2 = ""
            End If
        Else
            TraerDato2 = ""
            Err.Clear
        End If
        
    End With
    
    sqlDato = ""
    
    If rsDato.State = 1 Then
        rsDato.Close
        Set rsDato = Nothing
    End If
    
If Err Then GrabarLog "TraerDato", Err.Number & " " & Err.Description, "Procedimientos"
End Function

Public Sub t_borrarFila(id As Long, vtabla As String)
Dim vsql As String


If MsgBox("Está seguro de borrar la fila seleccionada ?", vbYesNo, "Borrando fila") = vbYes Then

    vsql = "delete from clientes where idClientes=" + Str(id)
    Call EjecutarScript(vsql)
    
End If

End Sub

Public Sub fMostarDocumentosImpagos(vcodigo As String)
' muestra las facturas inpagas


With frmBuscarFactura
    .txtCliente.Tag = vcodigo
    .cmbEstadoDocumento = "Adeudado"
    Call .cmdFiltrar_Click
End With


End Sub
Public Function NroComprobanteNuevo(ByVal vTipoDocumento As String, ByVal vLetra As String, ByVal vpuntoventa As String) As Long

If Not vTipoDocumento = "" And Not vLetra = "" And Not vpuntoventa = "" Then
    NroComprobanteNuevo = Val(EsNulo(GenerarDato("SELECT MAX(convert(NComprobante,unsigned)) AS NComp FROM Factura WHERE (Tipo = '" & vTipoDocumento & "') AND (Letra = '" & vLetra & "') AND (PuntoDeVenta = '" & vpuntoventa & "')", "NComp"))) + 1
End If

End Function


Public Function getNroRecibo() As Long
On Error Resume Next
Dim vsql As String
Dim vn As Long

vsql = "select max(numero) as m from t_nrorecibo"
getNroRecibo = 1
getNroRecibo = traerDatos2(vsql, "m", pathDBMySQL) + 1

vsql = "insert into t_nrorecibo (numero) values (" + Str(getNroRecibo) + ")"

Call EjecutarScript(vsql)

If Err Then
    getNroRecibo = 1
    vsql = "insert into t_nrorecibo (numero) values (" + Str(1) + ")"
End If
End Function


Public Function getMarcaIntarna() As Long
On Error Resume Next
Dim vsql As String
Dim vn As Long

vsql = "select max(numero) as m from t_marcainterna"

vn = traerDatos2(vsql, "m", pathDBMySQL) + 1

If vn > 999 Then
    vsql = "delete from t_marcainterna "
    Call EjecutarScript(vsql)
    getMarcaIntarna = 1
    
    MsgBox "Atención. Comenzamos con la marca interna 1", vbInformation
Else

getMarcaIntarna = vn
End If

If Err Then
getMarcaIntarna = 1
End If
End Function

Public Sub setMarcaInterna(vnumero As String)
Dim vsql As String
Dim vn As Long

vsql = "insert into t_marcainterna (numero) values (" + vnumero + ")"

Call EjecutarScript(vsql)

End Sub



Public Function descuento(vmonto, vdescuento) As Double
    descuento = vmonto * (100 - vdescuento) / 100
End Function

Public Sub Pase_Excel(Str_Sql As String)
Dim i, j As Integer
Dim Int_Columnas As Integer
Dim Int_Filas As Integer
Dim rs_main As New ADODB.Recordset
Dim excelApp As Excel.Application
Dim excellibro As Excel.Workbook
Dim excelhoja As Excel.Worksheet
Dim Conn As ADODB.Connection

Set Conn = New ADODB.Connection
Conn.ConnectionString = pathDBMySQL
Conn.Open

With rs_main
.CursorLocation = adUseClient
.CursorType = adOpenDynamic
.LockType = adLockBatchOptimistic
Set .ActiveConnection = Conn
.Open Str_Sql
End With

Set excelApp = New Excel.Application
Set excellibro = excelApp.Workbooks.Add
Set excelhoja = excellibro.ActiveSheet
Int_Columnas = rs_main.Fields.Count

For i = 1 To Int_Columnas
excelhoja.Cells(1, i) = rs_main.Fields(i - 1).Name
Next

If rs_main.RecordCount > 0 Then
rs_main.MoveFirst
For Int_Filas = 1 To rs_main.RecordCount
For j = 0 To Int_Columnas - 1
If IsNull(rs_main(j).Value) Then
excelhoja.Cells(Int_Filas + 2, j + 1) = ""
Else
excelhoja.Cells(Int_Filas + 2, j + 1) = CStr(rs_main(j).Value)
End If
Next
rs_main.MoveNext
Next
End If

excelApp.Visible = True

End Sub


Public Function ImprimirFlex(MSFlexGrid As Object, enca1 As String, enca2 As String)
      
      
    Dim Desde As Integer
    Dim Hasta As Integer
    Dim Copias As Integer
    Dim Orientation As Integer
    Dim i As Long
    Dim curi, curix As Long
    Dim Columna As Integer
    Dim j As Integer
    Dim x As Long
   
    
    On Error Resume Next
      

    
    
    For i = 1 To 1
        ' fuente y escala
        Printer.FontSize = 8
        Printer.ScaleMode = 1
        curi = 200
        curix = 0
        Columna = MSFlexGrid.Cols
        j = Printer.CurrentY
        ' recorre todas las filas
        
        Printer.Print enca1
        curi = curi + 240
        
        Printer.Print enca2
        
        
        For x = 0 To MSFlexGrid.Rows - 1
              
            If x <> 0 And x Mod 60 = 0 Then
                curi = 200
                Printer.NewPage
            Else
              
            curi = curi + 240
              
            End If
            ' recorre las columnas
            For j = 1 To Columna
                curix = MSFlexGrid.ColPos(j)
                ' posición x e y donde imprimir
                Printer.CurrentY = curi
                Printer.CurrentX = curix
                Printer.FontBold = False
                ' imprime
                
                
                If Val(Left(MSFlexGrid.TextMatrix(x, j), 2)) > 0 Then
                
                    Printer.Print Format(Format(MSFlexGrid.TextMatrix(x, j), "#,###,##0.00"), "@@@@@@@@@@@@")
                
                Else
                
                    Printer.Print MSFlexGrid.TextMatrix(x, j)
                End If
                
            Next
              
            curix = 0
          
        Next
        'Printer.Print ""
        'Printer.ForeColor = vbRed
        'Printer.FontItalic = True
        'Printer.Print "Total: " & MSFlexGrid.Rows - 1 & " registros"
        Printer.EndDoc ' manda el trabajo a la impresora
  
    Next
  
        Exit Function
          
error_Func:
     
   MsgBox Err.Description
   Exit Function
End Function



Public Sub formatGrilla(ByRef vg As MSHFlexGrid, Col As Integer, vformato As String)
Dim i As Integer


For i = 1 To vg.Rows - 1

    vg.TextMatrix(i, Col) = Format(vg.TextMatrix(i, Col), vformato)

Next

End Sub


Public Function strBool(S As String) As String

Select Case S
    Case "Verdadero"
        strBool = "True"
    Case "Falso"
        strBool = "False"
End Select

End Function


Public Function generarExcel(strArchivo As String, Vsgrd As MSHFlexGrid) As Boolean
On Error GoTo err_generarExcel


Dim xlApp As Excel.Application
Dim xlBook As Excel.Workbook
Dim xlSheet As Excel.Worksheet
Dim intFila As Integer
Dim intCol As Integer
generarExcel = False

'Verificar que se haya ingresado un nombre para el archivo.
If Len(Trim(strArchivo)) <= 0 Then
    MsgBox "Debe ingresar un nombre para el archivo."
Exit Function
End If

'Verificar que existan filas en el resultado de la consulta
'Se supone que la grilla tiene una fila para los titulos
If Vsgrd.Rows < 2 Then
    MsgBox "La consulta está vacia."
    Exit Function
End If

'Asignar referencias de objeto a las variables
Set xlApp = New Excel.Application
Set xlBook = xlApp.Workbooks.Add
Set xlSheet = xlBook.Worksheets.Add
'Recorrer la grilla e ingresar los valores a la planilla Excel


For intFila = 0 To Vsgrd.Rows - 1
    For intCol = 0 To Vsgrd.Cols - 1
        xlSheet.Cells(intFila, intCol).Value = Vsgrd.TextMatrix(intFila, intCol)
    Next intCol
Next intFila


'Salvar
xlSheet.SaveAs strArchivo
'Cerrar la hoja de cálculo
xlBook.Close
'Cerrar Excel
xlApp.Quit
'Liberar los objetos

Set xlApp = Nothing
Set xlBook = Nothing
Set xlSheet = Nothing
generarExcel = True
Exit Function

err_generarExcel:
If (Not xlBook Is Nothing) Then
'Cerrar la hoja de cálculo
    xlBook.Close False
End If

If (Not xlApp Is Nothing) Then
'   Cerrar Excel
    xlApp.Quit
End If
'Liberar los objetos

Set xlApp = Nothing
Set xlBook = Nothing
Set xlSheet = Nothing

Select Case Err.Number
    Case 1004 'El archivo ya existe y no quiere sobreescribirlo
    Case Else 'Cualquier otro error'
    MsgBox Err.Number & " - " & Err.Description
End Select

End Function

Public Sub settabla(ByVal vtabla As String, ByVal vcampos As String, ByVal vvalores As String)
On Error Resume Next
Dim vsql As String

vsql = "insert into " + vtabla + "(" + vcampos + ") values (" + vvalores + ")"

Call EjecutarScript(vsql)

If Err Then Exit Sub
End Sub


Public Function esProveedor(vcodigo As String) As Boolean
Dim valor, vsql As String
esProveedor = True

vsql = "select * from proveedores  where codigo = '" + vcodigo + "' and not tipoproveedor = 'Proveedor' and not tipoproveedor is null"

valor = traerDatos2(vsql, "codigo", pathDBMySQL)

If valor = "" Then
    esProveedor = False
End If

End Function


Public Function updateNrocheque(vid As String, vnro As Long)
On Error Resume Next
Dim vsql As String

vsql = "update bancos set nrocheque = " + Str(vnro) + " where idBancos='" + vid + "'"

Call EjecutarScript(vsql, pathDBMySQL)

If Err Then Exit Function
End Function

Public Function getDataSource(vsql As String) As ADODB.Recordset
Dim r As New ADODB.Recordset
  
Call r.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

Set getDataSource = r
End Function


Public Function getSaldoDisponible() As Double
On Error Resume Next
Dim vsa, vsp, vs, vd, vh, vss As Double
Dim vidBancos As String
Dim vsql, vcampos, vvalores, vnombre  As String

Dim rs As New ADODB.Recordset

vsql = "select * from bancos where tipodisponibilidad = 'Disponible' and not EsCaja = 'B' order by idBancos"

Call rs.Open(vsql, ConnDDBB, adOpenStatic, adLockPessimistic)

vss = 0

Do Until rs.EOF
    
    vidBancos = rs.Fields("idbancos")
    'vnombre = rs.Fields("descripcion")
    
    'vsa = sacierreTemp2(vidBancos, Date)
        
   ' Call spcierreTemp2(vidBancos, Date - 1, Date, vd, vh, Date - 1)
        
    vss = vss + spcierreTemp2(vidBancos, Date - 1, Date, vd, vh, Date - 1)

    rs.MoveNext
Loop
    
    getSaldoDisponible = vss

If Err Then
    getSaldoDisponible = 0
End If
End Function

Public Sub Log22(v As String)

On Error Resume Next
    v = Str(Date) + " " + Str(Time) + " -> " + v
    Dim v1, v2 As Variant
    Set v1 = CreateObject("Scripting.FileSystemObject")
    'set v2 = v1.CreateTextFile(App.Path + "\Log\Log.txt")
    Set v2 = v1.OpenTextFile(App.Path + "\Log\Log.txt", ForAppending, TristateFalse)
    
    v2.WriteLine v
    v2.Close
If Err Then
    Set v2 = v1.CreateTextFile(App.Path + "\Log\Log.txt")
    v2.WriteLine v
    v2.Close
End If

End Sub

Public Sub log(v As String)
On Error Resume Next
    v = Str(Date) + " " + Str(Time) + " -> " + v
    Dim v1, v2 As Variant
    Set v1 = CreateObject("Scripting.FileSystemObject")
    'set v2 = v1.CreateTextFile(App.Path + "\Log\Log.txt")
    Set v2 = v1.OpenTextFile(App.Path + "\Log\Log.txt", ForAppending, TristateFalse)
    
    v2.WriteLine v
    v2.Close
If Err Then
    Set v2 = v1.CreateTextFile(App.Path + "\Log\Log.txt")
    v2.WriteLine v
    v2.Close
End If

End Sub


Function spcierreTemp2grilla(ByVal vfh As Date) As String
' saldo anterior al cierre
On Error Resume Next

Dim vsql As String

vsql = " select " + _
    " bancosmovimientos.idbancos, " + _
    " sum(bancosmovimientos.debito - bancosmovimientos.credito) As sp " + _
    " From bancosmovimientos " + _
    " Where not idbancos = '098'" + _
    " Group By  bancosmovimientos.idBancos "

spcierreTemp2grilla = vsql

'  " Where not idbancos = '098' and bancosmovimientos.fecha <= '" + strfechaMySQL(vfh) + "' " + _

If Err Then
    Exit Function
End If

End Function


Function spcierreTemp2(ByVal vidBancos As String, ByVal vfd As Date, ByVal vfh As Date, ByRef vd, ByRef vh, ByRef vsp) As Double
' saldo anterior al cierre
On Error Resume Next

Dim vsql As String

vsql = " select " + _
    " sum(bancosmovimientos.debito) as d, sum(bancosmovimientos.credito) as h, " + _
    " sum(bancosmovimientos.debito - bancosmovimientos.credito) As sp " + _
    " From bancosmovimientos " + _
    " Where bancosmovimientos.fecha <= '" + strfechaMySQL(vfh) + "' " + _
    " and bancosmovimientos.idbancos = '" + vidBancos + "'" + _
    " Group By  bancosmovimientos.idBancos "

spcierreTemp2 = Val(traerDatos2(vsql, "sp", pathDBMySQL))


If Err Then
    Exit Function
End If

End Function


Public Sub log2(v As String)

Open App.Path + "\log2.txt" For Append As #1

'While Not EOF(1)

Print #1, v

'Wend

Close #1

End Sub


Public Sub SortByColumn(ByVal sort_column As Integer, MSFlexGrid1 As MSHFlexGrid)
    ' Hide the FlexGrid.
    MSFlexGrid1.Visible = False
    MSFlexGrid1.Refresh

    ' Sort using the clicked column.
    MSFlexGrid1.Col = sort_column
    MSFlexGrid1.ColSel = sort_column
    MSFlexGrid1.Row = 0
    MSFlexGrid1.RowSel = 0

    ' If this is a new sort column, sort ascending.
    ' Otherwise switch which sort order we use.
    If m_SortColumn <> sort_column Then
        m_SortOrder = flexSortGenericAscending
    ElseIf m_SortOrder = flexSortGenericAscending Then
        m_SortOrder = flexSortGenericDescending
    Else
        m_SortOrder = flexSortGenericAscending
    End If
    MSFlexGrid1.Sort = m_SortOrder

    ' Restore the previous sort column's name.
    If m_SortColumn >= 0 Then
        MSFlexGrid1.TextMatrix(0, m_SortColumn) = Mid$(MSFlexGrid1.TextMatrix(0, m_SortColumn), 3)
    End If

    ' Display the new sort column's name.
    m_SortColumn = sort_column
    If m_SortOrder = flexSortGenericAscending Then
        MSFlexGrid1.TextMatrix(0, m_SortColumn) = "> " & MSFlexGrid1.TextMatrix(0, m_SortColumn)
    Else
        MSFlexGrid1.TextMatrix(0, m_SortColumn) = "< " & MSFlexGrid1.TextMatrix(0, m_SortColumn)
    End If

    ' Display the FlexGrid.
    MSFlexGrid1.Visible = True
End Sub


Public Function feCodigoBarra(cuit, ccomprobante, sucursal, cae, fechacae) As String

Dim v, r As String

r = ""

v = Trim(Replace$(cuit, "-", ""))
r = r + v

v = Trim(Format(ccomprobante, "00"))
r = r + v

v = Trim(sucursal)
r = r + v

v = Trim(cae)

r = r + v

v = Trim(fechacae)
r = r + v


r = r + Trim(verificacionCodigoBarra(r))


feCodigoBarra = r

End Function
