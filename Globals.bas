Attribute VB_Name = "Globals"
Option Explicit

Public vqrnombre As String
Public vsaldo112, lblsaldocliente2 As Double

Public vmensajeGlobal As String

Public AlertCount As Integer

Private Const PI    As Double = 3.14159265358979
Private Const RADS  As Double = PI / 180    '<Degrees> * RADS = radians
Private Type PointAPI   'API Point structure
    x   As Long
    y   As Long
End Type


Public Type TDATransacciones
    idBancos As Integer
    idcuentas As Integer
    debe As Double
    haber As Double
    debito As Double
    credito As Double
End Type




Private Declare Function Polygon _
                Lib "gdi32" (ByVal hdc As Long, _
                             lpPoint As PointAPI, _
                             ByVal nCount As Long) As Long

Public Sub garbage()
On Error Resume Next
' verifica si hay que hacer rollback porque hubo una transacción que no se terminó
Dim vsql1 As String
Dim vnrointerno1 As Long

vsql1 = "select max(nrointerno) as c from t_rollback"
vnrointerno1 = Val(traerDatos2(vsql1, "c", pathDBMySQL))

If vnrointerno1 = 0 Then Exit Sub

vsql1 = "delete from t_rollback where nrointerno=" + Str(vnrointerno1)
Call EjecutarScript(vsql1, pathDBMySQL)

MsgBox "Se borró la transacción por no guardarse correctamente: " + Str(vnrointerno1), vbCritical

Call GrabarLog("Rollback", Str(vnrointerno1), "Borrar")
If Err Then Exit Sub
End Sub
Public Sub DrawAngle(picDraw As PictureBox, _
                     ByVal fAngle As Single)

    Dim iSize       As Integer
    Dim iFillStyle  As Integer
    Dim lFillColor  As Long
    Dim lForeColor  As Long
    Dim lRet        As Long
    Dim uaPts(3)    As PointAPI

    'Size arrow to best fit picDraw at any angle
    iSize = IIf(picDraw.ScaleHeight < picDraw.ScaleWidth, Int(picDraw.ScaleHeight / PI), Int(picDraw.ScaleWidth / PI))
    
    'Setup the 4 points of the arrow using the first point
    'as the center and the other points offset from the center.
    uaPts(0).x = picDraw.ScaleWidth / 2
    uaPts(0).y = picDraw.ScaleHeight / 2
    uaPts(1).x = uaPts(0).x - iSize
    uaPts(1).y = uaPts(0).y - iSize
    uaPts(2).x = uaPts(0).x + iSize
    uaPts(2).y = uaPts(0).y
    uaPts(3).x = uaPts(0).x - iSize
    uaPts(3).y = uaPts(0).y + iSize
    
    'Rotate the arrow to the correct angle
    Call RotatePoints(uaPts(0), uaPts, fAngle)
    
    'Save picDraw settings
    iFillStyle = picDraw.FillStyle
    lFillColor = picDraw.FillColor
    lForeColor = picDraw.ForeColor
    
    'Setup picDraw to fill the arrow
    picDraw.FillStyle = vbFSSolid   'Solid Fill
    picDraw.FillColor = &HFFFFFF    'Inside = White
    picDraw.ForeColor = &H0&        'Border = Black
    
    'Draw the filled arrow
    lRet = Polygon(picDraw.hdc, uaPts(0), 4)
    
    'Restore picDraw settings
    picDraw.FillStyle = iFillStyle
    picDraw.FillColor = lFillColor
    picDraw.ForeColor = lForeColor

    'Free the memory
    Erase uaPts
    
End Sub
Private Sub RotatePoints(uAxisPt As PointAPI, _
                         uRotatePts() As PointAPI, _
                         fDegrees As Single)

    'Rotates an array of PointAPI points around a center point by fDegrees

    Dim lIdx        As Long
    Dim fDX         As Single
    Dim fDY         As Single
    Dim fRadians    As Single

    fRadians = fDegrees * RADS
    
    For lIdx = 0 To UBound(uRotatePts)
        fDX = uRotatePts(lIdx).x - uAxisPt.x
        fDY = uRotatePts(lIdx).y - uAxisPt.y
        uRotatePts(lIdx).x = uAxisPt.x + ((fDX * Cos(fRadians)) + (fDY * Sin(fRadians)))
        uRotatePts(lIdx).y = uAxisPt.y + -((fDX * Sin(fRadians)) - (fDY * Cos(fRadians)))
    Next lIdx
    
End Sub
'----------------------------------Destruir---Huerfanos-------------------------------------------------------------------------
Public Sub BuscarHuerfanas(vremito)
On Error Resume Next
    
    Dim connFDetalleH As New ADODB.Connection
    Dim rsFDetalleH As New ADODB.Recordset
    Dim sqlFDetalleH As String

    With connFDetalleH
        .ConnectionString = pathDBMySQL
        .Open
    End With

    sqlFDetalleH = "SELECT * FROM fdetalle ORDER BY remito"
    
    With rsFDetalleH
        Call .Open(sqlFDetalleH, connFDetalleH, adOpenStatic, adLockReadOnly)
        
        Do Until .EOF = True
             If (BuscarRelacion("Factura", "(remito = " & .Fields("Remito").Value & ")") = False) Or (BuscarRelacion("CuentaCorriente", "(remito = " & .Fields("Remito").Value & ")") = False) Then
                BorrarBase "FDetalle WHERE remito = " & .Fields("Remito").Value & "", pathDBMySQL
            End If
            .MoveNext
        Loop
    End With
    
    sqlFDetalleH = ""
    
    rsFDetalleH.Close
    Set rsFDetalleH = Nothing
    
    connFDetalleH.Close
    Set connFDetalleH = Nothing
    
If Err Then GrabarLog "BuscarHuerfanas", Err.Number & " " & Err.Description, "Global"
End Sub
Private Function BuscarRelacion(vtabla As String, vsql As String) As Boolean
On Error Resume Next

    Dim connRelacion As New ADODB.Connection
    Dim rsRelacion As New ADODB.Recordset
    Dim sqlRelacion As String

    With connRelacion
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlRelacion = "SELECT * FROM " & vtabla & " WHERE " & vsql & ""

    With rsRelacion
        Call .Open(sqlRelacion, connRelacion, adOpenStatic, adLockReadOnly)
            
        BuscarRelacion = Not .EOF
    
    End With
    
    sqlRelacion = ""
    
    rsRelacion.Close
    Set rsRelacion = Nothing
    
    connRelacion.Close
    Set connRelacion = Nothing

If Err Then GrabarLog "BuscarRelacion", Err.Number & " " & Err.Description, "Global"
End Function
'------------------------------------------------------------------------------------------------------------


Public Sub controlarFacturaDetalles()
Dim sql As String

sql = "SELECT NComprobante, Fecha, Nombre, Totales, subtotal" & _
" From  view01 t1  INNER JOIN factura t2 ON (t1.remito=t2.remito)" & _
" Where " & _
" not(t1.`Totales` = t2.`subTotal`)  order by fecha desc"
 
frmConsultas.Buscar (sql)
frmConsultas.Show

End Sub




Public Sub fMostrarGrilla(vsql As String)

'frmConsultas.vcampoMostrar = vcampoMostrar
'frmConsultas.vcampoID = vcampoID
'frmConsultas.vcontrol = vcontrol
frmConsultas.vgsql = vsql

'Set frmConsultas.vform = vform

frmConsultas.Show

End Sub


Public Sub fbuscarGrilla(ByVal vtabla As String, ByVal vcampoMostrar As String, ByVal vcampoID As String, ByVal vcontrol As String, vform As Form, Optional vcampo2 As String, Optional enComuna As Boolean)

frmConsultas.vcampoMostrar = vcampoMostrar
frmConsultas.vcampoID = vcampoID
frmConsultas.vcontrol = vcontrol
frmConsultas.vtabla = vtabla
frmConsultas.vcampo2 = vcampo2
frmConsultas.enComuna = enComuna

Set frmConsultas.vform = vform
Call frmConsultas.vbuscando_Change

frmConsultas.Show
frmConsultas.vbuscando.SetFocus
End Sub



'Public Sub doTransaccion(vidConcepto As Integer, vimporte As Double)
'
'Dim vsql, vvalores As String
'Dim vd, vh As Double
'
'vd = 0
'vh = 0
'
'Dim r As Recordset
'
'vsql = "select * from conceptos2 where idConceptos=" + Str(vidConcepto)
'
''Call getRegistro(r, vsql)
'
'
'If r.Fields("debito") Then vd = vimporte
'
'If r.Fields("credito") Then vh = vimporte
'
'vvalores = Str(r.Fields("idbancos")) + "," + Str(vd) + "," + srt(vh)
'vsql = "insert into bancosmovimientos (idbancos,debito,credito) values (" + vvalores + ")"
'
'Call EjecutarScript(vsql, pathDBMySQL)
'
'
''--- cargar asiento automaticamente
'
'If r.Fields("debe") Then vd = vimporte
'If r.Fields("haber") Then vh = vimporte
'
'Call doAsiento(r.Fields("idCuentas"), vd, vh)
'
'
'End Sub


Private Sub doAsiento(vfecha As Date, vcta As String, vd As Double, vh As Double)


' --- variables
Dim vsql As String
Dim vcampos, vvalores As String
Dim vnrointerno As Long
Dim vnroasiento, vnrobalance As Integer
Dim vcodc, vcodp As Double


vnrointerno = UltimoNroInterno2

vnroasiento = Val(GenerarDato("SELECT MAX(Numero) as NroAsiento FROM Asientos where balance=" + Str(vnrobalance), "NroAsiento")) + 1 ' paso 2 para asiento

vnrobalance = TraerDato("balances", " Activo='S' order by NroBalance Desc", "NroBalance", pathDBMySQL)

vcampos = "fecha,numero,leyenda,tipoMovimiento,nrobalance,nrointerno,codigoproveedor,codigocliente,marca"

vvalores = strfecha2(vfecha) + "," + Str(vnroasiento) + ",''," + Str(vnrobalance) + "," + Str(vnrointerno) + "," + Str(vcodc) + "," + Str(vcodp) + ",'Normal'"


vsql = "insert into asientos (" + vcampos + ") values (" + vvalores + ")"
Call EjecutarScript(vsql, pathDBMySQL)





End Sub


Private Sub getRegistro(ByRef vr As Recordset, vsql As String)
On Error Resume Next

 '   Call vr.Open(vsql, connFDetalleH, adOpenStatic, adLockReadOnly)

If Err Then Exit Sub
End Sub



Public Sub vaciarControl(ByRef c As control)
On Error Resume Next

    c.Text = ""

If Err Then Exit Sub
End Sub


Public Function getNroCheque(vid As String) As Long
On Error Resume Next

Dim rs As New ADODB.Recordset
Dim vsql, vr  As String

vsql = "select * from bancos where bancos.EsCaja='N' and bancos.idBancos='" + vid + "'"


Call rs.Open(vsql, ConnDDBB, adOpenStatic, adLockReadOnly)

If rs.RecordCount > 0 Then
    getNroCheque = rs.Fields("nrocheque")
Else
    getNroCheque = -1
End If

If Err Then Exit Function
End Function


Function IsFormLoaded(FormToCheck As Form) As Integer
Dim y As Integer

For y = 0 To Forms.Count - 1
If Forms(y) Is FormToCheck Then
    IsFormLoaded = True
    Exit Function
    End If
Next
IsFormLoaded = False
End Function


Public Sub prenderCartel()
    frmCartel.Show
End Sub

Public Sub apagarCartel()
   Unload frmCartel
End Sub

Public Function NroRemitoNuevo() As Long
    On Error Resume Next

    Dim vnro As Long
    Dim vsql As String
    
    vsql = "select max(numero) as c from t_nroremito "
    
    
    ' vnro = TraerDato2("select * from t_nroremito order by numero desc", "numero", pathDBMySQL)
    
    vnro = TraerDato2(vsql, "c", pathDBMySQL)
    
    vsql = "insert into t_nroremito (numero) values (" + Str(vnro + 1) + ")"
    
    Call EjecutarScript(vsql, pathDBMySQL)
    
    NroRemitoNuevo = vnro + 1
    
    If Err < 0 Then
            MsgBox "Cuidado!", vbCritical
            NroRemitoNuevo = 0
            Exit Function
    End If
End Function



Public Function escodigodebarra(vdatos As String) As Boolean
escodigodebarra = False

If Val(vdatos) > 0 And Len(vdatos) > 6 Then
    escodigodebarra = True
End If

End Function

