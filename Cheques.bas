Attribute VB_Name = "Cheques"
Public Function fdsChequesToString()
Dim v As String

With gbldsCheques

v = " > Nro:" + .Ncheque & _
" > F.Acreditación:" + Str(.FechaAcreditacion) & _
" > Banco:" + EsNulo(traerDatos2("select * from bancos where idBancos='" + .idBancos + "'", "Descripcion", pathDBMySQL)) & _
" > Sucursal:" + EsNulo(traerDatos2("select * from bancoscuentas where idBancos='" + Str(.idBancosCuentas) + "'", "Descripcion", pathDBMySQL))

End With


fdsChequesToString = v

End Function


Public Sub chequesEnCartera(vcpInstancia As String, vcod As String, vnombre As String)

If vcpInstancia = "cobro" Then Exit Sub ' si es un cobro no puede seleccionar cheques

frmCheques.ComListado.Text = "En Cartera"
frmCheques.CombOrdenamiento.Text = "marcainterna"

Call frmCheques.PBFiltrar_Click
frmCheques.Show
frmCheques.WindowState = vmaximizar
frmCheques.vViene = vcpInstancia

gbldsCheques.Codigo = vcod ' Me.txtCliente(0).Text
gbldsCheques.Nombre = vnombre 'Me.txtCliente(1).Text

frmCheques.vbusca.SetFocus

'frmcheques.seleccionar(cpinstancia,"cartera")
End Sub


Public Sub cambiarCajaCheque(vid As Long, vcodCaja As String)
On Error Resume Next

Dim vsql As String


vsql = "update cheques set idcustodia ='" + Trim(vcodCaja) + "' where idcheques=" + Trim(Str(vid))

Call EjecutarScript(vsql, pathDBMySQL)

If Err Then Exit Sub
End Sub



