Attribute VB_Name = "Invariantes"
Public Sub invNroRemito()
Dim vsql As String
Dim vnf, vnpf As Long

vsql = " select max(t.Remito) as c from pfactura t "
vsql = " select max(t.Remito) as c from factura t "
vsql = " select max(t.numero) as c from t_nroremito t "

End Sub

Public Sub doEventos()
On Error Resume Next
Dim vsql, vmensaje, vaccion, vsql1 As String

Dim valor As String
Dim rec As New ADODB.Recordset


vsql1 = "select * from eventosadmin where estado='Activo'"


With rec

    Call .Open(vsql1, ConnDDBB, adOpenDynamic, adLockBatchOptimistic)


Do Until .EOF
    vsql1 = .Fields("sql")
    vmensaje = .Fields("mensaje")
    vaccion = .Fields("accion")
    
    valor = TraerDato2(vsql, "c", pathDBMySQL)
    
            If Not valor = "" Then
                frmAlarmas.Show
                frmAlarmas.log.AddItem ">  " + vmensaje + valor
                EbExecuteLine StrPtr(vaccion), 0, 0, 0
            End If
    
    .MoveNext
Loop



End With

If Err Then Exit Sub


End Sub

