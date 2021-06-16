Attribute VB_Name = "Indicadores"
Private Sub main()
On Error Resume Next

Dim rec As New Recordset
' recorrer la tabla t_auditoria


vsql = "select * from t_auditoria"

Call rec.Open(vsql, pathDBMySQL, adOpenDynamic, adLockReadOnly)


Do Until rec.EOF


    rec.MoveNext
Loop



If Err Then Exit Sub
End Sub
