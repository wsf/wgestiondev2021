Attribute VB_Name = "Informes2Mails"
' Generar informes para mails

Public Sub make_fiels(vsql As String, vPathDB As String)

Dim r  As New ADODB.Recordset

Call r.Open(vsql, vPathDB, adOpenDynamic, adLockReadOnly)
Dim linea As String

Call a1

Do Until r.EOF
    
    For i = 0 To r.Fields.Count
        linea = ""
        linea = linea + r.Fields(i) + " - "
        Call a2(linea)
    Next
    
    
    r.MoveNext
Loop

Call a3
End Sub


Public Sub a1()
On Error Resume Next
    Open App.Path + "\" + "Cuerpo.txt" For Output As 1
If Err Then Exit Sub
End Sub

Public Sub a3()
   Close 1
End Sub

Public Sub a2(vl As String)
On Error Resume Next
   ' Write #1, vl;
    Print #1, vl;
    'Call log.AddItem vl
If Err Then Exit Sub
End Sub


