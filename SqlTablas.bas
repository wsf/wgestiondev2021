Attribute VB_Name = "SqlTablas"
Public Sub InsertarEnTabla(vtabla As String, ByVal vcampos As String, ByVal vvalores As String, Optional vPathDB)
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
    
    
    vsql = "INSERT INTO " + Trim(vtabla) + " (" + Trim(vcampos) + ")" + "VALUE (" + vvalores + ")"
    
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
