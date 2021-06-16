Attribute VB_Name = "Json"


Public Function getDic(key As String, ByRef arr2) As String

getDic = ""

For i = 0 To 99
    If arr2(i, 0) = key Then
        getDic = arr2(i, 1)
    End If
Next
        
End Function


Public Sub setDic(k As String, d As String, ByRef arr2)
Dim i As Integer

For i = 0 To 99
    If arr2(i, 0) = "" Then
        arr2(i, 0) = k
        arr2(i, 1) = d
        Exit Sub
    End If
        'End Sub
Next

End Sub

Public Sub initDic(ByRef arr2)
Dim i As Integer

For i = 0 To 99

    arr2(i, 0) = ""
    arr2(i, 1) = ""
    
Next

End Sub


