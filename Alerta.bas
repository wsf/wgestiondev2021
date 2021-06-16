Attribute VB_Name = "Module1"
Public Sub alerta(vmensaje As String, vtime As Long)

If Not valertaModulo = "" Then
    valertaModulo = valertaModulo + Chr(13) + "-> " + vmensaje
    vmensaje = valertaModulo
End If

    Dim AlertBox As frmAlert
    Set AlertBox = New frmAlert
    
    
        AlertBox.DisplayAlert vmensaje, vtime
        

End Sub

