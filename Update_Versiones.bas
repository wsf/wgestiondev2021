Attribute VB_Name = "Update_Versiones"
Public Sub actualizar()
Dim v As Double
Exit Sub
If Not Dir(App.Path + "\Actualizaciones\" + "script") = "" Then
       
    If MsgBox("Hay una nueva actualización para este sistema. " + Chr(13) + "Desea instalarla ? ", vbYesNo) = vbYes Then
            
            'a = Shell("cd " + App.Path + "\Actualizaciones", 1)
            
            a = Shell(App.Path + "\Actualizaciones\" + "scriptBase.bat /s", vbMaximizedFocus)
            
           ' a = Shell("del " + App.Path + "\Actualizaciones\" + "script /s", 1)
            
            MsgBox "Las actualizaciones fueron instalada adecuadamente", vbInformation
    End If



End If

End Sub
 
    

