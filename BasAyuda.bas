Attribute VB_Name = "BasAyuda"
Option Explicit
Public Sub VerAyuda(vFormulario As String)
On Error Resume Next


If Err Then GrabarLog "VerAyuda", Err.Number & " " & Err.Description, "BasAyuda"
End Sub
