Attribute VB_Name = "CPEstructuras"
Option Explicit
Public Type CobrosPagos
        FormPersona As Form
        FormBuscaDocumento As Form 'frmbuscafactuar, frmbuscacompra
        TablaPersona As String ' cliente, proveedor
        TablaCtaCte As String 'cuentascorrientes, pcuentascorrientes, etc
End Type

Public CobrosPagos As CobrosPagos
