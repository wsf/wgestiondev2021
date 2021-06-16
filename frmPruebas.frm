VERSION 5.00
Begin VB.Form frmPruebas 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   10560
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmPruebas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sqlArticulos As String
Dim rsArticulos As ADODB.Recordset
Private Sub GrabarArticulo()
On Error Resume Next

    sqlArticulos = "SELECT * FROM Articulos WHERE (idArticulos = 25)"
        
    Set rsArticulos = New ADODB.Recordset
    
    With rsArticulos
        .CursorLocation = adUseServer
        Call .Open(sqlArticulos, ConnDDBB, adOpenDynamic, adLockOptimistic)

            .Fields("Codigo").Value = EsNulo(.Fields("Codigo").Value)
            .Fields("Descrip").Value = EsNulo(.Fields("Descrip").Value)
            .Fields("CodigoNum").Value = Val(.Fields("Codigo").Value)
            .Fields("CodigoBarra").Value = " "
            .Fields("idSubRubros").Value = EsNulo("001")
            .Fields("idRubros").Value = EsNulo("001")
            'Call GuardarFoto(rsArticulos, phtArticulo.PhotoFileName)
        
            'Ficha
            .Fields("idPorcentajeIva").Value = EsNulo(2)
            .Fields("idProveedor").Value = EsNulo("001")
            .Fields("idFabricantes").Value = EsNulo("001")
            .Fields("PCosto").Value = Val(10)
            
            Dim i As Integer
            For i = 1 To 6
                .Fields("PVenta" & i).Value = Val(Format(60, "#####0.00"))
            Next
            
            'Tecnica
            .Fields("FechaAlta").Value = strfechaMySQL(Date)
            
            'If chkActualizacionDePrecio.Value = xtpUnchecked Then
            '    If Val(vCostoAnterior) <> Val(txtFicha(7).Text) Then
            '        .Fields("FechaModificacion").Value = strfechaMySQL(Date)
            '    Else
            '        If Val(vPVentaAnterior) <> Val(txtFicha(8).Text) Then
            '            .Fields("FechaModificacion").Value = strfechaMySQL(Date)
            '        End If
            '    End If
            'Else
                .Fields("FechaModificacion").Value = strfechaMySQL(Date)
            'End If
            
            .Fields("Peso_U").Value = Val(20)
            .Fields("Peso_T").Value = Val(30)
            .Fields("UnidadesPorBulto").Value = Val(10)
            .Fields("Dimensiones").Value = EsNulo("2x5")
            .Fields("MensajeEmergente").Value = EsNulo("ada")
            .Fields("CodigoConcepto").Value = Val(1)
            .Fields("Observaciones").Value = EsNulo("Muy BUeno")

            .Fields("Porcentaje").Value = 50
            .Fields("idMoneda").Value = "001"
            
            'Stock
            .Fields("Stock").Value = Val(10)
            .Fields("StockMin").Value = Val(1)
            .Fields("StockMax").Value = Val(100)
            .Fields("idDepositos").Value = EsNulo("001")
            
            '.Fields("TimeStamp").Value = Now()
            .Update
        
        
            If Err.Description = "" Then
            
            Else
                MsgBox Err.Description
            End If
            
            
        
    End With

    sqlArticulos = ""
    
    If rsArticulos.State = 1 Then
        rsArticulos.Close
        Set rsArticulos = Nothing
    End If

If Err Then
    
    GrabarLog "GrabarArticulo", Err.Number & " " & Err.Description, Me.Caption
    
End If
End Sub

Private Sub Form_Load()
On Error Resume Next

    Me.Show

    GrabarArticulo

    Unload Me
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
