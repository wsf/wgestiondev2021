VERSION 5.00
Object = "{D9AF33E0-7C55-11D5-9151-0000E856BC17}#1.0#0"; "fiscal010724.ocx"
Object = "{AFD24A52-2823-4FBD-B75D-C282C11E1D98}#1.0#0"; "IFEpson.ocx"
Begin VB.Form frmImpresionFiscal 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin EPSON_Impresora_Fiscal.PrinterFiscal FiscalEpson 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin FiscalPrinterLibCtl.HASAR FiscalHasar 
      Left            =   600
      OleObjectBlob   =   "frmImpresionFiscal.frx":0000
      Top             =   0
   End
End
Attribute VB_Name = "frmImpresionFiscal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Me.Show
                With Me.FiscalEpson
                    
                        .BaudRate = 9600
                        .PortNumber = 1
                        .MessagesOn = True
                    
                    
                    'Call .DNFHCreditCard("", "", "", "", "", "", "", "", "", "", "", "")
                                        
                    If .Status = True Then
                        Call .OpenInvoice("T", "C", "A", "1", "P", "12", "Monotributo", "Responsable Inscripto", "Adrian Bortoli", "", "CUIT", "20-29379389-2", "N", "-", "-", "-", "", "", "G")
                        Call .SendInvoiceItem("Una Notebook", "1", "10000", ".21", "M", "0", "0", "", "", "", "0")
                        Call .SendInvoicePayment("Descuento por pagoA", "200.00", "D")
                        Call .SendInvoicePayment("Recargo por x motit", "150", "R")
                        Call .CloseInvoice("T", "A", "Total")
                    End If
            End With

End Sub
