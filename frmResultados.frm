VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmResultados 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Resultados"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1635
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEjecutar 
      Caption         =   "Ejecutar"
      Height          =   375
      Left            =   30
      TabIndex        =   5
      Top             =   870
      Width           =   6255
   End
   Begin VB.Frame FraGeneral 
      Height          =   825
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.CheckBox chkSeparadas 
         Caption         =   "Separar rubros en cada página"
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   360
         Width           =   2775
      End
      Begin MSComCtl2.DTPicker dtpCuentas 
         Height          =   315
         Left            =   1590
         TabIndex        =   1
         Top             =   330
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   68354049
         CurrentDate     =   39448
      End
      Begin VB.Label lblDatos 
         Alignment       =   1  'Right Justify
         Caption         =   "> Listar al Día:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   390
         Width           =   1425
      End
   End
   Begin MSComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   1320
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
End
Attribute VB_Name = "frmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
On Error Resume Next

    With Me
        .Left = 0
        .Top = 0
        .Width = 6420
        .Height = 1965
        .KeyPreview = True
    End With

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdEjecutar_Click()
On Error Resume Next
    
    Call BorrarBase("TempCuentas", pathDBMySQL)
    Call Cuentas
    Call MostrarReporte
   
If Err Then GrabarLog "cmdEjecutar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Cuentas()
On Error Resume Next

    Dim rsCuentas As New ADODB.Recordset, sConsultaCuentas As String
    
    sConsultaCuentas = "SELECT * FROM cuentas"
    
    With rsCuentas
        Call .Open(sConsultaCuentas, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            .MoveFirst
            
            Barra.Value = 0
            Barra.Max = .RecordCount
            
            Do Until .EOF = True
                GuardarTemp .Fields("CodigoCuenta").Value, .Fields("Cuenta").Value, .Fields("Indice").Value
                .MoveNext
                Barra.Value = Barra.Value + 1
            Loop
        
        End If
    
    End With
    
    sConsultaCuentas = ""
    
    If rsCuentas.State = 1 Then
        rsCuentas.Close
        Set rsCuentas = Nothing
    End If
    
If Err Then GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub GuardarTemp(vCodCuenta As Long, vCuenta As String, vCodRubro As String)
On Error Resume Next

    Dim rstemp As New ADODB.Recordset, sConsultaTemp As String
        
    sConsultaTemp = "SELECT * FROM TempCuentas"
    
    With rstemp
        Call .Open(sConsultaTemp, ConnDDBB, adOpenDynamic, adLockOptimistic)
        
        .AddNew
        
        .Fields("Codigo").Value = vCodCuenta
        .Fields("Cuenta").Value = vCuenta
        .Fields("Rubro").Value = vCodRubro
        'Uso SaldoAnteriorH como Saldo Anterior
        .Fields("SaldoAnteriorH").Value = Val(Format(CalcularSaldo(vCodCuenta, True), "########0.000"))
        
        'Uso SaldoH como Saldo Actual
        .Fields("SaldoH").Value = Val(Format(CalcularSaldo(vCodCuenta, False), "########0.000"))

        'Uso SaldoD como Saldo Total
        .Fields("SaldoD").Value = .Fields("SaldoAnteriorH").Value + .Fields("SaldoH").Value
            
        .Update
    End With
    
    sConsultaTemp = ""
    
    rstemp.Close
    Set rstemp = Nothing

    
If Err Then GrabarLog "GuardarTemp", Err.Number & " " & Err.Description, Me.Name
End Sub
Function CalcularSaldo(ByRef vCodCuenta As Long, vAnterior As Boolean) As Double
On Error Resume Next

    Dim rsSaldo As New ADODB.Recordset, sConsultaSaldo As String

    
    If vAnterior = True Then
        sConsultaSaldo = "SELECT Asientos.Codigo, Max(Asientos.NCuenta) AS Cuenta, Sum(Asientos.Debe) AS Debe, Sum(Asientos.Haber) AS Haber FROM Asientos WHERE (((Asientos.Fecha) < '" & strfechaMySQL(dtpCuentas.Value) + "')) GROUP BY Asientos.Codigo HAVING (((Asientos.Codigo)= " & vCodCuenta & "))"
    Else
        sConsultaSaldo = "SELECT Asientos.Codigo, Max(Asientos.NCuenta) AS Cuenta, Sum(Asientos.Debe) AS Debe, Sum(Asientos.Haber) AS Haber FROM Asientos WHERE (((Asientos.Fecha) >= '" & strfechaMySQL(dtpCuentas.Value) + "')) GROUP BY Asientos.Codigo HAVING (((Asientos.Codigo)= " & vCodCuenta & "))"
    End If
    
    With rsSaldo
        .CursorLocation = adUseClient
        Call .Open(sConsultaSaldo, ConnDDBB, adOpenStatic, adLockReadOnly)
        
        If Not .EOF = True Then
            CalcularSaldo = Val(Format(.Fields("Debe").Value, "##########0.00")) - Val(Format(.Fields("Haber").Value, "##########0.00"))
        Else
            CalcularSaldo = 0
        End If
    
    End With
    
    sConsultaSaldo = ""
    
    rsSaldo.Close
    Set rsSaldo = Nothing

If Err Then GrabarLog "CalcularSaldo", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub MostrarReporte()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento

    MsgBox "Prepare la Impresora!!!", vbInformation, "Mensaje ..."
    
    With Mantenimiento.rsRubrosResultados
        .Source = "SHAPE {SELECT * FROM rubros} AS RubrosResultados APPEND ({SELECT * FROM TempCuentas}  AS TempResultados RELATE 'Codigo' TO 'Rubro') AS TempResultados"
        
        If Not .State = 1 Then
            .Open
            .Close
            .Open
        Else
            .Close
            .Open
        End If
        
        If .RecordCount = 0 Then
            MsgBox "El sistema no presenta una Tabla con los Rubros Correspondientes", vbExclamation, "Mensaje ..."
            Exit Sub
        End If

    End With
    
    With drResultados
    
        .Sections("Rubros_Header").ForcePageBreak = chkSeparadas.Value
        
'        .Sections(2).Controls("snombre").Caption = gnombre
'        .Sections(2).Controls("sdirtel").Caption = gdireccion & "  /  " & gtelefono
'        .Sections(2).Controls("slocalidad").Caption = glocalidad
'        .Sections(2).Controls("semail").Caption = gemail

        .Show
    End With
    
If Err Then GrabarLog "MostrarReporte", Err.Number & " " & Err.Description, Me.Name
End Sub
