VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmUltimaCompra 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listados de última compra"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   5760
   Begin VB.ListBox l 
      Height          =   1035
      Left            =   360
      TabIndex        =   15
      Top             =   2880
      Width           =   5085
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   285
      Left            =   90
      TabIndex        =   12
      Top             =   5010
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ComboBox cboreparto 
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   1965
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   150
      TabIndex        =   7
      Top             =   1230
      Width           =   5295
      Begin VB.OptionButton opno 
         Caption         =   "Clientes S/ Compra"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   275
         Width           =   1695
      End
      Begin VB.OptionButton opsi 
         Caption         =   "Clientes C/  Compra"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   275
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   5295
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   72286209
         CurrentDate     =   38023
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   285
         Left            =   3600
         TabIndex        =   4
         Top             =   210
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   72286209
         CurrentDate     =   38023
      End
      Begin VB.Label Label4 
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2790
         TabIndex        =   6
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label3 
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   5
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.CheckBox chkFechas 
      Caption         =   "Anular fechas"
      Height          =   195
      Left            =   4080
      TabIndex        =   1
      Top             =   960
      Value           =   1  'Checked
      Width           =   1305
   End
   Begin VB.CommandButton cmdVerListado 
      Caption         =   "Ver listado"
      Height          =   375
      Left            =   3240
      TabIndex        =   0
      Top             =   2370
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc bclireparto 
      Height          =   330
      Left            =   120
      Top             =   6480
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "bclireparto"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc bnocompra_temp 
      Height          =   330
      Left            =   120
      Top             =   7200
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "bnocompra_temp"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc bfactura 
      Height          =   330
      Left            =   120
      Top             =   7560
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "bfactura"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc bcliente 
      Height          =   330
      Left            =   120
      Top             =   6840
      Visible         =   0   'False
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "bcliente"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label pasan 
      Caption         =   "Cantidad de Registros:"
      Height          =   285
      Left            =   450
      TabIndex        =   14
      Top             =   4590
      Width           =   4815
   End
   Begin VB.Label vregistros 
      Caption         =   "Cantidad de Registros:"
      Height          =   285
      Left            =   480
      TabIndex        =   13
      Top             =   4260
      Width           =   4815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "> Reparto :"
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
      Left            =   150
      TabIndex        =   10
      Top             =   2010
      Width           =   975
   End
End
Attribute VB_Name = "frmUltimaCompra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboreparto_GotFocus()
On Error Resume Next

    Call CargarCombo("clireparto", "Decrip", cboReparto, False)
    
If Err Then GrabarLog "cboreparto_GotFocus", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CerrarBase(vConexion As Integer)
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
            
    MsgBox "   Prepare la Impresora   ", vbInformation, "Mensaje ..."
            
    Select Case vConexion
    
        Case 0

            With Mantenimiento.rsnocompra_temp
                If .State = 1 Then
                    .Close
                    .Open
                Else
                    .Open
                    .Close
                    .Open
                End If
        
            End With
    
        Case 1
        
            With Mantenimiento.rsucompra
                If .State = 1 Then
                    .Close
                    .Open
                Else
                    .Open
                    .Close
                    .Open
                End If
            End With
    End Select
        
If Err Then GrabarLog "CerrarBase", Err.Number & " " & Err.Description, Me.Name
End Sub
    
Private Sub cmdVerListado_Click()
On Error Resume Next

    Dim sql, filter, filter2 As String
    Dim i As Integer
    
    i = 0
    
    sql = ""
    filter = ""
    filter2 = ""
    
    
    If Not Trim(cboReparto.Text) = "" Then
        sql = sql + " and reparto = '" + Trim((cboReparto)) + "'"
        filter = filter + " and últimodereparto = '" + Trim((cboReparto)) + "'"
        filter2 = filter2 + " and reparto = '" + Trim((cboReparto)) + "'"
    End If
        
    
    If opno.Value = True Then 'Sin compras
       
        BorrarBase "nocompra_temp", pathDBMySQL
        
        With bnocompra_temp
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM nocompra_temp"
            .Refresh
        End With
      
        '  -------------------------
        l.Clear
        With bcliente
            If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
            .RecordSource = "SELECT * FROM clientes WHERE (reparto = '" + Trim((cboReparto)) + "') AND (pasivo = 'NO') ORDER BY nombre"
            .Refresh
      
            Barra.Max = .Recordset.RecordCount
            Barra.Value = 0
        
        End With
        
        Do Until bcliente.Recordset.EOF = True
      
            l.AddItem ("> " & bcliente.Recordset("nombre").Value)
            '--- por cada cliente, veo si hay movimiento en el rango de fecha
        
            With bfactura
                If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
                .RecordSource = "select * from factura where fecha >= '" & strfechaMySQL(fdesde.Value) + "' and fecha <= '" & strfechaMySQL(fhasta.Value) + "' and codigo = '" + Trim(bcliente.Recordset("codigo")) + "'"
                .Refresh
            End With
        
            If bfactura.Recordset.EOF Then ' en el caso que no encuentra la factura debe guardar el nombre del cliente
                
                l.AddItem (">>>>>>>>>>>>> " & bcliente.Recordset("nombre").Value)
                i = i + 1
                With bnocompra_temp
                    .Recordset.AddNew
                    
                    .Recordset("codigo").Value = bcliente.Recordset("codigo").Value
                    .Recordset("nombre").Value = bcliente.Recordset("nombre").Value
                    .Recordset("Saldo").Value = 0
                    .Recordset("reparto").Value = bcliente.Recordset("reparto").Value
                    
                    .Recordset.Update
                End With
            End If
            
            bcliente.Recordset.MoveNext
            Barra.Value = Barra.Value + 1
        
        Loop
        
        bnocompra_temp.Refresh
        
        vregistros.Caption = "Cantidad de Registros: " & bnocompra_temp.Recordset.RecordCount
        pasan.Caption = "Pasaron: " & i
        

        CerrarBase 0
        With drNoCompra
            .Sections("TituloEmpresa").Controls("vfechas").Caption = "Desde : " & fdesde.Value & " hasta: " & fhasta.Value
            .Show
        End With
        
        
    Else 'Con Compras

        If Not chkFechas.Value = 1 Then
            sql = sql + " and UFecha >= '" + Str(fdesde) + "' and UFecha <= '" + Str(fhasta) + "'"
            filter = filter + " and UFecha >= '" + strfecha2(fdesde) + "' and UFecha <= '" + strfecha2(fhasta) + "'"
        End If
        
        CerrarBase 1
        
        With Mantenimiento.rsucompra
            .filter = "Codigo > 0" + filter
            .Sort = "ÚltimoDeNombre ASC"
        End With
        
        With drucompra
            .Sections("TituloEmpresa").Controls("vreparto").Caption = Me.cboReparto.Text
            
            .Show
        End With
    
    End If

    If Err Then
        MsgBox "Faltan ingresar datos.", vbCritical, "Error ..."
        GrabarLog "cmdVerListado_Click", Err.Number & " " & Err.Description, Me.Name
    End If
End Sub

Private Sub chkFechas_Click()
On Error Resume Next

    fdesde.Enabled = CBool(chkFechas.Value - 1)
    fhasta.Enabled = CBool(chkFechas.Value - 1)

If Err Then GrabarLog "chkFechas_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Form_Load()
On Error Resume Next
    
    Width = 5850
    Height = 5895

    fdesde.Value = Date
    fhasta.Value = Date

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Function freparto(vreparto As String) As String

    With bclireparto
        If .ConnectionString = "" Then .ConnectionString = pathDBMySQL
        .RecordSource = "SELECT * FROM clireparto"
        .Refresh
        .Recordset.Find ("descrip = '" + Trim(vreparto) + "'")

        If Not .Recordset.EOF = True Then
            freparto = .Recordset("nreparto")
        End If

    End With

End Function

