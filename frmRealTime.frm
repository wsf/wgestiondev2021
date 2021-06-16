VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmRealTime 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   9780
   ClientLeft      =   10245
   ClientTop       =   540
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9780
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   14420
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Timer tmrActualizador 
      Interval        =   1000
      Left            =   1200
      Top             =   8160
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   8175
      Left            =   4440
      TabIndex        =   1
      Top             =   0
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   14420
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   11274
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRealTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
On Error Resume Next



If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub FDetalle()
On Error Resume Next

    Dim connFdetalle As New ADODB.Connection
    Dim rsFdetalle As New ADODB.Recordset
    Dim sqlFDetalle As String

    With connFdetalle
        If .State = 1 Then .Close
        .ConnectionString = pathDBMySQL
        .Open
    End With

    sqlFDetalle = "SELECT * FROM fdetalle_temp"

    With rsFdetalle
        If .State = 1 Then
            
            .Close
        End If
        .CursorLocation = adUseClient
        Call .Open(sqlFDetalle, connFdetalle, adOpenStatic, adLockReadOnly)
                
        Set DataGrid1.DataSource = rsFdetalle
        
    End With

If Err Then GrabarLog "FDetalle", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Factura()
On Error Resume Next

    Dim connFactura As New ADODB.Connection
    Dim rsFactura As New ADODB.Recordset
    Dim sqlFactura As String

    With connFactura
        If .State = 1 Then .Close
        .ConnectionString = pathDBMySQL
        .Open
    End With

    sqlFactura = "SELECT * FROM Factura_temp"

    With rsFactura
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        Call .Open(sqlFactura, connFactura, adOpenStatic, adLockReadOnly)
                
        Set DataGrid2.DataSource = rsFactura
        
    End With

If Err Then GrabarLog "Factura", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub tmrActualizador_Timer()

    FDetalle
    Factura
End Sub
