VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmInicio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Realizando tareas de inicio ..."
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab SSTab1 
      Height          =   7785
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   13732
      _Version        =   327681
      TabHeight       =   520
      TabCaption(0)   =   "Avisos"
      TabPicture(0)   =   "Inicio.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dgErrores"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdArreglarVista"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lstError"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "c3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "c2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "c1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "Moras"
      TabPicture(1)   =   "Inicio.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command5"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Command4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Command3"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "MSHFlexGrid1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "Conf."
      TabPicture(2)   =   "Inicio.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command5 
         Caption         =   "Aplicar nuevo recargo"
         Height          =   435
         Left            =   -73500
         TabIndex        =   22
         Top             =   6660
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Porcentaje de recargo: "
         Height          =   705
         Left            =   -74850
         TabIndex        =   20
         Top             =   900
         Width           =   4005
         Begin VB.TextBox vrecargo 
            Height          =   315
            Left            =   2160
            TabIndex        =   21
            Top             =   240
            Width           =   1725
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Confirmar"
         Height          =   435
         Left            =   -74850
         TabIndex        =   18
         Top             =   7260
         Width           =   4065
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Sacar recargos "
         Height          =   435
         Left            =   -74880
         TabIndex        =   17
         Top             =   6660
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Height          =   525
         Left            =   150
         TabIndex        =   12
         Top             =   1440
         Width           =   4035
         Begin VB.CommandButton Command2 
            Caption         =   "ver ..."
            Height          =   285
            Left            =   3360
            TabIndex        =   14
            Top             =   180
            Width           =   585
         End
         Begin VB.CheckBox Check1 
            Appearance      =   0  'Flat
            Caption         =   "Recargos por deudas atrasadas:"
            Enabled         =   0   'False
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   90
            TabIndex        =   13
            Top             =   180
            Width           =   2685
         End
         Begin VB.Label vmora 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF8080&
            Height          =   255
            Left            =   2670
            TabIndex        =   15
            Top             =   180
            Width           =   615
         End
      End
      Begin VB.CheckBox c1 
         Appearance      =   0  'Flat
         Caption         =   "Verificando cheques entrantes."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   210
         TabIndex        =   11
         Top             =   510
         Width           =   3735
      End
      Begin VB.CheckBox c2 
         Appearance      =   0  'Flat
         Caption         =   "Acreditación de cheques automáticos."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   210
         TabIndex        =   10
         Top             =   810
         Width           =   3735
      End
      Begin VB.CheckBox c3 
         Appearance      =   0  'Flat
         Caption         =   "Cargando configuración."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   210
         TabIndex        =   9
         Top             =   1110
         Width           =   3735
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   825
         Left            =   150
         TabIndex        =   5
         Top             =   2070
         Width           =   4035
         Begin VB.CommandButton Command1 
            Caption         =   "Actualizar Errores"
            Height          =   435
            Left            =   1800
            TabIndex        =   6
            Top             =   9270
            Width           =   2055
         End
         Begin VB.Label vinfo 
            Height          =   375
            Left            =   1050
            TabIndex        =   8
            Top             =   9300
            Width           =   3645
         End
         Begin VB.Label vinfo2 
            Height          =   375
            Left            =   1020
            TabIndex        =   7
            Top             =   9360
            Width           =   3675
         End
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar ventana "
         Height          =   345
         Left            =   2790
         TabIndex        =   4
         Top             =   7230
         Width           =   1275
      End
      Begin VB.ListBox lstError 
         Height          =   645
         Left            =   150
         TabIndex        =   3
         Top             =   2970
         Width           =   3945
      End
      Begin VB.CommandButton cmdArreglarVista 
         Caption         =   "Arreglar Vista"
         Height          =   345
         Left            =   270
         TabIndex        =   2
         Top             =   7230
         Width           =   2415
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid dgErrores 
         Bindings        =   "Inicio.frx":0054
         Height          =   3495
         Left            =   150
         TabIndex        =   1
         Top             =   3630
         Width           =   3945
         _ExtentX        =   6959
         _ExtentY        =   6165
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4875
         Left            =   -74880
         TabIndex        =   16
         Top             =   1710
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   8599
         _Version        =   393216
         SelectionMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label1 
         Caption         =   "Clientes a los que se le aplicarán el recargo:"
         Height          =   285
         Left            =   -74820
         TabIndex        =   19
         Top             =   600
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCantErrores As Long
Dim connErrores As ADODB.Connection
Dim rsErrores  As ADODB.Recordset
Private Sub cmdArreglarVista_Click()
On Error Resume Next

    frmErrorVistas.Show

If Err Then GrabarLog "cmdArreglarVista_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdCerrar_Click()
On Error Resume Next
    
    Unload Me
    
If Err Then GrabarLog "cmdCerrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Command1_Click()
ControlErrores
End Sub


Private Sub Form_Load()
On Error Resume Next

    lstError.Clear
    ControlErrores
    ControlErroresVista
   
If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ControlErrores()
Dim i As Integer
On Error Resume Next

    Set connErrores = New ADODB.Connection
    Set rsErrores = New ADODB.Recordset
    Dim sqlErrores As String

    With connErrores
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlErrores = "SELECT * FROM Errores"

    With rsErrores
        .CursorLocation = adUseClient
        Call .Open(sqlErrores, connErrores, adOpenStatic, adLockPessimistic)
        
        Set dgErrores.DataSource = rsErrores
        
        If Not .EOF = True Then
            .MoveLast
            vCantErrores = .Fields("Errores").Value
        
            If Not .RecordCount = 0 Then .MoveFirst
        
            For i = 0 To .RecordCount
                FormatoCelda (i)
                .MoveNext
            Next
        
        Else
        
        End If
    
    End With
    
    FormatoGrilla
    
    sqlErrores = ""
    
If Err Then GrabarLog "ControlErrores", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub ControlErroresVista()
On Error Resume Next

    Dim connErroresVista As New ADODB.Connection
    Dim rsErroresVista As New ADODB.Recordset
    Dim sqlErroresVista As String

    With connErroresVista
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlErroresVista = "SELECT * FROM ErrorVista2"

    With rsErroresVista
        Call .Open(sqlErroresVista, connErroresVista, adOpenStatic, adLockPessimistic)
        
        If Not .EOF = True Then

            If (.RecordCount) <> Val(vCantErrores) Then

                lstError.AddItem ("ATENCION !!!!!")
                lstError.AddItem ("Hay " & Trim((.RecordCount - vCantErrores))) & "nuevos/arreglados"
                
                
                rsErrores.AddNew
                rsErrores.Fields("Errores").Value = .RecordCount
                rsErrores.Fields("fecha").Value = Date
                rsErrores.Update
            
            Else
                
                lstError.AddItem ("Estado Normal")
            
            End If
            
        End If

    End With
    
    sqlErroresVista = ""
    
    rsErroresVista.Close
    Set rsErroresVista = Nothing
    
    connErroresVista.Close
    Set connErroresVista = Nothing
    
If Err Then GrabarLog "ControlErroresVista", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub VerificarChequesEntrantes()
On Error Resume Next

    c1.Value = 1
    
    Dim connCheques As New ADODB.Connection
    Dim rsCheques As New ADODB.Recordset
    Dim sqlCheques As String

    With connCheques
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlCheques = "SELECT * FROM cheques WHERE (cp = 'p') AND (estado = 'No Acreditado') AND (deposito <= '" & Str(Date) & "')"
    
    With rsCheques
        Call .Open(sqlCheques, connCheques, adOpenStatic, adLockReadOnly)
    
        If Not .EOF = True Then vinfo.Caption = "> Se han acreditados cheques"
        
    End With
    
    sqlCheques = ""
    
    rsCheques.Close
    Set rsCheques = Nothing
    
    connCheques.Close
    Set connCheques = Nothing

If Err Then GrabarLog "VerificarChequesEntrantes", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub VerificarAcreditaciones()
On Error Resume Next

    c2.Value = 1
    
    Dim connCheques As New ADODB.Connection
    Dim rsCheques As New ADODB.Recordset
    Dim sqlCheques As String

    With connCheques
        .ConnectionString = pathDBMySQL
        .Open
    End With
    
    sqlCheques = "SELECT * FROM cheques WHERE (estado = 'No Acreditado') AND (deposito <= '" & strfechaMySQL(Date) & "')"
    
    With rsCheques
        Call .Open(sqlCheques, connCheques, adOpenStatic, adLockPessimistic)
    
        If Not .EOF Then .MoveFirst

        Do Until .EOF = True
            .Fields("Estado").Value = "Acreditado"
            vinfo2.Caption = "> Hay cheques entrantes"
            .MoveNext
        Loop

    End With
    
    sqlCheques = ""
    
    rsCheques.Close
    Set rsCheques = Nothing
    
    connCheques.Close
    Set connCheques = Nothing
    
If Err Then GrabarLog "VerificarAcreditaciones", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    rsErrores.Close
    Set rsErrores = Nothing

    connErrores.Close
    Set connErrores = Nothing

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla()
On Error Resume Next
    
    With dgErrores
        .SelectionMode = flexSelectionByRow

        .ColWidth(0, 0) = 400
        .ColWidth(1, 0) = 1250
        .ColWidth(2, 0) = 900
        .ColWidth(3, 0) = 900
        .ColWidth(4, 0) = 0
    
        .Redraw = True
        
    End With

        
If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoCelda(vFila As Integer)
On Error Resume Next

    dgErrores.TextMatrix(vFila, 1) = Format(dgErrores.TextMatrix(vFila, 1), "MM/DD/YYYY")
    
If Err Then GrabarLog "FormatoCelda", Err.Number & " " & Err.Description, Me.Name
End Sub

