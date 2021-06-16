VERSION 5.00
Begin VB.Form frmBackup 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   10230
   Begin VB.PictureBox zz 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   3195
      Left            =   0
      Picture         =   "frmBackup.frx":0000
      ScaleHeight     =   3195
      ScaleWidth      =   10185
      TabIndex        =   0
      Top             =   0
      Width           =   10185
      Begin VB.FileListBox lstcopia 
         Height          =   2235
         Left            =   6915
         Pattern         =   "*.zip"
         TabIndex        =   7
         Top             =   0
         Width           =   2025
      End
      Begin VB.Frame Frame1 
         Height          =   1005
         Left            =   2505
         TabIndex        =   5
         Top             =   -105
         Width           =   4410
         Begin VB.Label Label1 
            Caption         =   "Módulo de Backup"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   825
            Left            =   75
            TabIndex        =   6
            Top             =   135
            Width           =   4245
         End
      End
      Begin VB.CommandButton cmdUnZip 
         Caption         =   "Restaurar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   2
         Top             =   1650
         Width           =   1215
      End
      Begin VB.CommandButton cmdZip 
         Caption         =   "Back up"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   90
         TabIndex        =   1
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label lblnombre 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1500
         TabIndex        =   8
         Top             =   1800
         Width           =   60
      End
      Begin VB.Label lblTempDir 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   2760
         Width           =   2910
      End
      Begin VB.Label lblCurDir 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1500
         TabIndex        =   3
         Top             =   1095
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmBackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTempPath _
                Lib "kernel32" _
                Alias "GetTempPathA" (ByVal nBufferLength As Long, _
                                      ByVal lpBuffer As String) As Long
Dim vNombreZip As String
Private Sub CargarCopias()
On Error Resume Next


If Err Then GrabarLog "CargarCopias", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdUnZip_Click()
    Width = 9000
End Sub
Private Sub cmdZip_Click()
    Dim oZip As CGZipFiles
    On Error GoTo vbErrorHandler
    Set oZip = New CGZipFiles

    
    With oZip
        .ZipFileName = vConfigGral.vDireccionDB & "Backup\" & vNombreZip & ".zip"
        .UpdatingZip = False
        .AddFile (vConfigGral.vDireccionDB & "backup\Temp\*.*")

        If .MakeZipFile <> 0 Then
            MsgBox .GetLastMessage
        End If

    End With
    
    Set oZip = Nothing
    
    MsgBox "Copia realizada exitosamente", vbInformation, "Mensaje ..."
    
    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Form1::cmdZip_Click" & " " & Err.Description

End Sub

Private Sub Descomprime()

    On Error GoTo vbErrorHandler

    '
    ' Unzip the ZIPTEST.ZIP file to the Windows Temp Directory
    '
    Dim oUnZip As CGUnzipFiles
    
    Set oUnZip = New CGUnzipFiles
    
    With oUnZip
        '
        ' What Zip File ?
        '
        .ZipFileName = "C:\ZIPTEST.ZIP"
        '
        ' Where are we zipping to ?
        '
        .ExtractDir = GetTempPathName
        '
        ' Keep Directory Structure of Zip ?
        '
        .HonorDirectories = False

        '
        ' Unzip and Display any errors as required
        '
        If .Unzip <> 0 Then
            MsgBox .GetLastMessage
        End If

    End With
    
    Set oUnZip = Nothing
    MsgBox "\ZIPTEST.ZIP Extracted Successfully to " & GetTempPathName

    Exit Sub

vbErrorHandler:
    MsgBox Err.Number & " " & "Form1::cmdUnZip_Click" & " " & Err.Description

End Sub

Private Sub Form_Load()
On Error Resume Next

    With vConfigGral
    
        Call CopiarArchivo(.vDireccionDB & .vEmpresa & ".mdb", .vDireccionDB & "Backup\Temp\" & .vEmpresa & ".mdb", True)
    
        lblTempDir.Caption = GetTempPathName
        lblCurDir.Caption = vConfigGral.vDireccionDB & "backup\Temp\"
        
        vNombreZip = Right(Date, 4) + Left(Right(Date, 7), 2) + Left(Date, 2) & Left(Time, 2) & Mid(Time, 4, 2)
    
        lblCurDir.AutoSize = True
    
        cmdUnZip.Enabled = True
    
        Width = 7005
        Height = 2730
    
        lstcopia.Path = .vDireccionDB & "Backup"
    
    End With
    
    CargarCopias

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    With vConfigGral
        Call BorrarArchivo(.vDireccionDB & "Backup\Temp\" & .vEmpresa & ".mdb")
    End With
    
    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)
    
If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Function GetTempPathName() As String
    Dim sBuffer As String
    Dim lRet As Long
    
    sBuffer = String$(255, vbNullChar)
    
    lRet = GetTempPath(255, sBuffer)
    
    If lRet > 0 Then
        sBuffer = Left$(sBuffer, lRet)
    End If

    GetTempPathName = sBuffer
    
End Function
Private Sub lblnombre_Change()

    If lblNombre.Caption = "" Then
        cmdUnZip.Enabled = False
    Else
        cmdUnZip.Enabled = True
    End If

End Sub
Private Sub lstcopia_Click()
On Error Resume Next

    lblNombre.Caption = Mid(lstcopia.FileName, 7, 2) & "/" & Mid(lstcopia.FileName, 5, 2) & "/" & Left(lstcopia.FileName, 4) & " - " & Mid(lstcopia, 9, 2) & ":" & Mid(lstcopia, 11, 2)

If Err Then GrabarLog "lstcopia_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
