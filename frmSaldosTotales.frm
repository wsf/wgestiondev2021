VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmSaldosTotales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Resumenes agrupados por mes."
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   15165
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc bsaldosTotales 
      Height          =   405
      Left            =   5400
      Top             =   240
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   714
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
      Appearance      =   1
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
      Caption         =   "bsaldosTotales"
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
   Begin TabDlg.SSTab SSTTab0 
      Height          =   6555
      Left            =   360
      TabIndex        =   27
      Top             =   840
      Width           =   14715
      _ExtentX        =   25956
      _ExtentY        =   11562
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Listado"
      TabPicture(0)   =   "frmSaldosTotales.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "dgPagosPorMes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Parametros"
      TabPicture(1)   =   "frmSaldosTotales.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Predicciones"
      TabPicture(2)   =   "frmSaldosTotales.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin MSDataGridLib.DataGrid dgPagosPorMes 
         Bindings        =   "frmSaldosTotales.frx":0054
         Height          =   6045
         Left            =   180
         TabIndex        =   28
         Top             =   360
         Width           =   14430
         _ExtentX        =   25453
         _ExtentY        =   10663
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16777215
         ForeColor       =   4210752
         HeadLines       =   2
         RowHeight       =   15
         RowDividerStyle =   4
         FormatLocked    =   -1  'True
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
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "anomes"
            Caption         =   "Ano/Mes"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "#### / ##"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "SumaDeTotal_ctacte"
            Caption         =   "Total CtaCte"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "SumaDetotal_cdo"
            Caption         =   "Total Ctdo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "SumaDeTotal"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "SumaDepago"
            Caption         =   "Cobrado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "resta"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "SumaDesueldo"
            Caption         =   "Sueldos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   """$"" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   11274
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "eficiencia"
            Caption         =   "% Cobro"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "%##0.00"
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
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1800
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   390
      TabIndex        =   1
      Top             =   7470
      Width           =   14655
      Begin VB.PictureBox Picture1 
         Height          =   345
         Index           =   4
         Left            =   360
         ScaleHeight     =   285
         ScaleWidth      =   765
         TabIndex        =   16
         Top             =   2010
         Width           =   825
      End
      Begin VB.PictureBox Picture1 
         Height          =   345
         Index           =   3
         Left            =   11670
         ScaleHeight     =   285
         ScaleWidth      =   765
         TabIndex        =   15
         Top             =   2010
         Width           =   825
      End
      Begin VB.PictureBox Picture1 
         Height          =   345
         Index           =   2
         Left            =   8040
         ScaleHeight     =   285
         ScaleWidth      =   765
         TabIndex        =   14
         Top             =   2010
         Width           =   825
      End
      Begin VB.PictureBox Picture1 
         Height          =   345
         Index           =   1
         Left            =   4200
         ScaleHeight     =   285
         ScaleWidth      =   765
         TabIndex        =   13
         Top             =   2010
         Width           =   825
      End
      Begin VB.Label vpromecobro 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   6510
         TabIndex        =   24
         Top             =   1650
         Width           =   1305
      End
      Begin VB.Label vpromeventa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   6510
         TabIndex        =   23
         Top             =   1350
         Width           =   1305
      End
      Begin VB.Label vventaporrepartidor 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   2280
         TabIndex        =   22
         Top             =   1680
         Width           =   1425
      End
      Begin VB.Label vventaporcliente 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   2280
         TabIndex        =   21
         Top             =   1380
         Width           =   1425
      End
      Begin VB.Label tsaldo 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Left            =   11580
         TabIndex        =   20
         Top             =   330
         Width           =   2955
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Saldo Total:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   555
         Index           =   4
         Left            =   7950
         TabIndex        =   19
         Top             =   330
         Width           =   3645
      End
      Begin VB.Label tcobro 
         Alignment       =   2  'Center
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   4200
         TabIndex        =   18
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label tventa 
         Alignment       =   2  'Center
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   420
         TabIndex        =   17
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   705
         Index           =   1
         Left            =   11700
         TabIndex        =   12
         Top             =   1410
         Width           =   2745
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   675
         Index           =   0
         Left            =   8040
         TabIndex        =   11
         Top             =   1410
         Width           =   3495
      End
      Begin VB.Label lblEstimacionDe 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Estimacion de Ganancia"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   11580
         TabIndex        =   10
         Top             =   870
         Width           =   2955
      End
      Begin VB.Label lblEficienciaDe 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Quebrantos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   7950
         TabIndex        =   9
         Top             =   870
         Width           =   3645
      End
      Begin VB.Label lblEficienciaDe 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         Caption         =   "Eficiencia de Ventas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   330
         TabIndex        =   8
         Top             =   870
         Width           =   3615
      End
      Begin VB.Label lblPromedios 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Caption         =   "Eficiencia de Cobros"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   3930
         TabIndex        =   7
         Top             =   870
         Width           =   4035
      End
      Begin VB.Label lblPorcentajeDe 
         Caption         =   "> % Cobro por repartidor:"
         Height          =   255
         Index           =   3
         Left            =   3990
         TabIndex        =   6
         Top             =   1680
         Width           =   2445
      End
      Begin VB.Label lblPorcentajeDe 
         Caption         =   "> % Cobro por cliente:"
         Height          =   255
         Index           =   2
         Left            =   4020
         TabIndex        =   5
         Top             =   1350
         Width           =   2505
      End
      Begin VB.Label lblPorcentajeDe 
         Caption         =   "> % venta por repartidor:"
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   4
         Top             =   1710
         Width           =   1905
      End
      Begin VB.Label lblPorcentajeDe 
         Caption         =   "> % de venta por clientes:"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   1410
         Width           =   2115
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   14685
      Begin VB.CommandButton cmdRecalcular 
         Caption         =   "Re - Calcular"
         Height          =   345
         Left            =   13050
         TabIndex        =   26
         Top             =   180
         Width           =   1545
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir"
         Height          =   345
         Left            =   11400
         TabIndex        =   25
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label lblEstadisticasDe 
         Caption         =   "Estadisticas de pagos y cobros por mes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   210
         Width           =   14175
      End
   End
End
Attribute VB_Name = "frmSaldosTotales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdImprimir_Click()
On Error Resume Next

    Unload Mantenimiento
    Load Mantenimiento
    
    MsgBox "Prepare la Impresora !!!", vbInformation, "Mensaje ..."

    With Mantenimiento.rssaldoPorMes
        If .State = 1 Then .Close
        
        .Source = bsaldosTotales.RecordSource
        
        If .State = 0 Then .Open
        .Close
        .Open
    
        If .RecordCount = 0 Then Exit Sub
    End With
    
    With drSaldoPorMes
        .Show
    End With
    
    
If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CalculoInicial()
On Error Resume Next

    Dim vtctacte, vtcdo, vtotal, vtresta, vtpago, vtsueldo, vteficiencia, vcontador As Double

    vtctacte = 0
    vtcdo = 0
    vtotal = 0
    vtresta = 0
    vtpago = 0
    vtsueldo = 0
    vteficiencia = 0
    vcontador = 0

    With bsaldosTotales
        .ConnectionString = pathDBMySQL
        '.RecordSource = "SELECT * FROM SaldosPorMes"
        .RecordSource = "SELECT concat(mid(fdetalle.Fecha, 1,4) , mid(fdetalle.Fecha, 6,2)) + 0 AS anomes,  Sum(fdetalle.Total_ctacte) AS SumaDeTotal_ctacte,  Sum(fdetalle.total_cdo) AS SumaDetotal_cdo,  Sum(fdetalle.Total) AS SumaDeTotal,  Sum(fdetalle.pago) AS SumaDepago,  Sum(fdetalle.resta) AS resta,  Sum(fdetalle.sueldo) AS SumaDesueldo,  Sum(fdetalle.pago) * 100 / Sum(fdetalle.Total) AS eficiencia, Max(fdetalle.Fecha) AS ÚltimoDeFecha FROM fdetalle INNER JOIN Factura ON fdetalle.Remito = Factura.Remito GROUP BY concat(Year(fdetalle.fecha) , Month(fdetalle.fecha));"
        
        .Refresh

        Do Until .Recordset.EOF = True
        
            vtctacte = vtctacte + .Recordset(1).Value
            vtcdo = vtcdo + .Recordset(2).Value
            vtotal = vtotal + .Recordset(3).Value
            vtresta = vtresta + .Recordset(5).Value
            vtpago = vtpago + .Recordset(4).Value
            vtsueldo = vtsueldo + .Recordset(6).Value
            vteficiencia = vteficiencia + .Recordset(7).Value

            vcontador = vcontador + 1

            .Recordset.MoveNext

        Loop
    
    End With
    
    
    Call FormatoGrilla
        
    tsaldo.Caption = vtresta
    tventa.Caption = vtotal
    tcobro.Caption = vtpago
    vpromecobro.Caption = vteficiencia / vcontador

If Err Then GrabarLog "CalculoInicial", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdRecalcular_Click()
On Error Resume Next

    CalculoInicial

If Err Then GrabarLog "cmdRecalcular_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub dgPagosPorMes_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    Call OrdenarDataGrid(ColIndex, bsaldosTotales.Recordset, dgPagosPorMes)

If Err Then GrabarLog "dgPagosPorMes_HeadClick", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub Form_Load()
    On Error Resume Next
    
    Me.Show
    CalculoInicial
    
    If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub FormatoGrilla()
On Error Resume Next
    

If Err Then GrabarLog "FormatoGrilla", Err.Number & " " & Err.Description, Me.Name
End Sub

