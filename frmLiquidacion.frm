VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form Liquidación 
   Caption         =   "breloj.Recordset(1)"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8625
   Begin VB.CheckBox Check1 
      Caption         =   "Todas las fechas"
      Height          =   165
      Left            =   6510
      TabIndex        =   20
      Top             =   1260
      Width           =   1905
   End
   Begin VB.TextBox vfechadesde 
      Height          =   285
      Left            =   4830
      TabIndex        =   19
      Top             =   990
      Width           =   1395
   End
   Begin VB.TextBox vfechahasta 
      Height          =   285
      Left            =   4830
      TabIndex        =   18
      Top             =   1350
      Width           =   1395
   End
   Begin VB.Frame Frame6 
      Height          =   765
      Left            =   7320
      TabIndex        =   16
      Top             =   5280
      Width           =   1215
      Begin VB.CommandButton Command15 
         Caption         =   "Salir"
         Height          =   495
         Left            =   90
         Picture         =   "frmLiquidacion.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         ToolTipText     =   "Salir del módulo de insumos"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.Frame Frame5 
      Height          =   765
      Left            =   90
      TabIndex        =   12
      Top             =   870
      Width           =   2865
      Begin VB.CommandButton Command2 
         Caption         =   "Ejecutar Consulta"
         Height          =   495
         Left            =   1410
         Picture         =   "frmLiquidacion.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Confecciona el listado de la liquidación de sueldos"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Actualizar datos"
         Height          =   495
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Presionar para actualizar el archivo sin salir del programa"
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Frame verinfo 
      Caption         =   "Informe de errores : "
      ForeColor       =   &H00000080&
      Height          =   3435
      Left            =   90
      TabIndex        =   9
      Top             =   1770
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton Command5 
         Caption         =   "Salir"
         Height          =   255
         Left            =   7020
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir los datos procesados anteriormente "
         Top             =   3120
         UseMaskColor    =   -1  'True
         Width           =   1245
      End
      Begin VB.TextBox info 
         Height          =   2865
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   8175
      End
   End
   Begin VB.Frame Frame4 
      Height          =   765
      Left            =   3450
      TabIndex        =   7
      Top             =   5280
      Width           =   1425
      Begin VB.CommandButton Command6 
         Caption         =   "Ver Informe >>"
         Height          =   465
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir los datos procesados anteriormente "
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   1245
      End
   End
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   90
      TabIndex        =   4
      Top             =   5280
      Width           =   3345
      Begin VB.CommandButton Command3 
         Caption         =   "Imprimir Listado"
         Height          =   495
         Left            =   2010
         Picture         =   "frmLiquidacion.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir los datos procesados anteriormente "
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Resumen por Empleados"
         Height          =   495
         Left            =   120
         Picture         =   "frmLiquidacion.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Imprimir los datos procesados anteriormente "
         Top             =   180
         UseMaskColor    =   -1  'True
         Width           =   1905
      End
   End
   Begin MSAdodcLib.Adodc breloj 
      Height          =   330
      Left            =   960
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VbProg\Personal\Datos\Personal.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VbProg\Personal\Datos\Personal.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from reloj,empleado where codigo = empleados order by fecha"
      Caption         =   "breloj"
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
   Begin MSAdodcLib.Adodc bliqui 
      Height          =   330
      Left            =   5640
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VbProg\Personal\Datos\Personal.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VbProg\Personal\Datos\Personal.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Liqui"
      Caption         =   "bliqui"
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
   Begin MSAdodcLib.Adodc bempleados 
      Height          =   330
      Left            =   3360
      Top             =   4560
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VbProg\Personal\Datos\Personal.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\VbProg\Personal\Datos\Personal.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Empleado"
      Caption         =   "bempleados"
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
   Begin RichTextLib.RichTextBox t 
      Height          =   795
      Left            =   1890
      TabIndex        =   2
      Top             =   3210
      Visible         =   0   'False
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   1402
      _Version        =   393217
      Enabled         =   -1  'True
      FileName        =   "C:\Excel\Registro.Rei"
      TextRTF         =   $"frmLiquidacion.frx":0408
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmLiquidacion.frx":0D70
      Height          =   3435
      Left            =   90
      TabIndex        =   3
      Top             =   1770
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   6059
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
      ColumnCount     =   15
      BeginProperty Column00 
         DataField       =   "Empleados"
         Caption         =   "Cod. Empleado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Fecha"
         Caption         =   "Fecha"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "Hora"
         Caption         =   "Hora"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "ES"
         Caption         =   "ES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "reloj.id"
         Caption         =   "reloj.id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "Codigo"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "Empleado"
         Caption         =   "Empleado"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "Precio_hora"
         Caption         =   "Precio_hora"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "Adicional"
         Caption         =   "Adicional"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "Descuentos"
         Caption         =   "Descuentos"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "Fnacimiento"
         Caption         =   "Fnacimiento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column11 
         DataField       =   "Ftrabajo"
         Caption         =   "Ftrabajo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column12 
         DataField       =   "Domicilio"
         Caption         =   "Domicilio"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column13 
         DataField       =   "Comentario"
         Caption         =   "Comentario"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column14 
         DataField       =   "Empleado.id"
         Caption         =   "Empleado.id"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1080
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   3539.906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column11 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column12 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column13 
            ColumnWidth     =   14.74
         EndProperty
         BeginProperty Column14 
            ColumnWidth     =   14.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      ForeColor       =   &H00000080&
      Height          =   705
      Left            =   60
      TabIndex        =   1
      Top             =   90
      Width           =   8415
      Begin VB.TextBox vempleado 
         Height          =   345
         Left            =   3690
         TabIndex        =   0
         Top             =   210
         Width           =   4575
      End
      Begin VB.Label Label2 
         Caption         =   "Ingresar Empleado que desea consultar :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   15
         Top             =   270
         Width           =   3585
      End
   End
   Begin VB.Label Label1 
      Caption         =   "> Hasta :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   3840
      TabIndex        =   22
      Top             =   1380
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "> Desde :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   3840
      TabIndex        =   21
      Top             =   1050
      Width           =   885
   End
End
Attribute VB_Name = "Liquidación"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vcodigo As String
Dim otrodia As Double
Dim i As Double ' variable global para ver d donde sigue

Private Sub Check1_Click()
    vfechadesde.Enabled = Not vfechadesde.Enabled
    vfechahasta.Enabled = Not vfechahasta.Enabled
End Sub

Private Sub Command1_Click()
liquidation
End Sub

Private Sub Command15_Click()
Unload Me
End Sub

Private Sub Command2_Click()
On Error Resume Next

info.Text = ""

calliqui

If vempleado = "" Then
    breloj.RecordSource = "select * from reloj,empleado where codigo = empleados and fecha >= #" + vfechadesde + "# and fecha <= #" + vfechahasta + "# order by fecha, hora"
Else
    breloj.RecordSource = "select * from reloj,empleado where (codigo = empleados and fecha >= #" + vfechadesde + "# and fecha <= #" + vfechahasta + "#) and (empleados = '" + Trim(vcodigo) + "') order by fecha, hora"
End If

breloj.Refresh

If Err Then Exit Sub
End Sub
Private Sub liquidation()
On Error Resume Next

t.FileName = "c:\excel\Registro.rei"
t.Refresh

Dim ultfecha As Date

    'borrar_reloj
    
    breloj.RecordSource = "select * from reloj,empleado where (codigo = empleados)"
    breloj.Refresh
      
    tolast (ultimafecha)
    
Do 'acá donde hago la lectura del archivo. No es la mejor forma de hacerlo; pero anda.
    breloj.Recordset.AddNew
    
    t.SelStart = 0 + i
    t.SelLength = 8
    breloj.Recordset("empleados") = t.SelText
   
    t.SelStart = 9 + i
    t.SelLength = 8
    breloj.Recordset("Fecha") = t.SelText
        
    t.SelStart = 18 + i
    t.SelLength = 5
    breloj.Recordset("hora") = t.SelText
        
    t.SelStart = 24 + i
    t.SelLength = 1
    breloj.Recordset("ES") = t.SelText
      
    breloj.Recordset.Update
    
    i = i + 33
    If Trim(t.SelText) = "" Then
        Exit Sub
    End If
Loop

If Err Then Exit Sub
End Sub
Private Sub tolast(vfecha As Date)
Dim vvfecha As Date

i = 0

Do 'acá donde hago la lectura del archivo. No es la mejor forma de hacerlo; pero anda.
    
    t.SelStart = 0 + i
    t.SelStart = 9 + i
    t.SelLength = 8
    vvfecha = t.SelText
    If vvfecha > vfecha Then
        'i = i + 1  ' lo dejo para q comience del principio
        If i < 0 Then i = 0
        Exit Sub
    End If
           
    t.SelStart = 18 + i
    t.SelStart = 24 + i
    
        
    i = i + 33
    t.SelLength = 1
    If Trim(t.SelText) = "" Then
        Exit Sub
    End If
Loop

End Sub

Function ultimafecha() As Date
On Error Resume Next
    breloj.RecordSource = "select * from reloj,empleado where codigo = empleados order by fecha"
    breloj.Refresh
    breloj.Recordset.MoveLast
    ultimafecha = breloj.Recordset("fecha")
If Err Then
    ultimafecha = Date - 10000000
    Exit Function
End If
End Function
Private Sub calliqui()
On Error Resume Next

borra_liqui

bempleados.Refresh
bempleados.Recordset.MoveFirst

Do Until bempleados.Recordset.EOF
    ' Calcula la liquidación de por cada uno de los empleados
Select Case bempleados.Recordset("tjornada")
    Case 1
        liqui_t1 bempleados.Recordset("Codigo"), bempleados.Recordset("Empleado"), bempleados.Recordset("Tolerancia")
    Case 2
        liqui_t2 bempleados.Recordset("Codigo"), bempleados.Recordset("Empleado"), bempleados.Recordset("Tolerancia")
    Case 3
        liqui_t3 bempleados.Recordset("Codigo"), bempleados.Recordset("Empleado"), bempleados.Recordset("Tolerancia")
    Case 4
        liqui_t4 bempleados.Recordset("Codigo"), bempleados.Recordset("Empleado"), bempleados.Recordset("Tolerancia")
End Select
    bempleados.Recordset.MoveNext
Loop

If Err Then Exit Sub
End Sub
Private Sub liqui_t1(vcodigo As String, vempleado As String, vtolerancia As Double)

breloj.RecordSource = "select * from reloj where empleados = '" + Trim(vcodigo) + "' and fecha >= #" + vfechadesde + "# and fecha <= #" + vfechahasta + "# order by fecha,hora"
breloj.Refresh

' me para en la primera entrada
breloj.Recordset.MoveFirst
breloj.Recordset.Find ("es = 'e'")


Dim vhora1, vhora2 As Timer
Dim vfecha1, vfecha2 As Date

Dim can_jornadas As Integer ' se guardan la cantidad de jornadas trbajadas

Dim vthora, vtotal, vtotalnojornadas As Double
Dim t11, t12, t21, t22 As Double

vthora = 0
breloj.Recordset.MoveFirst

Do Until breloj.Recordset.EOF
    
    ' --------------- primera hora ------------------------
    vhora1 = (breloj.Recordset("hora"))
    vfecha1 = (breloj.Recordset("fecha"))
    ' -----------------------------------------------------
    
    breloj.Recordset.MoveNext
    If breloj.Recordset.EOF Then Exit Do
     
    vfecha2 = (breloj.Recordset("fecha"))
     
    '--------------- segunda hora ----------------
    t11 = (Hour((breloj.Recordset("hora"))) * 60)
    t12 = Minute(breloj.Recordset("hora"))
    
    t21 = (Hour(vhora1) * 60)
    t22 = Minute(vhora1)
    
    
    If vfecha2 > vfecha1 Then
        ' Tipo de jornada 1, entonces esto no puede ocurrir
        info.Text = info.Text + " La jornada de trabajo no es del Tipo 1 (E1,S1,E2,S2) Fecha1 : " + Str(vfecha1) + "  Fecha2 :" + Str(vfecha2)
    Else
        vthora = vthora + ((t11 + t12) - (t21 + t22)) / 60
        'If breloj.Recordset(4) = 25863957 Then MsgBox Str(((t11 + t12) - (t21 + t22)) / 60)
    End If
    
    breloj.Recordset.MoveNext
    
    If (12 - ((t11 + t12) - (t21 + t22)) / 60) < vtolerancia Then ' si la jornada de trabajo es cumplida
        can_jornada = can_jornada + 1
    Else
        info.Text = info.Text + Chr(13) + " Jornada incompleta. Fecha : " + Str(vfecha1)
        vtotalnojornadas = vtotalnojornadas + ((t11 + t12) - (t21 + t22)) / 60 ' acumulación de horas trabajadas cuando no c cumplen las jornadas
    End If
    
Loop

    vthora = can_jornada * 12 + vtotalnojornadas ' cantidad total de horas que se le deben computar
    
    vtotal = bempleados.Recordset("precio_hora") * (vthora) ' le sumo 1, es la hora del armuerzo

bliqui.Refresh
bliqui.Recordset.AddNew
    bliqui.Recordset("Codigo") = vcodigo
    bliqui.Recordset("Empleado") = vempleado
    bliqui.Recordset("Precio_Hora") = bempleados.Recordset("precio_hora")
    bliqui.Recordset("Horas") = vthora
    bliqui.Recordset("Adicional") = bempleados.Recordset("adicional")
    bliqui.Recordset("descuento") = bempleados.Recordset("descuentos")
    bliqui.Recordset("total") = vtotal + bempleados.Recordset("adicional") + bempleados.Recordset("descuentos")
bliqui.Recordset.Update
End Sub

Private Sub liqui_t2(vcodigo As String, vempleado As String, vtolerancia As Double)

' le tengo que sumar una día a la última fecha

breloj.RecordSource = "select * from reloj where empleados = '" + Trim(vcodigo) + "' and fecha >= " + vfechadesde + " and fecha <= #" + Strfecha("#vfechahasta#" + 1) + "# order by fecha,hora"
breloj.Refresh



'-------  me para en la segunda entrada del día. ----------
breloj.Recordset.Find ("es = 'e'")
breloj.Recordset.MoveNext
breloj.Recordset.MoveNext
'---------------------------------------------------------


Dim vhora1, vhora2 As Timer
Dim vfecha1, vfecha2 As Date

Dim can_jornadas  As Integer ' se guardan la cantidad de jornadas trbajadas
Dim intervalo As Integer ' indica si está en el intervalo 1 o 2

Dim vthora, vtotal, vtotalnojornadas As Double
Dim t11, t12, t21, t22 As Double

vthora = 0
'breloj.Recordset.MoveFirst

intervalo = 1

Dim t3 As Integer
t3 = 0
Do Until breloj.Recordset.EOF

'--------- verifico si estoy en el último día ------------
If breloj.Recordset("fecha") = Me.vfechahasta + 1 Then
    If t3 = 4 Then Exit Sub ' lel 4to. mov. pertenece al otro mes.
    t3 = t3 + 1
End If

' -------------------------------------
' identifica en que intervalo está
intervalo = intervalo + 1
If intervalo = 3 Then intervalo = 1
' -------------------------------------
    
    
    ' --------------- primera hora ------------------------
    vhora1 = (breloj.Recordset("hora"))
    vfecha1 = (breloj.Recordset("fecha"))
    ' -----------------------------------------------------
    
    breloj.Recordset.MoveNext
    If breloj.Recordset.EOF Then Exit Do
     
    vfecha2 = (breloj.Recordset("fecha"))
     
    '--------------- segunda hora ----------------
    t11 = (Hour((breloj.Recordset("hora"))) * 60)
    t12 = Minute(breloj.Recordset("hora"))
    
    t21 = (Hour(vhora1) * 60)
    t22 = Minute(vhora1)
    
    
    If vfecha2 > vfecha1 Then
        ' Tipo de jornada 1, entonces esto no puede ocurrir
        If intervalo = 2 Then info.Text = info.Text + " La jornada de trabajo no es del Tipo 3 E1,(S1,E2,S2) Fecha1 : " + Str(vfecha1) + "  Fecha2 :" + Str(vfecha2)
        otrodia = (24 * 60) + (t11 + t12)
        vthora = vthora + (-(t21 + t22) + otrodia) / 60
    Else
        vthora = vthora + ((t11 + t12) - (t21 + t22)) / 60
        'If breloj.Recordset(4) = 25863957 Then MsgBox Str(((t11 + t12) - (t21 + t22)) / 60)
    End If
    
    breloj.Recordset.MoveNext
    
    If (12 - ((t11 + t12) - (t21 + t22)) / 60) < vtolerancia Then ' si la jornada de trabajo es cumplida
        can_jornada = can_jornada + 1
    Else
        info.Text = info.Text + Chr(13) + " Jornada incompleta. Fecha : " + Str(vfecha1)
        vtotalnojornadas = vtotalnojornadas + ((t11 + t12) - (t21 + t22)) / 60 ' acumulación de horas trabajadas cuando no c cumplen las jornadas
    End If
    
Loop

    vthora = can_jornada * 12 + vtotalnojornadas ' cantidad total de horas que se le deben computar
    
    vtotal = bempleados.Recordset("precio_hora") * (vthora) ' le sumo 1, es la hora del armuerzo

bliqui.Refresh
bliqui.Recordset.AddNew
    bliqui.Recordset("Codigo") = vcodigo
    bliqui.Recordset("Empleado") = vempleado
    bliqui.Recordset("Precio_Hora") = bempleados.Recordset("precio_hora")
    bliqui.Recordset("Horas") = vthora
    bliqui.Recordset("Adicional") = bempleados.Recordset("adicional")
    bliqui.Recordset("descuento") = bempleados.Recordset("descuentos")
    bliqui.Recordset("total") = vtotal + bempleados.Recordset("adicional") + bempleados.Recordset("descuentos")
bliqui.Recordset.Update
End Sub




Private Sub liqui_t3(vcodigo As String, vempleado As String, vtolerancia As Double)

' le tengo que sumar una día a la última fecha

breloj.RecordSource = "select * from reloj where empleados = '" + Trim(vcodigo) + "' and fecha >= #" + Strfecha("#vfechadesde#" + 1) + "# and fecha <= #" + Strfecha("#vfechahasta#" + 1) + "# order by fecha,hora"
breloj.Refresh



'-------  me para en la segunda entrada del día. ----------
breloj.Recordset.Find ("es = 'e'")
breloj.Recordset.MoveNext
breloj.Recordset.MoveNext
'---------------------------------------------------------


Dim vhora1, vhora2 As Timer
Dim vfecha1, vfecha2 As Date

Dim can_jornadas  As Integer ' se guardan la cantidad de jornadas trbajadas
Dim intervalo As Integer ' indica si está en el intervalo 1 o 2

Dim vthora, vtotal, vtotalnojornadas As Double
Dim t11, t12, t21, t22 As Double

vthora = 0
'breloj.Recordset.MoveFirst

intervalo = 1

Dim t3 As Integer
t3 = 0
Do Until breloj.Recordset.EOF

'--------- verifico si estoy en el último día ------------
If breloj.Recordset("fecha") = Me.vfechahasta + 1 Then
    If t3 = 4 Then Exit Sub ' lel 4to. mov. pertenece al otro mes.
    t3 = t3 + 1
End If

' -------------------------------------
' identifica en que intervalo está
intervalo = intervalo + 1
If intervalo = 3 Then intervalo = 1
' -------------------------------------
    
    
    ' --------------- primera hora ------------------------
    vhora1 = (breloj.Recordset("hora"))
    vfecha1 = (breloj.Recordset("fecha"))
    ' -----------------------------------------------------
    
    breloj.Recordset.MoveNext
    If breloj.Recordset.EOF Then Exit Do
     
    vfecha2 = (breloj.Recordset("fecha"))
     
    '--------------- segunda hora ----------------
    t11 = (Hour((breloj.Recordset("hora"))) * 60)
    t12 = Minute(breloj.Recordset("hora"))
    
    t21 = (Hour(vhora1) * 60)
    t22 = Minute(vhora1)
    
    
    If vfecha2 > vfecha1 Then
        ' Tipo de jornada 1, entonces esto no puede ocurrir
        If intervalo = 2 Then info.Text = info.Text + " La jornada de trabajo no es del Tipo 3 E1,(S1,E2,S2) Fecha1 : " + Str(vfecha1) + "  Fecha2 :" + Str(vfecha2)
        otrodia = (24 * 60) + (t11 + t12)
        vthora = vthora + (-(t21 + t22) + otrodia) / 60
    Else
        vthora = vthora + ((t11 + t12) - (t21 + t22)) / 60
        'If breloj.Recordset(4) = 25863957 Then MsgBox Str(((t11 + t12) - (t21 + t22)) / 60)
    End If
    
    breloj.Recordset.MoveNext
    
    If (12 - ((t11 + t12) - (t21 + t22)) / 60) < vtolerancia Then ' si la jornada de trabajo es cumplida
        can_jornada = can_jornada + 1
    Else
        info.Text = info.Text + Chr(13) + " Jornada incompleta. Fecha : " + Str(vfecha1)
        vtotalnojornadas = vtotalnojornadas + ((t11 + t12) - (t21 + t22)) / 60 ' acumulación de horas trabajadas cuando no c cumplen las jornadas
    End If
    
Loop

    vthora = can_jornada * 12 + vtotalnojornadas ' cantidad total de horas que se le deben computar
    
    vtotal = bempleados.Recordset("precio_hora") * (vthora) ' le sumo 1, es la hora del armuerzo

bliqui.Refresh
bliqui.Recordset.AddNew
    bliqui.Recordset("Codigo") = vcodigo
    bliqui.Recordset("Empleado") = vempleado
    bliqui.Recordset("Precio_Hora") = bempleados.Recordset("precio_hora")
    bliqui.Recordset("Horas") = vthora
    bliqui.Recordset("Adicional") = bempleados.Recordset("adicional")
    bliqui.Recordset("descuento") = bempleados.Recordset("descuentos")
    bliqui.Recordset("total") = vtotal + bempleados.Recordset("adicional") + bempleados.Recordset("descuentos")
bliqui.Recordset.Update
End Sub

Private Sub liqui_t4(vcodigo As String, vempleado As String, vtolerancia As Double)

' le tengo que sumar una día a la última fecha

breloj.RecordSource = "select * from reloj where empleados = '" + Trim(vcodigo) + "' and fecha >= #" + vfechadesde + "# and fecha <= #" + Strfecha("#vfechahasta#" + 1) + "# order by fecha,hora"
breloj.Refresh



' me para en la primera entrada
breloj.Recordset.MoveFirst
breloj.Recordset.Find ("es = 'e'")


Dim vhora1, vhora2 As Timer
Dim vfecha1, vfecha2 As Date

Dim can_jornadas  As Integer ' se guardan la cantidad de jornadas trbajadas
Dim intervalo As Integer ' indica si está en el intervalo 1 o 2

Dim vthora, vtotal, vtotalnojornadas As Double
Dim t11, t12, t21, t22 As Double

vthora = 0
breloj.Recordset.MoveFirst

intervalo = 1

Dim t4 As Integer

t4 = 0

Do Until breloj.Recordset.EOF

'--------- verifico si estoy en el último día ------------
If breloj.Recordset("fecha") = Me.vfechahasta + 1 Then
    If t4 = 1 Then Exit Sub
    t4 = 1
End If

' -------------------------------------
' identifica en que intervalo está
intervalo = intervalo + 1
If intervalo = 3 Then intervalo = 1
' -------------------------------------
    
    
    ' --------------- primera hora ------------------------
    vhora1 = (breloj.Recordset("hora"))
    vfecha1 = (breloj.Recordset("fecha"))
    ' -----------------------------------------------------
    
    breloj.Recordset.MoveNext
    If breloj.Recordset.EOF Then Exit Do
     
    vfecha2 = (breloj.Recordset("fecha"))
     
    '--------------- segunda hora ----------------
    t11 = (Hour((breloj.Recordset("hora"))) * 60)
    t12 = Minute(breloj.Recordset("hora"))
    
    t21 = (Hour(vhora1) * 60)
    t22 = Minute(vhora1)
    
    
    If vfecha2 > vfecha1 Then
        ' Tipo de jornada 1, entonces esto no puede ocurrir
        If intervalo = 1 Then info.Text = info.Text + " La jornada de trabajo no es del Tipo 4 (E1,S1,E2),S2 Fecha1 : " + Str(vfecha1) + "  Fecha2 :" + Str(vfecha2)
        otrodia = (24 * 60) + (t11 + t12)
        vthora = vthora + (-(t21 + t22) + otrodia) / 60
    Else
        vthora = vthora + ((t11 + t12) - (t21 + t22)) / 60
        'If breloj.Recordset(4) = 25863957 Then MsgBox Str(((t11 + t12) - (t21 + t22)) / 60)
    End If
    
    breloj.Recordset.MoveNext
    
    If (12 - ((t11 + t12) - (t21 + t22)) / 60) < vtolerancia Then ' si la jornada de trabajo es cumplida
        can_jornada = can_jornada + 1
    Else
        info.Text = info.Text + Chr(13) + " Jornada incompleta. Fecha : " + Str(vfecha1)
        vtotalnojornadas = vtotalnojornadas + ((t11 + t12) - (t21 + t22)) / 60 ' acumulación de horas trabajadas cuando no c cumplen las jornadas
    End If
    
Loop

    vthora = can_jornada * 12 + vtotalnojornadas ' cantidad total de horas que se le deben computar
    
    vtotal = bempleados.Recordset("precio_hora") * (vthora) ' le sumo 1, es la hora del armuerzo

bliqui.Refresh
bliqui.Recordset.AddNew
    bliqui.Recordset("Codigo") = vcodigo
    bliqui.Recordset("Empleado") = vempleado
    bliqui.Recordset("Precio_Hora") = bempleados.Recordset("precio_hora")
    bliqui.Recordset("Horas") = vthora
    bliqui.Recordset("Adicional") = bempleados.Recordset("adicional")
    bliqui.Recordset("descuento") = bempleados.Recordset("descuentos")
    bliqui.Recordset("total") = vtotal + bempleados.Recordset("adicional") + bempleados.Recordset("descuentos")
bliqui.Recordset.Update
End Sub



Function Hour1(vhora As String) As Double
If Right(vhora, 4) = "p.m." Then
    Hour1 = Hour(vhora) + 20
Else
    Hour1 = Hour(vhora) + 20
End If

End Function

Private Sub borrar_reloj()
On Error Resume Next
    breloj.Refresh
    
    breloj.Recordset.AddNew
    breloj.Recordset("empleados") = 1
    breloj.Recordset.Update
    
    breloj.Refresh
    
    breloj.Recordset.MoveFirst
    Do Until breloj.Recordset.EOF
        breloj.Recordset.Delete
        breloj.Recordset.MoveNext
    Loop
If Err Then Exit Sub
End Sub
Private Sub borra_liqui()
On Error Resume Next

    bliqui.Recordset.AddNew
    bliqui.Recordset("empleado") = "1"
    bliqui.Recordset.Update
    
    bliqui.Refresh
    bliqui.Recordset.MoveFirst
    Do Until bliqui.Recordset.EOF
        bliqui.Recordset.Delete
        bliqui.Recordset.MoveNext
    Loop
If Err Then Exit Sub
End Sub

Private Sub Command3_Click()
MousePointer = vbHourglass
Me.Refresh
Unload Me
Me.Show
refre
mantenimiento.rscliqui.Sort = "codigo"
drliqui.Refresh
drliqui.Show
MousePointer = vbDefault
End Sub
Private Sub refre()
On Error Resume Next
mantenimiento.rscliqui.Close
mantenimiento.rscliqui.Open
If Err Then Exit Sub
End Sub

Private Sub Command4_Click()

If Not Trim(vempleado.Text) = "" Then
    mantenimiento.rscresumen.Filter = "liqui.codigo = '" + vcodigo + "'"
End If
drdetalle.Show
End Sub

Private Sub Command5_Click()
verinfo.Visible = False
End Sub

Private Sub Command6_Click()
verinfo.Visible = True
End Sub

Private Sub Form_Click()
'MsgBox DateDiff("d", #8/5/2004#, #8/6/2004#)
'MsgBox #7/31/2004# + 1
End Sub

Private Sub Form_Load()
Me.Left = 1600
Me.Top = 400
Me.Width = 8745
Me.Height = 6525

vfechadesde = Date
vfechahasta = Date

'liquidation

End Sub

Private Sub vempleado_KeyPress(KeyAscii As Integer)
On Error Resume Next
If KeyAscii = 13 Then

bempleados.Refresh
bempleados.Recordset.Find ("empleado like '%" + vempleado.Text + "%'")

'bempleados.
'bemplaedos.RecordSource = "select * from empleado where empleado.empleado like '%" + Trim(vempleado) + "%'"
'bempleados.Refresh
'bempleados.Recordset.MoveFirst

vcodigo = bempleados.Recordset("codigo")
vempleado = bempleados.Recordset("empleado")
End If

If Err Then Exit Sub
End Sub

Private Sub modelo_liqui_t1(vcodigo As String, vempleado As String)
' (E1 S1 E2 S2)


breloj.RecordSource = "select * from reloj where empleados = '" + Trim(vcodigo) + "' and fecha >= #" + vfechadesde + "# and fecha <= #" + vfechahasta + "# order by fecha,hora"
breloj.Refresh

'If vcodigo = "25863963" Then
'    MsgBox "25863963"
'End If

Dim vhora1, vhora2 As Timer
Dim vfecha1, vfecha2 As Date

Dim vthora, vtotal As Double
Dim t11, t12, t21, t22 As Double

vthora = 0
breloj.Recordset.MoveFirst

Do Until breloj.Recordset.EOF
    
    ' --------------- primera hora ------------------------
    vhora1 = (breloj.Recordset("hora"))
    vfecha1 = (breloj.Recordset("fecha"))
    ' -----------------------------------------------------
    
    breloj.Recordset.MoveNext
    If breloj.Recordset.EOF Then Exit Do
     
    vfecha2 = (breloj.Recordset("fecha"))
     
    ' segunda hora
    t11 = (Hour((breloj.Recordset("hora"))) * 60)
    t12 = Minute(breloj.Recordset("hora"))
    
    t21 = (Hour(vhora1) * 60)
    t22 = Minute(vhora1)
    
    
    If vfecha2 > vfecha1 Then
        otrodia = (24 * 60) + (t11 + t12)
        vthora = vthora + (-(t21 + t22) + otrodia) / 60
        
       ' If breloj.Recordset(4) = 25863957 Then MsgBox Str(((t21 + t22) + otrodia) / 60)
    Else
        vthora = vthora + ((t11 + t12) - (t21 + t22)) / 60
        'If breloj.Recordset(4) = 25863957 Then MsgBox Str(((t11 + t12) - (t21 + t22)) / 60)
    End If
    
    breloj.Recordset.MoveNext
Loop

vtotal = bempleados.Recordset("precio_hora") * (vthora) ' le sumo 1, es la hora del armuerzo

bliqui.Refresh
bliqui.Recordset.AddNew
    bliqui.Recordset("Codigo") = vcodigo
    bliqui.Recordset("Empleado") = vempleado
    bliqui.Recordset("Precio_Hora") = bempleados.Recordset("precio_hora")
    bliqui.Recordset("Horas") = vthora
    bliqui.Recordset("Adicional") = bempleados.Recordset("adicional")
    bliqui.Recordset("descuento") = bempleados.Recordset("descuentos")
    bliqui.Recordset("total") = vtotal + bempleados.Recordset("adicional") + bempleados.Recordset("descuentos")
bliqui.Recordset.Update


End Sub

