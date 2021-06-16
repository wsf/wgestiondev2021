VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEstructuraCaja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estructura de conceptos de ingresos y egresos de caja"
   ClientHeight    =   9150
   ClientLeft      =   2490
   ClientTop       =   -2475
   ClientWidth     =   12690
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9150
   ScaleWidth      =   12690
   Begin VB.OptionButton OptEgreso 
      Caption         =   "Egreso"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7800
      TabIndex        =   45
      Top             =   1440
      Width           =   975
   End
   Begin VB.OptionButton OptIngreso 
      Caption         =   "Ingreso"
      Enabled         =   0   'False
      Height          =   255
      Left            =   6600
      TabIndex        =   44
      Top             =   1440
      Width           =   975
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   8640
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   1
   End
   Begin MSComctlLib.StatusBar BarraEstado 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   35
      Top             =   8895
      Width           =   12690
      _ExtentX        =   22384
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   8535
      Left            =   6240
      TabIndex        =   1
      Top             =   0
      Width           =   6375
      Begin VB.Frame fraBusqueda 
         Caption         =   "Busqueda de conceptos"
         Height          =   2895
         Left            =   120
         TabIndex        =   20
         Top             =   3600
         Width           =   6135
         Begin VB.CheckBox chkCliente 
            Caption         =   "Buscar ingresos"
            Height          =   255
            Left            =   4560
            TabIndex        =   43
            Top             =   840
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.CheckBox chkProv 
            Caption         =   "Buscar egresos"
            Height          =   255
            Left            =   4560
            TabIndex        =   42
            Top             =   1200
            Value           =   1  'Checked
            Width           =   1455
         End
         Begin VB.TextBox txtConceptoBuqueda 
            Height          =   285
            Left            =   1200
            TabIndex        =   38
            Top             =   1920
            Width           =   4815
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "Imprimir"
            Height          =   495
            Left            =   4080
            TabIndex        =   37
            Top             =   2280
            Width           =   975
         End
         Begin VB.CheckBox chkFecha 
            Caption         =   "Filtrar por fechas"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox txtProveedor 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   32
            Top             =   1200
            Width           =   3255
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "Buscar"
            Height          =   495
            Left            =   5160
            Picture         =   "frmEstructuraCaja.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   29
            ToolTipText     =   "Ejecutar búsqueda"
            Top             =   2280
            UseMaskColor    =   -1  'True
            Width           =   885
         End
         Begin VB.TextBox txtUsuario 
            Height          =   285
            Left            =   1200
            TabIndex        =   28
            Top             =   1560
            Width           =   4815
         End
         Begin VB.TextBox txtCliente 
            Height          =   285
            Left            =   1200
            TabIndex        =   26
            Top             =   840
            Width           =   3255
         End
         Begin MSComCtl2.DTPicker dtDesde 
            Height          =   315
            Left            =   2520
            TabIndex        =   21
            Top             =   360
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   71041025
            CurrentDate     =   40151
         End
         Begin MSComCtl2.DTPicker dtHasta 
            Height          =   315
            Left            =   4680
            TabIndex        =   22
            Top             =   360
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   71041025
            CurrentDate     =   40151
         End
         Begin VB.Label Label4 
            Caption         =   "> Concepto:"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label5 
            Caption         =   "> Proveedor:"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "> Usuario:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1560
            Width           =   855
         End
         Begin VB.Label Label2 
            Caption         =   "> Cliente:"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "> Hasta:"
            Height          =   255
            Left            =   3960
            TabIndex        =   24
            Top             =   390
            Width           =   735
         End
         Begin VB.Label lblFechaDesde 
            Caption         =   "> Desde:"
            Height          =   255
            Left            =   1800
            TabIndex        =   23
            Top             =   390
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdGuardar 
         Appearance      =   0  'Flat
         Caption         =   "Aceptar"
         Height          =   495
         Left            =   4560
         Picture         =   "frmEstructuraCaja.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Guardar y cerrar"
         Top             =   7920
         UseMaskColor    =   -1  'True
         Width           =   825
      End
      Begin VB.Frame fraOpera 
         Caption         =   "Operaciones sobre el concepto seleccionado"
         Height          =   1215
         Left            =   120
         TabIndex        =   7
         Top             =   6600
         Width           =   6135
         Begin VB.CommandButton cmdAñadir 
            Caption         =   "Agregar"
            Height          =   375
            Left            =   3960
            TabIndex        =   12
            ToolTipText     =   " Añadir el texto indicado al nodo seleccionado (como nodo hijo) "
            Top             =   720
            Width           =   1005
         End
         Begin VB.CommandButton cmdBorrar 
            Caption         =   "Borrar "
            Height          =   375
            Left            =   5040
            TabIndex        =   11
            Top             =   720
            Width           =   1005
         End
         Begin VB.CommandButton cmdRenombrar 
            Caption         =   "Renombrar "
            Height          =   375
            Left            =   2880
            TabIndex        =   10
            ToolTipText     =   " Sustituir el texto del nodo seleccionado por el indicado "
            Top             =   720
            Width           =   1005
         End
         Begin VB.TextBox txtConcepto 
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            Top             =   360
            Width           =   4815
         End
         Begin VB.Label lblConcepto1 
            Caption         =   "> Concepto : "
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informacion sobre el concepto seleccionado"
         Height          =   1575
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6135
         Begin VB.TextBox txtCodigoConcepto 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   14
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtConceptoVista 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1200
            TabIndex        =   5
            Top             =   360
            Width           =   4815
         End
         Begin VB.CheckBox chkTieneSubconcepto 
            Caption         =   "Tiene subconcepto"
            Enabled         =   0   'False
            Height          =   255
            Left            =   2760
            TabIndex        =   4
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label lblCodigoConcepto 
            Caption         =   "> Codigo : "
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   720
            Width           =   1515
         End
         Begin VB.Label lblConcepto 
            Caption         =   "> Concepto : "
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   360
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdCerrar 
         Appearance      =   0  'Flat
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   5400
         Picture         =   "frmEstructuraCaja.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cerrar"
         Top             =   7920
         UseMaskColor    =   -1  'True
         Width           =   795
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   4920
         Top             =   1920
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   4
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEstructuraCaja.frx":078E
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEstructuraCaja.frx":0D28
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEstructuraCaja.frx":12C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEstructuraCaja.frx":185C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label lblSaldoAnterior 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4440
         TabIndex        =   41
         Top             =   3200
         Width           =   1695
      End
      Begin VB.Label lblSalAnt 
         Caption         =   "> Saldo anterior:"
         Height          =   255
         Left            =   3120
         TabIndex        =   40
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label lblSal 
         Caption         =   "> Saldo:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label lblSaldo 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1080
         TabIndex        =   30
         Top             =   3200
         Width           =   1695
      End
      Begin VB.Label lblEgresoValue 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1080
         TabIndex        =   19
         Top             =   2600
         Width           =   1695
      End
      Begin VB.Label lblEgresos 
         Caption         =   "> Egresos:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lblIngresoValue 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1034
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   375
         Left            =   1080
         TabIndex        =   17
         Top             =   1995
         Width           =   1695
      End
      Begin VB.Label lblIngresos 
         Caption         =   "> Ingresos:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   975
      End
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   8415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   14843
      _Version        =   393217
      Style           =   7
      ImageList       =   "ImageList2"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmEstructuraCaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mostrarSaldos As Boolean
Private Calculado As Boolean
Private estaBuscando As Boolean
Private sqlCaja As String
Public leido As Boolean
Public saldototal As Double
Public vModo As Modo

Public Enum Modo
    Creacion
    Lectura
    Seleccion
End Enum

Private Sub cmdAdd_Click()
On Error Resume Next

    AgregarNodo

If Err Then GrabarLog "cmdAdd_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub AgregarNodo()
On Error Resume Next
    ' Añadir el contenido del txtConcepto al nodo actual
    Dim tNodo As Node
    Dim sP As String, sH As String
    Dim i As Long
    
    ' El nodo que está actualmente seleccionado
    Set tNodo = TreeView1.SelectedItem
   
    'If Me.TreeView1.SelectedItem Is Not Null Then
    
    sP = tNodo.Key
    ' Cantidad de hijos
    i = tNodo.Children
    
        Do
            i = i + 1
            sH = sP & "-" & CStr(i)
            Err = 0
            ' Añadirlo como nodo hijo del seleccionado
            TreeView1.Nodes.Add sP, tvwChild, sH, Me.txtConcepto.Text, 4, 4
            ' Si no da error, salir del bucle
        
            If Err.Description <> "" Then
                MsgBox "Debe seleccionar un nodo padre para el nuevo concepto"
                Exit Sub
            Else
                Exit Do
            End If
        
            Exit Do
        Loop
    
    
    Dim ingresoEgreso As Boolean
    Dim spSinLetras As String
    spSinLetras = Mid(sP, 3, ContarCaracteres(sP, "-") - 2)
    ingresoEgreso = ActualizarTieneSubconcepto(CLng(spSinLetras))
    GuardarConcepto CLng(spSinLetras), ingresoEgreso
    
    'Else
    'MsgBox "Debe seleccionar un nodo padre para el nuevo concepto"
    'End If
If Err Then
    MsgBox "Error al guardar el concepto " & Me.txtConcepto.Text
    GrabarLog "AgregarNodo", Err.Number & " " & Err.Description, Me.Name
End If

End Sub

'Retorna verdadero si es un ingreso, caso contrario falso
Private Function ActualizarTieneSubconcepto(Codigo As Long) As Boolean
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
    
    sqlConcepto = "Select * from Concepto"
  
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
    
    If Not rsConcepto.EOF Then rsConcepto.MoveFirst
    
    Do While Not rsConcepto.EOF
        If rsConcepto("Codigo") = Codigo Then
            ActualizarTieneSubconcepto = rsConcepto("IngresoEgreso").Value
            If CBool(rsConcepto("TieneSubconcepto").Value) = False Then
                rsConcepto("TieneSubconcepto") = True
                rsConcepto.Update
            End If
            Exit Function
        End If
        rsConcepto.MoveNext
    Loop
End Function
Private Function ActualizarTieneSubconceptoEnEliminacion(codigoPadre As Long) As Boolean
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    
    cmdConcepto.ActiveConnection = ConnDDBB
    
    'Verifico si el codigo del padre del nodo eliminado es padre de alguien mas
    sqlConcepto = "SELECT * FROM Concepto WHERE (codigoPadre = " & codigoPadre & ")"
  
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.CursorLocation = adUseClient
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
    
    'Si no es padre de nadie mas entonces ese nodo no tiene subconceptos
    If rsConcepto.RecordCount = 0 Then
        ActualizarConceptoPadre codigoPadre
    End If
        
End Function
Private Sub ActualizarConceptoPadre(codigoPadre As Long)
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    
    cmdConcepto.ActiveConnection = ConnDDBB
    
    sqlConcepto = "Select * from Concepto where codigo = " & codigoPadre
  
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.CursorLocation = adUseClient
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
    
    If rsConcepto.RecordCount = 1 Then
        rsConcepto.MoveFirst
        rsConcepto("TieneSubconcepto").Value = False
        rsConcepto.Update
    End If

End Sub

Private Sub GuardarConcepto(codigoPadre As Long, ingresoEgreso As Boolean)
On Error Resume Next
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
    
    sqlConcepto = "Select * from Concepto where 0"
  
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
    
    If Not rsConcepto.EOF Then
        rsConcepto.MoveFirst
    End If
    
    rsConcepto.AddNew
    rsConcepto("Descripcion") = Me.txtConcepto.Text
    rsConcepto("TieneSubconcepto") = False
    rsConcepto("IngresoEgreso") = ingresoEgreso
    rsConcepto("CodigoPadre") = codigoPadre
    rsConcepto.Update
If Err Then MsgBox "Error al guardar el nodo " & Me.txtConceptoVista.Text
End Sub

'Funcion recursiva que elimina el concepto y sus subconceptos
Private Sub EliminarConcepto(codigoPadre As Long, tieneSubconcepto As Boolean)
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
    
    sqlConcepto = "Select * from Concepto where CodigoPadre = " & codigoPadre
  
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
    
    If Not rsConcepto.EOF Then
        rsConcepto.MoveFirst
    End If
    
    Do While Not rsConcepto.EOF
        If rsConcepto("TieneSubconcepto") = True Then
            EliminarConcepto rsConcepto("Codigo"), rsConcepto("TieneSubconcepto")
            'rsConcepto("TieneSubconcepto") = False
            rsConcepto.Delete
            rsConcepto.Update
        Else
        
            If rsConcepto("Descripcion") <> "Ingresos" And rsConcepto("Descripcion") <> "Egresos" Then
                'Si no tiene subconcepto lo elimino
                rsConcepto.Delete
                rsConcepto.Update
            Else
                MsgBox "No se permite eliminar el concepto " & rsConcepto("Descripcion")
            End If
        End If
        rsConcepto.MoveNext
    Loop
    
End Sub
Private Sub AgregarNodoConcepto(idPadre As Long, idConcepto As Long, descripcion As String, esIngreso As Boolean, tieneSubconcepto As Boolean)
   If Not estaBuscando Then
    
    
    If idPadre = -1 Then
    
        Dim nodoPadre As Node
        Me.TreeView1.LineStyle = tvwTreeLines
        Me.TreeView1.Style = tvwTreelinesPlusMinusPictureText
        Set nodoPadre = TreeView1.Nodes.Add(, , "R " & CStr(idConcepto), descripcion, 3, 3)
        Me.TreeView1.Nodes("R " & CStr(idConcepto)).Expanded = True
        
    Else
        If Not tieneSubconcepto Then
            TreeView1.Nodes.Add "R " & CStr(idPadre), tvwChild, "R " & CStr(idConcepto), descripcion, 4, 4
        Else
            If esIngreso Then
                TreeView1.Nodes.Add "R " & CStr(idPadre), tvwChild, "R " & CStr(idConcepto), descripcion, 2, 2
            Else
                TreeView1.Nodes.Add "R " & CStr(idPadre), tvwChild, "R " & CStr(idConcepto), descripcion, 1, 1
            End If
        End If
        'Expandir el nodo
        Me.TreeView1.Nodes("R " & CStr(idConcepto)).Expanded = True
    End If
   Dim b  As Boolean
   b = ActualizarTieneSubconcepto(CLng(idPadre))
   Else
    
   End If
End Sub

Private Sub cmdBorrarNodo_Click()
    Dim tNodo As Node
    Dim i As Long
    '
    ' El nodo que está actualmente seleccionado
    Set tNodo = TreeView1.SelectedItem
    i = tNodo.Children
    '
    ' Avisar que se va a borrar un nodo que tiene hijos
    If i > 0 Then
        If MsgBox("¿Quiere borrar el nodo con " & CStr(i) & " hijos?", vbQuestion Or vbYesNo, "Borrar nodos") = vbNo Then
            Exit Sub
        End If
    End If
    TreeView1.Nodes.Remove tNodo.Index
End Sub


Private Sub cmdAñadir_Click()
    AgregarNodo
    Unload Me
    Load Me
    ActualizarDespuesOperar
End Sub

Private Sub BorrarNodo()
    Dim tNodo As Node
    Dim i As Long
    
    ' El nodo que está actualmente seleccionado
    Set tNodo = TreeView1.SelectedItem
    i = tNodo.Children
    Dim padre As String
    padre = tNodo.Parent.Key
    
    ' Avisar que se va a borrar un nodo que tiene hijos
    If i > 0 Then
        If MsgBox("¿Esta seguro que quiere borrar el concepto con " & CStr(i) & " subconceptos?", vbQuestion Or vbYesNo, "Borrar nodos") = vbNo Then
            Exit Sub
        End If
    End If
    TreeView1.Nodes.Remove tNodo.Index
    Dim padreAActualizar As Long
    If i = 0 Then
        
        padreAActualizar = EliminarUnico(tNodo.Key)
        
    Else
        EliminarConcepto Mid(tNodo.Key, 3, ContarCaracteres(tNodo.Key, "-") - 2), True
        padreAActualizar = EliminarUnico(tNodo.Key)
    End If
    
    'Actualizar tieneconcepto del padre
    ActualizarTieneSubconceptoEnEliminacion (padreAActualizar)
End Sub

Private Function EliminarUnico(claveNodo As String)
          'EliminarConcepto Mid(tNodo.Key, 3, ContarCaracteres(tNodo.Key, "-") - 2), False
          Dim cmdConcepto As New ADODB.Command
          Dim sqlConcepto As String
          cmdConcepto.ActiveConnection = ConnDDBB
          
          sqlConcepto = "SELECT * FROM Concepto WHERE (Codigo = " & Mid(claveNodo, 3, ContarCaracteres(claveNodo, "-") - 2) & ")"
        
          Dim rsConcepto As New ADODB.Recordset
          
          If rsConcepto.State = 0 Then
              rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
          Else
              Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
          End If
          
          If Not rsConcepto.EOF Then
              rsConcepto.MoveFirst
          End If
              
          If rsConcepto("Descripcion").Value <> "Ingresos" And rsConcepto("Descripcion").Value <> "Egresos" Then
                'Si no tiene subconcepto y el valor en la caja de este concepto = 0 lo elimino
                
                If CalcularTotalConcepto(rsConcepto("Codigo").Value, rsConcepto("IngresoEgreso").Value) > 0 Then
                    MsgBox "No se puede eliminar el concepto " & rsConcepto("Descripcion") & vbCrLf & _
                    "debido a que esta en uso en la caja. ", vbInformation
                Else
                    Dim padreAActualizar As Long
                    padreAActualizar = rsConcepto("CodigoPadre").Value
                    rsConcepto.Delete
                    rsConcepto.Update
                End If
            Else
                MsgBox "No se permite eliminar el concepto " & rsConcepto("Descripcion").Value
            End If
            
            EliminarUnico = padreAActualizar
End Function

Private Sub cmdLlenarTree_Click()
    LlenarArbol
End Sub

Private Sub LlenarArbol()
    ActivarBotones False
    Me.TreeView1.Nodes.Clear
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
    
    sqlConcepto = "SELECT * FROM Concepto"
  
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    End If
    
    If Not rsConcepto.EOF Then rsConcepto.MoveFirst
    
    Dim idPadre As Long
    Dim idConcepto As Long
    Dim idConceptoAux As Long
    
    'Primero busco la raiz
    Do While Not rsConcepto.EOF
        If rsConcepto("Descripcion") = "Raiz" Then
            idConcepto = rsConcepto("Codigo").Value
            AgregarNodoConcepto -1, rsConcepto("Codigo").Value, rsConcepto("Descripcion").Value, rsConcepto("IngresoEgreso").Value, rsConcepto("TieneSubconcepto").Value
            DoEvents
            Exit Do
        End If
        rsConcepto.MoveNext
    Loop
    If mostrarSaldos = True Then
        If Calculado = True Then
            ArmarResumenConceptoCalculado idConcepto
        Else
            ArmarResumenConceptos idConcepto
        End If
    Else
        ArmarArbol idConcepto
    End If
    ActivarBotones True
End Sub

Private Sub ArmarArbol(idRaiz As Long)
   
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
    
    If Not idRaiz = 1 Then
        sqlConcepto = "Select * from Concepto ORDER BY Descripcion"
    Else
        sqlConcepto = "Select * from Concepto"
    End If
    
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    End If
    
    If Not rsConcepto.EOF = True Then
        rsConcepto.MoveFirst
    End If
    
    Dim idPadre As Long
    Dim idConcepto As Long
    Dim idConceptoAux As Long

    Dim i As Integer
        
    Do While Not rsConcepto.EOF = True
        If rsConcepto("CodigoPadre").Value = idRaiz Then
            AgregarNodoConcepto idRaiz, rsConcepto("Codigo"), rsConcepto("Descripcion"), rsConcepto("IngresoEgreso"), rsConcepto("TieneSubconcepto")
            DoEvents
            If CBool(rsConcepto("TieneSubconcepto").Value) = True Then
                ArmarArbol rsConcepto("Codigo").Value
            End If
            idConcepto = rsConcepto("Codigo").Value
        End If
        rsConcepto.MoveNext
    Loop
          
End Sub
Private Sub ArmarResumenConceptoCalculado(idRaiz As Long)
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
    
    sqlConcepto = "SELECT * FROM Concepto"
  
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    End If
    
    If Not rsConcepto.EOF Then
        rsConcepto.MoveFirst
    End If
    
    Dim idPadre As Long
    Dim idConcepto As Long
    Dim idConceptoAux As Long

    Dim i As Integer
    Dim totalconcepto As Double
    Do While Not rsConcepto.EOF
        If rsConcepto("CodigoPadre") = idRaiz Then
            'Si no es una hoja sumar los valores de sus hijos
            
            totalconcepto = CalcularTotalConceptoRec(rsConcepto("Codigo"), rsConcepto("TieneSubConcepto"), rsConcepto("IngresoEgreso"))
            AgregarNodoConcepto idRaiz, rsConcepto("Codigo"), rsConcepto("Descripcion") & " $ " & totalconcepto, rsConcepto("IngresoEgreso"), rsConcepto("TieneSubconcepto")
            DoEvents
            
            If Me.ProgressBar1.Value < Me.ProgressBar1.Max Then
                Dim inc As Double
                If estaBuscando Then
                    inc = 1 / 2
                Else
                    inc = 1
                End If
                Me.ProgressBar1.Value = Me.ProgressBar1.Value + inc
            End If
            
            If CBool(rsConcepto("TieneSubconcepto").Value) = True Then
                ArmarResumenConceptoCalculado rsConcepto("Codigo")
            End If
            idConcepto = rsConcepto("Codigo").Value
            
             If rsConcepto("Descripcion") = "Ingresos" Then
                Me.lblIngresoValue.Caption = Format(totalconcepto, "#######0.00")
            End If
            If rsConcepto("Descripcion") = "Egresos" Then
                Me.lblEgresoValue.Caption = Format(totalconcepto, "#######0.00")
            End If
        End If
        
       
        rsConcepto.MoveNext
    Loop
          
End Sub

Private Function CalcularTotalConceptoPadre(codigoPadre As Long)
    Dim total As Double
    Dim cmdCaja As New ADODB.Command
    Dim sqlCaja As String
    cmdCaja.ActiveConnection = ConnDDBB
      
    sqlCaja = "Select * from (Concepto co inner join Caja ca on co.Codigo = ca.CodigoConcepto) where co.CodigoPadre = " & Str(codigoPadre)
    
    Dim rsCaja As New ADODB.Recordset
      
    If rsCaja.State = 0 Then
        rsCaja.Open sqlCaja, ConnDDBB, 3, 3
    Else
        Set rsCaja = ConnDDBB.Execute(sqlCaja)
    End If
      
    If Not rsCaja.EOF Then
        rsCaja.MoveFirst
    End If
    
    Do While Not rsCaja.EOF
        total = total + rsCaja("Importe")
        rsCaja.MoveNext
    Loop
    
    CalcularTotalConceptoPadre = total
End Function

Private Function ArmarResumenConceptos(idRaiz As Long) As Double
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
    
    sqlConcepto = "SELECT * FROM Concepto"
  
    Dim rsConcepto As New ADODB.Recordset
    
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    End If
    
    If Not rsConcepto.EOF Then rsConcepto.MoveFirst
    
    
    Dim idPadre As Long
    Dim idConcepto As Long
    Dim idConceptoAux As Long
    Dim saldoConcepto  As Double
    Dim i As Integer
        
    Do While Not rsConcepto.EOF = True
        If rsConcepto("CodigoPadre").Value = idRaiz Then
            saldoConcepto = CalcularTotalConceptoRec(rsConcepto("Codigo").Value, rsConcepto("TieneSubconcepto").Value, rsConcepto("IngresoEgreso").Value)
            
            AgregarNodoConcepto idRaiz, rsConcepto("Codigo").Value, rsConcepto("Descripcion").Value & " $" & saldoConcepto, rsConcepto("IngresoEgreso").Value, rsConcepto("TieneSubconcepto").Value
            
            ProgressBar1.Max = ProgressBar1.Max + 1
            
            If CBool(rsConcepto("TieneSubconcepto").Value) = True Then
                ArmarResumenConceptos = ArmarResumenConceptos + ArmarResumenConceptos(rsConcepto("Codigo"))
            End If
            GuardarTotalConcepto rsConcepto("Codigo").Value, ArmarResumenConceptos
            idConcepto = rsConcepto("Codigo").Value
        End If
        rsConcepto.MoveNext
    Loop
          
End Function

Private Sub GuardarTotalConcepto(Codigo As Long, saldo As Double)
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
      
    sqlConcepto = "Select * from Concepto where Codigo = " & Str(Codigo)
    
    Dim rsConcepto As New ADODB.Recordset
      
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
      
    If Not rsConcepto.EOF Then
        rsConcepto.MoveFirst
        rsConcepto("SaldoTotal").Value = saldo
        rsConcepto.Update
    End If
    
    
End Sub
Private Sub cmdBorrar_Click()
On Error Resume Next

    BorrarNodo
    Unload Me
    Load Me
    ActualizarDespuesOperar

If Err Then GrabarLog "cmdBorrar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdBuscar_Click()
On Error Resume Next

    If ValidarFecha() = True Then
        Buscar
        lblSaldoAnterior = Format(CalcularSaldoAnterior(), "#######0.00")
    End If

If Err Then GrabarLog "cmdBuscar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Public Sub Buscar()
On Error Resume Next

    ActivarBotones False
    CambiarStatus "Buscando", "Filtrando por los campos seleccionados...", App.Path & "\Imagenes\iconos\250.ico"
    MousePointer = 11
   
    InicializarLlenadoArbol
    
    ActivarBotones True
    MousePointer = Default
    CambiarStatus "WSF", "WSF - Sistema integral de caja diaria", App.Path & "\Imagenes\iconos\144.ico"

If Err Then GrabarLog "Buscar", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub cmdGuardar_Click()
On Error Resume Next

    SetearDatosCaja

If Err Then GrabarLog "cmdGuardar_Click", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub SetearDatosCaja()
On Error Resume Next

    If vModo = Modo.Seleccion Then
        If tieneSubconcepto(Val(txtCodigoConcepto.Text)) = True Then
            MsgBox "No se puede asignarle un valor a este concepto ya que tiene subconceptos asociados."
        Else
            Select Case vVieneConcepto
            
                Case "frmCaja"
                    With frmCaja
                        .txtConcepto.Text = txtConceptoVista.Text
                        .esIngreso = Me.OptIngreso.Value
                        .vCodigoConcepto = Val(Me.txtCodigoConcepto.Text)
                        Unload Me
                    End With
                        
                Case "frmArticulosAlta"
                    With frmArticulosAlta
                        .txtTecnica(5).Text = Val(txtCodigoConcepto.Text)
                        .txtTecnica(6).Text = txtConceptoVista.Text
                        Unload Me
                    End With
            End Select
        
        End If
    Else
        Unload Me
    End If

If Err Then GrabarLog "SetearDatosCaja", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Function CalcularSaldoAnterior() As Double
    Dim cmd As New ADODB.Command
    Dim sql As String
    cmd.ActiveConnection = ConnDDBB
    
    Dim sqlCliProv, sqlCli, sqlPro  As String
    Dim sqlFecha As String
    Dim sqlUsu As String
    Dim sqlCon As String
        
    If (chkCliente.Value = 1) And Not Trim(txtCliente.Text) = "" Then
        sqlCli = "cli.Nombre like '%" & Me.txtCliente.Text & "%'"
    Else
        sqlCli = 0
    End If
    
    If chkProv.Value = 1 And Not Trim(txtProveedor.Text) = "" Then
        sqlPro = "p.Nombre like '%" & Me.txtProveedor.Text & "%'"
    Else
        sqlPro = 0
    End If
    If Not sqlCli = 0 And Not sqlPro = 0 Then
        sqlCliProv = " and (" & sqlCli & " or " & sqlPro & ")"
    
    End If
    
    If chkFecha.Value = 1 Then
        sqlFecha = " and c.Fecha < '" & strfechaMySQL(dtDesde.Value) & "'"
    Else
        CalcularSaldoAnterior = 0
        Exit Function
    End If
       
    If Not Trim(txtUsuario.Text) = "" Then
        sqlUsu = " and c.Usuario like '%" & Me.txtUsuario.Text & "%'"
    End If
    
    If Not Trim(txtConceptoBuqueda.Text) = "" Then
        sqlCon = " and co.Descripcion like '%" & Trim(txtConceptoBuqueda.Text) & "%'"
    End If
    
    sql = "SELECT fecha, p.Nombre & cli.Nombre AS Nombre, usuario, Descripcion, IF(IngresoEgreso = True,  Importe, 0) AS Ingreso, IF(IngresoEgreso = False,  Importe, 0) AS Egreso, IF(IngresoEgreso = True,  Importe, - Importe) AS Saldo FROM ((caja AS c LEFT JOIN clientes AS cli ON c.CodigoCliente = cli.Codigo) LEFT JOIN Proveedores AS p ON c.CodigoProveedor = p.Codigo) LEFT JOIN Concepto AS co ON co.Codigo = c.CodigoConcepto"
    sql = sql & " Where 1 " & sqlCliProv & _
                 sqlUsu & sqlFecha & " and NroCheque = ''" & sqlCon
    

    Dim rsConsultaCaja As New ADODB.Recordset
    
    If rsConsultaCaja.State = 0 Then
        rsConsultaCaja.Open sql, ConnDDBB, 3, 3
    Else
        Set rsConsultaCaja = ConnDDBB.Execute(sql)
    End If
    
    Dim saldo As Double
    If Not rsConsultaCaja.EOF Then
        rsConsultaCaja.MoveFirst
    End If
    
    Do While Not rsConsultaCaja.EOF = True
        saldo = saldo + Val(Format(rsConsultaCaja("Saldo").Value, "#####0.00"))
        rsConsultaCaja.MoveNext
    Loop
    
    CalcularSaldoAnterior = saldo
    
End Function

Private Function ValidarFecha() As Boolean
On Error Resume Next

    ValidarFecha = True
    If dtDesde.Value > dtHasta.Value Then
        MsgBox "La fecha desde debe ser menor o igual a la fecha hasta"
        ValidarFecha = False
    End If

If Err Then GrabarLog "ValidarFecha", Err.Number & " " & Err.Description, Me.Name
End Function
Private Sub cmdImprimir_Click()
On Error Resume Next

If ValidarFecha() = True Then
    Dim cmd As New ADODB.Command
    Dim sql As String
    
    cmd.ActiveConnection = ConnDDBB
    
    Dim sqlCliProv, sqlCli, sqlPro  As String
    Dim sqlFecha As String
    Dim sqlUsu As String
    Dim sqlCon As String
    Dim SaldoAnterior As Double
    
    SaldoAnterior = CalcularSaldoAnterior()
        
    If (chkCliente.Value = 1) Then
        sqlCli = "cli.Nombre LIKE '%" & txtCliente.Text & "%'"
    Else
        sqlCli = " AND 1=2"
    End If
    
    If Not chkProv.Value = 1 And Not (txtProveedor.Text) = "" Then
        sqlPro = "p.Nombre like '%" & txtProveedor.Text & "%'"
    Else
        sqlPro = " AND 1=2"
    End If
    
    If chkFecha.Value = 1 Then
        sqlFecha = " and c.Fecha between '" & strfechaMySQL(Me.dtDesde.Value) & "' and '" & strfechaMySQL(Me.dtHasta.Value) & "'"
    End If
       
    If Not Trim(txtUsuario.Text) = "" Then
        sqlUsu = " and (c.Usuario LIKE '%" & Trim(txtUsuario.Text) & "%')"
    End If
    
    If Not Trim(txtConceptoBuqueda.Text) = "" Then
        sqlCon = " and (co.Descripcion like '%" & Trim(txtConceptoBuqueda.Text) & "%')"
    End If
    
    sql = "SELECT fecha, p.Nombre & cli.Nombre AS Nombre, usuario, Descripcion, c.Comentario, IF(IngresoEgreso = True,  Importe, 0) AS Ingreso, IF(IngresoEgreso = False,  Importe, 0) AS Egreso, IF(IngresoEgreso = True, Importe, - Importe) AS Saldo, " & SaldoAnterior & " as SaldoAnterior FROM ((caja AS c LEFT JOIN clientes AS cli ON c.CodigoCliente = cli.Codigo) LEFT JOIN Proveedores AS p ON c.CodigoProveedor = p.Codigo) LEFT JOIN Concepto AS co ON co.Codigo = c.CodigoConcepto"
    sql = sql & " Where 1 " & sqlCliProv & _
                 sqlUsu & sqlFecha & " and NroCheque = ''" & sqlCon
    
    Dim rsConsultaCaja As New ADODB.Recordset
    
    If rsConsultaCaja.State = 0 Then
        rsConsultaCaja.CursorLocation = adUseClient
        rsConsultaCaja.Open sql, ConnDDBB, 3, 3
    Else
        Set rsConsultaCaja = ConnDDBB.Execute(sql)
    End If

    Dim rsCajaImpresion As New ADODB.Recordset
      
    If rsCajaImpresion.State = 0 Then
        rsCajaImpresion.Open "CajaImpresion", ConnDDBB, 3, 3
    Else
        rsCajaImpresion.Close
        rsCajaImpresion.Open "CajaImpresion", ConnDDBB, 3, 3
    End If
    
    Dim i As Integer
    
    'Elimino los datos anteriores
    
    Call BorrarBase("CajaImpresion", pathDBMySQL)
    
    Dim saldoRenglonAnt As Double
    
    If rsConsultaCaja.RecordCount = 0 Then
        rsCajaImpresion.AddNew
        rsCajaImpresion("SaldoAnterior").Value = SaldoAnterior
        rsCajaImpresion.Update
    Else
    
    Do While Not rsConsultaCaja.EOF
        
        rsCajaImpresion.AddNew
        For i = 0 To 8
            
            If Not IsNull(rsConsultaCaja(i).Value) = True Then 'Nueva Reforma
                rsCajaImpresion(i).Value = rsConsultaCaja(i).Value
            End If
        
        Next i
        
        
        rsCajaImpresion("Saldo").Value = saldoRenglonAnt + Val(Format(rsConsultaCaja("Saldo").Value, "#######0.00"))
        
        
        saldoRenglonAnt = Val(Format(rsCajaImpresion("Saldo").Value, "#######0.00"))
        rsCajaImpresion.Update
        rsConsultaCaja.MoveNext
    Loop
    End If
    rsConsultaCaja.Close
    rsCajaImpresion.Close
    
    Unload Mantenimiento
    Load Mantenimiento
    
    With drcaja
        '.Sections(2).Controls("lblSaldoAnterior").Caption = CalcularSaldoAnterior(fdesde.Value)
        
        If chkFecha.Value = 0 Then
            .Sections(2).Controls("gFechaDesde").Caption = "Todas"
            .Sections(2).Controls("gFechaHasta").Caption = "Todas"
            .Refresh
        Else
            .Sections(2).Controls("gFechaDesde").Caption = Me.dtDesde.Value
            .Sections(2).Controls("gFechaHasta").Caption = Me.dtHasta.Value
            .Refresh
        End If
        
        If chkCliente.Value = 1 Then
            If Me.txtCliente.Text <> "" Then
                .Sections(2).Controls("gCliente").Caption = Me.txtCliente.Text
            Else
                .Sections(2).Controls("gCliente").Caption = "Todos"
            End If
        Else
            .Sections(2).Controls("gCliente").Caption = "Ninguno"
        End If
        
        If chkProv.Value = 1 Then
            If Me.txtProveedor.Text <> "" Then
                .Sections(2).Controls("gProveedor").Caption = Me.txtProveedor.Text
            Else
                .Sections(2).Controls("gProveedor").Caption = "Todos"
            End If
        Else
            .Sections(2).Controls("gProveedor").Caption = "Ninguno"
        End If
        
        If txtUsuario.Text <> "" Then
            .Sections(2).Controls("gUsuario").Caption = Me.txtUsuario.Text
        Else
            .Sections(2).Controls("gUsuario").Caption = "Todos"
        End If
        
        If txtConcepto.Text <> "" Then
            .Sections(2).Controls("gConcepto").Caption = Me.txtConcepto.Text
        Else
            .Sections(2).Controls("gConcepto").Caption = "Todos"
        End If
        
        .Refresh
        
        .Show
    
    End With
End If

If Err Then GrabarLog "cmdImprimir_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub cmdRenombrar_Click()
    Dim tNodo As Node
   
    ' El nodo que está actualmente seleccionado
    Set tNodo = TreeView1.SelectedItem
   
    ' Sustituir el texto del nodo seleccionado
    ' por el contenido del cboConcepto
    TreeView1.SelectedItem.Text = txtConcepto.Text
    
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
      
    sqlConcepto = "Select * from Concepto where Codigo = " & Mid(tNodo.Key, 3, ContarCaracteres(tNodo.Key, "-") - 2)
    
    Dim rsConcepto As New ADODB.Recordset
      
    If rsConcepto.State = 0 Then
        rsConcepto.CursorLocation = adUseClient
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
    
    If rsConcepto.RecordCount = 1 Then
        rsConcepto.MoveFirst
        rsConcepto("Descripcion").Value = txtConcepto.Text
        rsConcepto.Update
    End If
End Sub

Private Sub cmdBorrarSeleccionado_Click()
    BorrarNodo
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub chkCliente_Click()
If Me.chkCliente.Value = 1 Then
     Me.txtCliente.Enabled = True
     txtCliente.BackColor = &H80000005
Else
    txtCliente.Enabled = False
    txtCliente.Text = ""
    txtCliente.BackColor = &H8000000F
End If
End Sub

Private Sub chkFecha_Click()
If Me.chkFecha.Value = 1 Then
    Me.dtDesde.Enabled = True
    Me.dtHasta.Enabled = True
Else
    Me.dtDesde.Enabled = False
    Me.dtHasta.Enabled = False
End If
End Sub

Private Sub chkProv_Click()
If Me.chkProv.Value = 1 Then
     Me.txtProveedor.Enabled = True
     txtProveedor.BackColor = &H80000005
Else
    txtProveedor.Enabled = False
    txtProveedor.Text = ""
    txtProveedor.BackColor = &H8000000F
End If
End Sub
Private Sub Form_Activate()
Me.dtDesde.Value = Date
Me.dtHasta.Value = Date

Actualizar
Dim vSaldoAnterior As Double
vSaldoAnterior = CalcularSaldoAnterior
Me.lblSaldoAnterior = vSaldoAnterior
 If vSaldoAnterior >= 0 Then
        Me.lblSaldoAnterior.ForeColor = &HC000&
Else
    Me.lblSaldoAnterior.ForeColor = &HFF&
End If

End Sub

Public Sub Actualizar()
If leido = False Then
    MousePointer = vbHourglass
    ActivarBotones False
    Me.fraBusqueda.Enabled = False = False
    Me.Top = 0
    CambiarStatus "Cargando", "Inicializando arbol...", App.Path & "\iconos\im29.ico"
    InicializarLlenadoArbol
    leido = True
    CambiarStatus "WSF", "WSF - Sistema integral de caja diaria", App.Path & "\iconos\144.ico"
    MostrarControles
End If
End Sub

Private Sub MostrarControles()
    If vModo = Modo.Creacion Then
        fraBusqueda.Visible = False
        fraOpera.Visible = True
    End If
    
    If vModo = Modo.Lectura Then
        fraBusqueda.Visible = True
        fraOpera.Visible = False
    End If
    
    If vModo = Modo.Seleccion Then
        fraBusqueda.Visible = False
        fraOpera.Visible = False
    End If
End Sub

Private Sub ActualizarDespuesOperar()
    MousePointer = vbHourglass
    ActivarBotones False
    Me.fraBusqueda.Enabled = False = False
    Me.Top = 0
    CambiarStatus "Cargando", "Inicializando arbol...", App.Path & "\Imagenes\iconos\im29.ico"
    InicializarLlenadoArbol
    leido = True
    CambiarStatus "WSF", "WSF - Sistema integral de caja diaria", App.Path & "\Imagenes\iconos\144.ico"
    MostrarControles
End Sub
Private Sub ActivarBotones(v As Boolean)
On Error Resume Next

cmdAñadir.Enabled = v
cmdBorrar.Enabled = v
cmdBuscar.Enabled = v
cmdCerrar.Enabled = v
cmdGuardar.Enabled = v
cmdRenombrar.Enabled = v
cmdImprimir.Enabled = v

If Err Then GrabarLog "ActivarBotones", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub CambiarStatus(Nombre As String, descripcion As String, imPath As String)
On Error Resume Next

    With BarraEstado
        .Panels.Clear
        .Panels.Add , "Hora", , sbrTime
        .Panels.Add , "Fecha", , sbrDate

        .Panels.Add , Nombre, descripcion, sbrText, LoadPicture(imPath)
        .Panels(1).AutoSize = sbrContents
        .Panels(2).AutoSize = sbrContents
        .Panels(3).width = 10000
        '.Panels(3).AutoSize = sbrSpring
    
    End With

If Err Then GrabarLog "CambiarStatus", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub InicializarLlenadoArbol()

    Me.ProgressBar1.Min = 0
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Refresh

    ' Configuramos manualmente el Treeview
    With TreeView1
        .Style = tvwTreelinesPlusMinusText
        .LineStyle = tvwRootLines
        .PathSeparator = "\"
        .Indentation = Screen.TwipsPerPixelX * 5 '256
        '
        ' No permitir la edición automática del texto
        .LabelEdit = tvwManual
        ' Para que se pueda expandir al seleccionar un nodo,
        ' cambia este valor a True,
        ' si se deja en False, se expande al hacer doble-click
        .SingleSel = False
        ' Para que al perder el foco,
        ' se siga viendo el que está seleccionado
        .HideSelection = False
        '
        .Refresh
    End With
    '
    PrepararImageList
    ' Llenar el Treeview con los nodos de la tabla Concepto
    If mostrarSaldos = True Then
        LlenarArbol
        DoEvents
        Calculado = True
        'Me.TreeView1.Nodes.Remove 61
        Me.lblIngresos.Visible = True
        Me.lblEgresos.Visible = True
        Me.lblSal.Visible = True
        Me.lblSaldo.Visible = True
        Me.lblSalAnt.Visible = True
        Me.lblSaldoAnterior.Visible = True
        Me.fraBusqueda.Visible = True
    Else
        Me.lblIngresos.Visible = False
        Me.lblEgresos.Visible = False
        Me.lblSal.Visible = False
        Me.lblSaldo.Visible = False
        Me.fraBusqueda.Visible = False
        Me.lblSalAnt.Visible = False
        Me.lblSaldoAnterior.Visible = False
    End If
    LlenarArbol
    saldototal = 0
  
    If Me.lblIngresoValue.Caption <> "" And Me.lblEgresoValue.Caption <> "" Then
        saldototal = CDbl(Me.lblIngresoValue.Caption) - CDbl(Me.lblEgresoValue.Caption)
    End If
    Me.lblSaldo.Caption = Format(saldototal, "#######0.00")
    If saldototal > 0 Then
        Me.lblSaldo.ForeColor = &HC000&
    Else
        Me.lblSaldo.ForeColor = &HFF&
    End If
    Me.height = 9570
    Me.width = 12780
    Me.ProgressBar1.Value = Me.ProgressBar1.Min
    MousePointer = Default
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mostrarSaldos = False
    Calculado = False
    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

End Sub
Private Sub TreeView1_DblClick()
    SetearDatosCaja
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim s As String
    s = Node.Text
    If Node.Children > 0 Then
        s = s & ", tiene " & Node.Children & " hijos"
    Else
        s = s & ", no tiene hijos"
    End If
        
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    cmdConcepto.ActiveConnection = ConnDDBB
      
    sqlConcepto = "SELECT * FROM Concepto WHERE (Codigo = " & Mid(Node.Key, 3, ContarCaracteres(Node.Key, "-") - 2) & ")"
    
    Dim rsConcepto As New ADODB.Recordset
      
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
      
    If Not rsConcepto.EOF Then
        rsConcepto.MoveFirst
        Me.txtConceptoVista.Text = rsConcepto("Descripcion").Value
        Me.txtConcepto.Text = rsConcepto("Descripcion").Value
        Me.txtCodigoConcepto.Text = rsConcepto("Codigo").Value
        If CBool(rsConcepto("IngresoEgreso").Value) = True Then
            Me.OptIngreso.Value = True
        Else
            Me.OptEgreso.Value = True
        End If
        If CBool(rsConcepto("TieneSubconcepto").Value) = True Then
            Me.chkTieneSubconcepto.Value = 1
        Else
            Me.chkTieneSubconcepto.Value = 0
        End If
    End If
End Sub

Private Sub PrepararImageList()
   Set Me.TreeView1.ImageList = ImageList2
End Sub



Private Function CalcularTotalConcepto(Codigo As Long, esIngreso As Boolean)
    Dim total As Double
    Dim cmdCaja As New ADODB.Command
    
    Dim sqlCliProv, sqlCli, sqlPro  As String
    Dim sqlFecha As String
    cmdCaja.ActiveConnection = ConnDDBB
    
    If Me.chkCliente.Value = 1 Then
        sqlCli = "cli.Nombre like '%" & Me.txtCliente.Text & "%'"
    Else
        sqlCli = 0
    End If
    
    If Me.chkProv.Value = 1 Then
        sqlPro = "p.Nombre like '%" & Me.txtProveedor.Text & "%'"
    Else
        sqlPro = 0
    End If
    
    sqlCliProv = " and (" & sqlCli & " or " & sqlPro & ")"
    
    If Me.chkFecha.Value = 1 Then
        sqlFecha = " and ca.Fecha between ''" & Me.dtDesde.Value & "'' and ''" & Me.dtHasta.Value & "''"
    End If
    
    sqlCaja = "Select * from ((Caja ca  inner join Proveedores p on ca.CodigoProveedor = p.Codigo) " & _
                 " inner join Clientes cli on ca.CodigoCliente = cli.Codigo)" & _
                 " Where CodigoConcepto = " & Str(Codigo) & _
                  sqlCliProv & _
                 " and ca.Usuario like '%" & Me.txtUsuario.Text & "%'" & sqlFecha

    
        
    Dim rsCaja As New ADODB.Recordset
      
    If rsCaja.State = 0 Then
        rsCaja.Open sqlCaja, ConnDDBB, 3, 3
    Else
        Set rsCaja = ConnDDBB.Execute(sqlCaja)
    End If
      
    If Not rsCaja.EOF Then
        rsCaja.MoveFirst
    End If
    
    Do While Not rsCaja.EOF
        total = total + rsCaja("Importe").Value
        rsCaja.MoveNext
    Loop
    
    CalcularTotalConcepto = total
End Function
Private Function CalcularTotalConceptoRec(Codigo As Long, tieneSubconcepto As Boolean, esIngreso As Boolean)
    Dim total As Double
    Dim cmdCaja As New ADODB.Command
    Dim sqlCaja As String
    Dim sqlCliProv, sqlUsuario, sqlFecha  As String
    cmdCaja.ActiveConnection = ConnDDBB
    
    If esIngreso = True Then
        If chkCliente.Value = 1 Then
            sqlCliProv = " AND (cli.Nombre like '%" & Trim(txtCliente.Text) & "%')"
        Else
            sqlCliProv = " AND 2=1"
        End If
    Else
        If chkProv.Value = 1 Then
            sqlCliProv = " AND (p.Nombre like '%" & txtProveedor.Text & "%')"
        Else
            sqlCliProv = " AND 2=1"
        End If
    End If
    
    If chkFecha.Value = 1 Then
        sqlFecha = " AND (ca.Fecha between '" & strfechaMySQL(dtDesde.Value) & "' and '" & strfechaMySQL(dtHasta.Value) & "')"
    End If
    
    If Not Trim(txtUsuario.Text) = "" Then
        sqlUsuario = " and (ca.Usuario like '%" & Trim(txtUsuario.Text) & "%')"
    End If
    
    sqlCaja = "SELECT * FROM ((Caja ca left join Proveedores p on ca.CodigoProveedor = p.Codigo) " & _
                 " LEFT JOIN Clientes cli on ca.CodigoCliente = cli.Codigo)" & _
                 " WHERE 1=1 AND (CodigoConcepto = " & Str(Codigo) & ") " & _
                  sqlCliProv & sqlUsuario & sqlFecha & " AND (NroCheque = '' OR NroCheque Is Null OR NroCheque = 0)"
    
    Dim rsCaja As New ADODB.Recordset
      
    If rsCaja.State = 0 Then
        rsCaja.Open sqlCaja, ConnDDBB, 3, 3
    Else
        Set rsCaja = ConnDDBB.Execute(sqlCaja)
    End If
      
    If Not rsCaja.EOF Then rsCaja.MoveFirst
    
    If tieneSubconcepto = True Then
        'Busco los hijos del concepto
        Dim cmdConc As New ADODB.Command
        Dim sqlConc As String
        cmdConc.ActiveConnection = ConnDDBB
          
        sqlConc = "SELECT * FROM Concepto WHERE (CodigoPadre = " & Trim(Codigo) & ")"
        
        Dim rsConc As New ADODB.Recordset
          
        If rsConc.State = 0 Then
            rsConc.Open sqlConc, ConnDDBB, 3, 3
        Else
            Set rsConc = ConnDDBB.Execute(sqlConc)
        End If
          
        If Not rsConc.EOF Then
            rsConc.MoveFirst
        End If
        
        Do While Not rsConc.EOF
            CalcularTotalConceptoRec = CalcularTotalConceptoRec + CalcularTotalConceptoRec(rsConc("Codigo"), rsConc("TieneSubconcepto"), rsConc("IngresoEgreso"))
            rsConc.MoveNext
        Loop
        
    Else
        Do While Not rsCaja.EOF = True
            total = total + Val(Format(rsCaja("Importe").Value, "#####0.00"))
            
            rsCaja.MoveNext
        Loop
        CalcularTotalConceptoRec = total
    End If
    
End Function
Private Function tieneSubconcepto(Codigo As Long) As Boolean
    Dim cmdConcepto As New ADODB.Command
    Dim sqlConcepto As String
    
    cmdConcepto.ActiveConnection = ConnDDBB
      
    sqlConcepto = "Select * from Concepto WHERE (Codigo = " & Str(Codigo) & ")"
    
    Dim rsConcepto As New ADODB.Recordset
      
    If rsConcepto.State = 0 Then
        rsConcepto.Open sqlConcepto, ConnDDBB, 3, 3
    Else
        Set rsConcepto = ConnDDBB.Execute(sqlConcepto)
    End If
      
    If Not rsConcepto.EOF Then
        rsConcepto.MoveFirst
        tieneSubconcepto = rsConcepto("TieneSubconcepto").Value
    End If
    
End Function
