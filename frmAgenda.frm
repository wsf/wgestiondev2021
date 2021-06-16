VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAgenda 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agenda"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9150
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin MSAdodcLib.Adodc bconceptos 
      Height          =   375
      Left            =   0
      Top             =   7800
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
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
      Caption         =   "bconceptos"
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
   Begin MSAdodcLib.Adodc bagenda 
      Height          =   375
      Left            =   0
      Top             =   7440
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
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
      Caption         =   "bagenda"
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
   Begin TabDlg.SSTab tab_agenda 
      Height          =   6675
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   11774
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Ingreso"
      TabPicture(0)   =   "frmAgenda.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ImageList2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "tab_alta"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Mantenimiento"
      TabPicture(1)   =   "frmAgenda.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cftarea"
      Tab(1).Control(1)=   "citarea"
      Tab(1).Control(2)=   "dgagenda"
      Tab(1).Control(3)=   "Frame10"
      Tab(1).Control(4)=   "busqueda"
      Tab(1).Control(5)=   "Line2"
      Tab(1).Control(6)=   "lblAgenda"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Modo Calendario"
      TabPicture(2)   =   "frmAgenda.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "calendario"
      Tab(2).Control(2)=   "gdcalendario"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Finalizadas"
      TabPicture(3)   =   "frmAgenda.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Line1"
      Tab(3).Control(1)=   "Label8"
      Tab(3).Control(2)=   "DataGrid1"
      Tab(3).ControlCount=   3
      Begin VB.CommandButton cftarea 
         Caption         =   "Finalizar Tarea"
         Height          =   255
         Left            =   -67440
         TabIndex        =   21
         Top             =   5760
         Width           =   1215
      End
      Begin VB.CommandButton citarea 
         Caption         =   "Iniciar Tarea"
         Height          =   255
         Left            =   -68640
         TabIndex        =   56
         Top             =   5760
         Width           =   1215
      End
      Begin TabDlg.SSTab tab_alta 
         Height          =   5415
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   9551
         _Version        =   393216
         TabOrientation  =   1
         Style           =   1
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "Tarea"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblHora"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblfecha"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label6"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblComentario"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "vfecha"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "vhora"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "vperiodo"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "fraagenda"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "vtarea"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Notas"
         TabPicture(1)   =   "frmAgenda.frx":0070
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "lblNotaPara"
         Tab(1).Control(1)=   "Label2"
         Tab(1).Control(2)=   "Linea_nota"
         Tab(1).Control(3)=   "Label17"
         Tab(1).Control(4)=   "cbonota"
         Tab(1).Control(5)=   "vnota"
         Tab(1).ControlCount=   6
         Begin VB.TextBox vtarea 
            Height          =   2055
            Left            =   1680
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   840
            Width           =   6615
         End
         Begin VB.TextBox vnota 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3735
            Left            =   -74880
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   720
            Width           =   8295
         End
         Begin VB.ComboBox cbonota 
            Height          =   315
            ItemData        =   "frmAgenda.frx":008C
            Left            =   -73800
            List            =   "frmAgenda.frx":009C
            TabIndex        =   43
            Top             =   4560
            Width           =   7215
         End
         Begin VB.Frame fraagenda 
            Caption         =   " Programar Pagos"
            Height          =   1935
            Left            =   240
            TabIndex        =   36
            Top             =   2880
            Width           =   8145
            Begin VB.ComboBox vconcepto 
               Height          =   315
               ItemData        =   "frmAgenda.frx":00BF
               Left            =   1320
               List            =   "frmAgenda.frx":00C9
               TabIndex        =   47
               Top             =   720
               Width           =   6615
            End
            Begin VB.TextBox vmonto 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1320
               TabIndex        =   39
               Top             =   1080
               Width           =   6615
            End
            Begin VB.TextBox vcliente 
               Height          =   315
               Left            =   1320
               TabIndex        =   38
               Top             =   360
               Width           =   6615
            End
            Begin VB.TextBox vtotal 
               Alignment       =   1  'Right Justify
               Height          =   315
               Left            =   1320
               TabIndex        =   37
               Top             =   1440
               Width           =   6615
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   ">Concepto :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   120
               TabIndex        =   48
               Top             =   720
               Width           =   1065
            End
            Begin VB.Label lblCredito 
               AutoSize        =   -1  'True
               Caption         =   "> Monto :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   360
               TabIndex        =   42
               Top             =   1080
               Width           =   795
            End
            Begin VB.Label lblCliente 
               AutoSize        =   -1  'True
               Caption         =   "> Cliente :"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   360
               TabIndex        =   41
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "> Total :"
               Height          =   195
               Left            =   600
               TabIndex        =   40
               Top             =   1440
               Width           =   585
            End
         End
         Begin VB.ComboBox vperiodo 
            Height          =   315
            ItemData        =   "frmAgenda.frx":00DF
            Left            =   1680
            List            =   "frmAgenda.frx":00EF
            TabIndex        =   34
            Top             =   480
            Width           =   6615
         End
         Begin MSComCtl2.DTPicker vhora 
            Height          =   315
            Left            =   5040
            TabIndex        =   32
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   125370370
            CurrentDate     =   38790
         End
         Begin MSComCtl2.DTPicker vfecha 
            Height          =   315
            Left            =   1680
            TabIndex        =   33
            Top             =   120
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Format          =   125370369
            CurrentDate     =   38790
         End
         Begin VB.Label Label17 
            Caption         =   "Escriba texto ""libre""..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008080&
            Height          =   195
            Left            =   -75000
            TabIndex        =   54
            Top             =   360
            Width           =   2385
         End
         Begin VB.Line Linea_nota 
            BorderColor     =   &H0000C000&
            X1              =   -75000
            X2              =   -66120
            Y1              =   240
            Y2              =   240
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "Notas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   -75000
            TabIndex        =   46
            Top             =   0
            Width           =   8895
         End
         Begin VB.Label lblNotaPara 
            AutoSize        =   -1  'True
            Caption         =   "Nota para :"
            Height          =   195
            Left            =   -74640
            TabIndex        =   45
            Top             =   4560
            Width           =   795
         End
         Begin VB.Label lblComentario 
            AutoSize        =   -1  'True
            Caption         =   "> Periodicidad :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   1395
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "> Tarea :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   720
            TabIndex        =   30
            Top             =   840
            Width           =   795
         End
         Begin VB.Label lblfecha 
            AutoSize        =   -1  'True
            Caption         =   "> Fecha :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   720
            TabIndex        =   29
            Top             =   120
            Width           =   810
         End
         Begin VB.Label lblHora 
            AutoSize        =   -1  'True
            Caption         =   "> Hora :"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3960
            TabIndex        =   28
            Top             =   120
            Width           =   690
         End
      End
      Begin VB.Frame Frame1 
         Height          =   555
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   8925
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   660
            Left            =   90
            TabIndex        =   18
            Top             =   150
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   1164
            ButtonWidth     =   1931
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList1"
            HotImageList    =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "Borrar "
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Guardar "
                  Object.ToolTipText     =   "<F2> Guarda "
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "Buscar"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "Imprimir"
                  ImageIndex      =   5
               EndProperty
            EndProperty
         End
         Begin VB.Label lblMantenimientoDe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[Mantenimiento de Agenda] "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   6120
            TabIndex        =   19
            Top             =   240
            Width           =   2400
         End
      End
      Begin MSDataGridLib.DataGrid gdcalendario 
         Bindings        =   "frmAgenda.frx":011B
         Height          =   3345
         Left            =   -74880
         TabIndex        =   16
         Top             =   2790
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   5900
         _Version        =   393216
         BackColor       =   -2147483628
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column02 
            DataField       =   "Hora"
            Caption         =   "Hora"
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
         BeginProperty Column03 
            DataField       =   "Tarea"
            Caption         =   "Tarea"
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
         BeginProperty Column04 
            DataField       =   "Notas"
            Caption         =   "Notas"
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
         BeginProperty Column05 
            DataField       =   "Periodo"
            Caption         =   "Periodo"
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
         BeginProperty Column06 
            DataField       =   "concepto"
            Caption         =   "concepto"
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
         BeginProperty Column07 
            DataField       =   "codigo_cli"
            Caption         =   "codigo_cli"
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
         BeginProperty Column08 
            DataField       =   "codigo_ctacte"
            Caption         =   "codigo_ctacte"
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
         BeginProperty Column09 
            DataField       =   "codigo_caja"
            Caption         =   "codigo_caja"
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
         BeginProperty Column10 
            DataField       =   "Estado"
            Caption         =   "Estado"
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
         BeginProperty Column11 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
         BeginProperty Column12 
            DataField       =   "usuario_nota"
            Caption         =   "usuario_nota"
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
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   3690.142
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   14.74
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
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid dgagenda 
         Bindings        =   "frmAgenda.frx":0131
         Height          =   3735
         Left            =   -75000
         TabIndex        =   15
         Top             =   1920
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   6588
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column02 
            DataField       =   "Hora"
            Caption         =   "Hora"
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
         BeginProperty Column03 
            DataField       =   "Tarea"
            Caption         =   "Tarea"
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
         BeginProperty Column04 
            DataField       =   "Notas"
            Caption         =   "Notas"
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
         BeginProperty Column05 
            DataField       =   "Periodo"
            Caption         =   "Periodo"
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
         BeginProperty Column06 
            DataField       =   "concepto"
            Caption         =   "concepto"
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
         BeginProperty Column07 
            DataField       =   "codigo_cli"
            Caption         =   "codigo_cli"
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
         BeginProperty Column08 
            DataField       =   "codigo_ctacte"
            Caption         =   "codigo_ctacte"
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
         BeginProperty Column09 
            DataField       =   "codigo_caja"
            Caption         =   "codigo_caja"
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
         BeginProperty Column10 
            DataField       =   "Estado"
            Caption         =   "Estado"
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
         BeginProperty Column11 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
         BeginProperty Column12 
            DataField       =   "usuario_nota"
            Caption         =   "usuario_nota"
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
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   4800.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   14.74
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
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame10 
         Height          =   555
         Left            =   -75000
         TabIndex        =   2
         Top             =   0
         Width           =   8925
         Begin MSComctlLib.Toolbar menu 
            Height          =   330
            Left            =   0
            TabIndex        =   4
            Top             =   120
            Width           =   5205
            _ExtentX        =   9181
            _ExtentY        =   582
            ButtonWidth     =   1826
            ButtonHeight    =   582
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            DisabledImageList=   "ImageList1"
            HotImageList    =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   5
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Agregar"
                  ImageIndex      =   1
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Borrar "
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "Guardar "
                  Object.ToolTipText     =   "<F2> Guarda "
                  ImageIndex      =   3
               EndProperty
               BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Buscar"
                  ImageIndex      =   4
               EndProperty
               BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "Imprimir"
                  ImageIndex      =   5
               EndProperty
            EndProperty
         End
         Begin VB.Label lblLabel2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "[Mantenimiento de Agenda]"
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   6840
            TabIndex        =   3
            Top             =   240
            Width           =   1950
         End
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   10080
         Top             =   -300
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":0147
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":0259
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":036B
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":047D
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":058F
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList ImageList2 
         Left            =   9120
         Top             =   7380
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483633
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":06A1
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":07B3
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":08C5
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":09D7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmAgenda.frx":0AE9
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab busqueda 
         Height          =   975
         Left            =   -75000
         TabIndex        =   5
         Top             =   600
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   1720
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Por Tarea"
         TabPicture(0)   =   "frmAgenda.frx":0BFB
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblcaja"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lblconcepto"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "vctarea"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Por Nota"
         TabPicture(1)   =   "frmAgenda.frx":0C17
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label1"
         Tab(1).Control(1)=   "vcnota"
         Tab(1).ControlCount=   2
         TabCaption(2)   =   "Por Estado"
         TabPicture(2)   =   "frmAgenda.frx":0C33
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Label4"
         Tab(2).Control(1)=   "vcestado"
         Tab(2).ControlCount=   2
         TabCaption(3)   =   "Por Fecha"
         TabPicture(3)   =   "frmAgenda.frx":0C4F
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "Label9"
         Tab(3).Control(1)=   "Label10"
         Tab(3).Control(2)=   "vchasta"
         Tab(3).Control(3)=   "vcdesde"
         Tab(3).Control(4)=   "f"
         Tab(3).ControlCount=   5
         Begin VB.ComboBox vcestado 
            Height          =   315
            ItemData        =   "frmAgenda.frx":0C6B
            Left            =   -73200
            List            =   "frmAgenda.frx":0C78
            TabIndex        =   55
            Top             =   480
            Width           =   6975
         End
         Begin VB.CheckBox f 
            Caption         =   "Anular Fechas"
            Height          =   255
            Left            =   -74760
            TabIndex        =   53
            Top             =   480
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker vcdesde 
            Height          =   315
            Left            =   -71880
            TabIndex        =   25
            Top             =   480
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   124780545
            CurrentDate     =   38798
         End
         Begin VB.TextBox vcnota 
            Height          =   315
            Left            =   -73080
            TabIndex        =   23
            Top             =   480
            Width           =   6855
         End
         Begin VB.TextBox vctarea 
            Height          =   315
            Left            =   2040
            TabIndex        =   22
            Top             =   480
            Width           =   6735
         End
         Begin VB.TextBox vccliente 
            Height          =   315
            Left            =   -73080
            TabIndex        =   7
            Top             =   525
            Width           =   6735
         End
         Begin VB.TextBox vccomentario 
            Height          =   315
            Left            =   -72840
            TabIndex        =   6
            Top             =   525
            Width           =   6495
         End
         Begin MSComCtl2.DTPicker fhasta 
            Height          =   315
            Left            =   -68520
            TabIndex        =   8
            Top             =   525
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
            Format          =   124780545
            CurrentDate     =   38028
         End
         Begin MSComCtl2.DTPicker fdesde 
            Height          =   315
            Left            =   -71400
            TabIndex        =   9
            Top             =   525
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
            Format          =   124780545
            CurrentDate     =   38028
         End
         Begin MSComCtl2.DTPicker vchasta 
            Height          =   315
            Left            =   -68040
            TabIndex        =   26
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   124780545
            CurrentDate     =   38798
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde :"
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
            Left            =   -73200
            TabIndex        =   52
            Top             =   510
            Width           =   1260
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta :"
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
            Left            =   -69480
            TabIndex        =   51
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "> Ingrese Estado :"
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
            Left            =   -74880
            TabIndex        =   24
            Top             =   555
            Width           =   1575
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "> Ingrese Nota :"
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
            Left            =   -74880
            TabIndex        =   20
            Top             =   555
            Width           =   1395
         End
         Begin VB.Label lblIngreseComentario 
            AutoSize        =   -1  'True
            Caption         =   "> Ingrese Comentario :"
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
            Left            =   -74880
            TabIndex        =   14
            Top             =   555
            Width           =   2040
         End
         Begin VB.Label lblconcepto 
            AutoSize        =   -1  'True
            Caption         =   "> Ingrese Tarea:"
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
            Left            =   120
            TabIndex        =   13
            Top             =   555
            Width           =   1425
         End
         Begin VB.Label lblcaja 
            AutoSize        =   -1  'True
            Caption         =   "- Condiciones de Busqueda de Agenda -"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004040&
            Height          =   240
            Left            =   4680
            TabIndex        =   12
            Top             =   0
            Width           =   4215
         End
         Begin VB.Label lblFechaDesde 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Desde :"
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
            Left            =   -72720
            TabIndex        =   11
            Top             =   550
            Width           =   1260
         End
         Begin VB.Label lblFechaHasta 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Hasta :"
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
            Left            =   -69840
            TabIndex        =   10
            Top             =   550
            Width           =   1260
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmAgenda.frx":0C9F
         Height          =   5535
         Left            =   -75000
         TabIndex        =   49
         Top             =   480
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9763
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   13
         BeginProperty Column00 
            DataField       =   "id"
            Caption         =   "id"
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
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
         BeginProperty Column02 
            DataField       =   "Hora"
            Caption         =   "Hora"
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
         BeginProperty Column03 
            DataField       =   "Tarea"
            Caption         =   "Tarea"
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
         BeginProperty Column04 
            DataField       =   "Notas"
            Caption         =   "Notas"
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
         BeginProperty Column05 
            DataField       =   "Periodo"
            Caption         =   "Periodo"
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
         BeginProperty Column06 
            DataField       =   "concepto"
            Caption         =   "concepto"
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
         BeginProperty Column07 
            DataField       =   "codigo_cli"
            Caption         =   "codigo_cli"
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
         BeginProperty Column08 
            DataField       =   "codigo_ctacte"
            Caption         =   "codigo_ctacte"
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
         BeginProperty Column09 
            DataField       =   "codigo_caja"
            Caption         =   "codigo_caja"
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
         BeginProperty Column10 
            DataField       =   "Estado"
            Caption         =   "Estado"
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
         BeginProperty Column11 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
         BeginProperty Column12 
            DataField       =   "usuario_nota"
            Caption         =   "usuario_nota"
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
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   14.74
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   2069.858
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1409.953
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   14.74
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
               ColumnWidth     =   1440
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   14.74
            EndProperty
         EndProperty
      End
      Begin MSACAL.Calendar calendario 
         Height          =   2505
         Left            =   -74940
         TabIndex        =   57
         Top             =   390
         Width           =   8805
         _Version        =   524288
         _ExtentX        =   15531
         _ExtentY        =   4419
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2006
         Month           =   3
         Day             =   22
         DayLength       =   0
         MonthLength     =   2
         DayFontColor    =   0
         FirstDay        =   1
         GridCellEffect  =   0
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Agenda de eventos - Modo Calendario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74610
         TabIndex        =   58
         Top             =   60
         Width           =   8385
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "Finalizada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -75000
         TabIndex        =   50
         Top             =   0
         Width           =   8895
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000F&
         BorderWidth     =   2
         X1              =   -75000
         X2              =   -66120
         Y1              =   260
         Y2              =   260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000004&
         X1              =   -75000
         X2              =   -66120
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Label lblAgenda 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         Caption         =   "Agenda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000006&
         Height          =   255
         Left            =   -75000
         TabIndex        =   1
         Top             =   1560
         Width           =   8895
      End
   End
   Begin MSAdodcLib.Adodc bpersonas 
      Height          =   375
      Left            =   0
      Top             =   8160
      Visible         =   0   'False
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
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
      Caption         =   "bpersonas"
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
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vtab As String
Dim vmodifica_agenda As Integer
Dim vid As Long
Dim vcod_cliente As String



Private Sub buscacli()
    
    With bpersonas
   
        .RecordSource = "select * from Clientes where (nombre = '" + Trim(vcliente) + "') or (codigo = '" + Trim(vcliente) + "')"
        .Refresh

        If .Recordset.EOF Then
            frmBuscarCliente.Show
            frmBuscarCliente.o = 4
            frmBuscarCliente.txtClientes.Text = vcliente.Text
            frmBuscarCliente.txtClientes.SetFocus
        Else
        
            vcliente = .Recordset("Nombre").Value
            vcod_cliente = .Recordset("codigo").Value
        End If
    
    End With

End Sub

'prueba check in

Private Sub Buscar()
    On Error Resume Next
    
    Dim sql As String

    If Not vctarea.Text = "" Then sql = sql + " and tarea like '%" + Trim(vctarea.Text) + "%'"
    If Not vcnota.Text = "" Then sql = sql + " and Notas like '%" + Trim(vcnota.Text) + "%'"
    If Not vcestado.Text = "" Then sql = sql + " and estado = '" + Trim(vcestado.Text) + "'"
    If Not f.Value = 1 Then sql = sql + " and (fecha >= '" & strfechaMySQL(vcdesde) & "' and fecha <= '" & strfechaMySQL(vchasta) & "')"

    bagenda.RecordSource = "Select * from agenda where 1=1" + sql
    bagenda.Refresh

    If Err Then GrabarLog "Buscar", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub calendario_Click()
    Dim sql As String

    If calendario.Value > Date Then
        gdcalendario.BackColor = &HFFFFFF       'blanco
    End If
    
    If calendario.Value < Date Then
        gdcalendario.BackColor = &HFFFFFF 'blanco
    End If
     
    If calendario.Value = Date Then
        gdcalendario.BackColor = &H80000018 'amarillo
    End If
    
    bagenda.RecordSource = "Select * from agenda where fecha = '" + Str(calendario.Value) + "'"
    bagenda.Refresh
    
End Sub

Private Sub cftarea_Click()

    'On Error Resume Next
    If Not MsgBox("¿Desea dar por terminada esta tarea?", vbYesNo + vbInformation, "Pregunta...") = vbYes Then Exit Sub
    
    If bagenda.Recordset.RecordCount = 0 Then Exit Sub
    
    bagenda.Recordset("Estado") = "Completada"
    bagenda.Recordset.Update
    
    bagenda.RecordSource = "select * from agenda where not Estado = 'Completada'"
    bagenda.Refresh
    
    Me.tab_agenda.tab = 3
    
    'If Err Then Exit Sub
End Sub

Private Sub citarea_Click()

    If MsgBox("¿Desea iniciar todas las tareas seleccionadas?", vbYesNo + vbInformation, "Pregunta...") = vbYes Then
        If bagenda.Recordset.RecordCount = 0 Then Exit Sub
        bagenda.Refresh
        bagenda.Recordset.MoveFirst

        Do Until bagenda.Recordset.EOF
            bagenda.Recordset("Estado") = "Iniciada"
            bagenda.Recordset("Fecha") = fejecucion(bagenda.Recordset("periodo"))
            bagenda.Recordset.Update
            bagenda.Recordset.MoveNext
        Loop

    Else
        bagenda.Recordset("Estado") = "Iniciada"
        bagenda.Recordset("Fecha") = fejecucion(bagenda.Recordset("periodo"))
        bagenda.Recordset.Update
        
        bagenda.RecordSource = "select * from agenda where not (Estado = 'Completada')"
        bagenda.Refresh
    End If

End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    OrdenarDataGrid ColIndex, bagenda.Recordset, DataGrid1
    
If Err Then GrabarLog "DataGrid1_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub dgagenda_DblClick()
    Modificar
End Sub

Private Sub dgagenda_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    OrdenarDataGrid ColIndex, bagenda.Recordset, dgagenda

If Err Then GrabarLog "dgagenda_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub f_Click()

    If f.Value = 1 Then
        vcdesde.Enabled = False
        vchasta.Enabled = False

    Else
        vcdesde.Enabled = True
        vchasta.Enabled = True
    End If

End Sub

Public Function fejecucion(periodo As String) As Date
    Dim fecha As Date
    
    Select Case periodo

        Case "Solo una vez"
            fecha = vfecha.Value

        Case "Diario"
            fecha = vfecha.Value + 1

        Case "Semanal"
            fecha = vfecha.Value + 7

        Case "Mensual"
            fecha = bagenda.Recordset("Fecha") + 30
    End Select

    fejecucion = fecha
End Function

Private Sub Form_KeyPress(KeyCode As Integer)
    'If KeyCode = vbKeyF1 Then nuevo
    'If KeyCode = vbKeyF2 Then grabar
    'If KeyCode = vbKeyF3 Then Borrar
    'If KeyCode = vbKeyF4 Then Buscar
    'If KeyCode = vbKeyF5 Then Imprimir
    'If KeyCode = 27 Then Unload Me
End Sub

Private Sub Form_Load()

    calendario.Value = Date

    With bagenda
        .ConnectionString = pathDBMySQL
        .RecordSource = "Agenda"
        .Refresh
    End With

    With bconceptos
        .ConnectionString = pathDBMySQL
        .RecordSource = "Concepto"
        .Refresh
    End With

    With bpersonas
        .ConnectionString = pathDBMySQL
        .RecordSource = "Clientes"
        .Refresh
    End With
    
    tab_agenda.tab = 0
    tab_alta.tab = 0
    vfecha = Date
    vhora = Time
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Name
End Sub
Private Sub gdcalendario_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next

    OrdenarDataGrid ColIndex, bagenda.Recordset, gdcalendario

If Err Then GrabarLog "gdcalendario_HeadClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Grabar()

    If vtarea.Text = "" And vnota.Text = "" Then
        MsgBox "Debe ingresar alguna tarea/nota!", vbInformation, "Pregunta..."
        vtarea.SetFocus
        Exit Sub
    End If

    With bagenda
        .Refresh

        If Not vmodifica_agenda = 1 Then
            .Recordset.AddNew
        Else
            .RecordSource = "Select * from agenda where id = " + Trim(Str(vid)) + ""
            .Refresh
        End If

        .Recordset("Fecha") = fejecucion(vperiodo)
        .Recordset("hora") = vhora
        .Recordset("usuario") = Trim(vConfigGral.vUser)
        .Recordset("periodo") = Trim(vperiodo)
        .Recordset("tarea") = Trim(vtarea)
        .Recordset("codigo_cli") = Trim(vcod_cliente)
        .Recordset("concepto") = Trim(vconcepto)
        .Recordset("estado") = "Iniciada"

        If Not cbonota.Text = "" Then
            .Recordset("usuario_nota") = Trim(cbonota.Text)
        Else
            .Recordset("usuario_nota") = Trim(vConfigGral.vUser)
        End If

        .Recordset.Update
    End With

    Nuevo
End Sub

Private Sub menu_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Index

        Case 1
            tab_agenda.tab = 0

        Case 2
            Borrar bagenda.object, True

        Case 4
            Buscar
    End Select

End Sub

Private Sub Modificar()
    On Error Resume Next

    With bagenda
        vfecha = .Recordset("Fecha").Value
        vhora = .Recordset("hora").Value
        vConfigGral.vUser = Format(.Recordset("usuario").Value, "#################")
        vperiodo = Format(.Recordset("periodo").Value, "##################")
        vtarea = Format(.Recordset("tarea").Value, "################################################################################################################################")
        vnota = Format(.Recordset("notas").Value, "################################################################################################################################")
        vid = .Recordset("id").Value
        tab_agenda.tab = 0
    End With

    vmodifica_agenda = 1
    
    If Err.Number = 380 Then Exit Sub
End Sub

Private Sub Nuevo()
    vfecha = Date
    vhora = Time
    vperiodo.Text = ""
    vtarea = ""
    cbonota = ""
    vnota = ""
End Sub

Private Sub tab_agenda_Click(PreviousTab As Integer)
    On Error Resume Next

    Select Case tab_agenda.tab

        Case 0

            'Alta
        Case 1
            bagenda.RecordSource = "Select * from agenda where not estado = 'Completada'"

        Case 2
            calendario.Value = Date
            calendario_Click

        Case 3
            bagenda.RecordSource = "Select * from agenda where estado = 'Completada'"
    End Select

    bagenda.Refresh

    If Err Then GrabarLog "tab_agenda_Click", Err.Number & " " & Err.Description, Me.Name
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next

    Select Case Button.Index

        Case 1
            Nuevo

        Case 3
            Grabar
    End Select

    If Err Then GrabarLog "Toolbar1_ButtonClick", Err.Number & " " & Err.Description, Me.Name
End Sub

Public Sub vcliente_Keypress(Keyascii As Integer)

    If Keyascii = 13 Then
        buscacli
    End If

End Sub

Private Sub vconcepto_GotFocus()
On Error Resume Next
    vconcepto.Clear
    bconceptos.Refresh
    bconceptos.Recordset.MoveFirst

    Do Until bconceptos.Recordset.EOF
        vconcepto.AddItem bconceptos.Recordset("Concepto")
        bconceptos.Recordset.MoveNext
    Loop
If Err Then Exit Sub
End Sub

Private Sub vmonto_Change()
    vtotal.Text = Format(vmonto, "#########0.00")
End Sub
    
