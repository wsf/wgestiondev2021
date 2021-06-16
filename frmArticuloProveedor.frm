VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.0#0"; "Codejock.Controls.v13.0.0.Demo.ocx"
Begin VB.Form frmArticuloProveedor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualización de Artículos desde la lista de Proveedores"
   ClientHeight    =   8325
   ClientLeft      =   3615
   ClientTop       =   1200
   ClientWidth     =   15810
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   15810
   ShowInTaskbar   =   0   'False
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grilla 
      Height          =   2115
      Left            =   30
      TabIndex        =   66
      Top             =   90
      Width           =   15765
      _ExtentX        =   27808
      _ExtentY        =   3731
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6045
      Left            =   30
      TabIndex        =   1
      Top             =   2250
      Width           =   15735
      _Version        =   851968
      _ExtentX        =   27755
      _ExtentY        =   10663
      _StockProps     =   68
      AllowReorder    =   -1  'True
      PaintManager.Layout=   1
      PaintManager.BoldSelected=   -1  'True
      PaintManager.HotTracking=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   2
      SelectedItem    =   1
      Item(0).Caption =   "Articulos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "Frame2"
      Item(1).Caption =   "Facturación automática"
      Item(1).ControlCount=   20
      Item(1).Control(0)=   "CheckBox2"
      Item(1).Control(1)=   "CheckBox1"
      Item(1).Control(2)=   "CheckBox3"
      Item(1).Control(3)=   "PushButton2"
      Item(1).Control(4)=   "txtPuntoVenta"
      Item(1).Control(5)=   "Label2"
      Item(1).Control(6)=   "vfecha"
      Item(1).Control(7)=   "Label3"
      Item(1).Control(8)=   "pb2"
      Item(1).Control(9)=   "Label4"
      Item(1).Control(10)=   "vreferencia"
      Item(1).Control(11)=   "PushButton3"
      Item(1).Control(12)=   "Label5"
      Item(1).Control(13)=   "vlog"
      Item(1).Control(14)=   "txtlinea"
      Item(1).Control(15)=   "Label6"
      Item(1).Control(16)=   "PushButton4"
      Item(1).Control(17)=   "PushButton5"
      Item(1).Control(18)=   "PushButton6"
      Item(1).Control(19)=   "PusLimpiarGrilla"
      Begin XtremeSuiteControls.PushButton PusLimpiarGrilla 
         Height          =   420
         Left            =   13680
         TabIndex        =   83
         Top             =   4185
         Width           =   1140
         _Version        =   851968
         _ExtentX        =   2011
         _ExtentY        =   741
         _StockProps     =   79
         Caption         =   "Limpiar Grilla"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton6 
         Height          =   435
         Left            =   6600
         TabIndex        =   82
         Top             =   5370
         Width           =   4395
         _Version        =   851968
         _ExtentX        =   7752
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Fijar los nros de compromante según afip"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton5 
         Height          =   435
         Left            =   1710
         TabIndex        =   81
         Top             =   5370
         Width           =   4725
         _Version        =   851968
         _ExtentX        =   8334
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Consultar parámetros"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   435
         Left            =   540
         TabIndex        =   80
         Top             =   5370
         Width           =   1035
         _Version        =   851968
         _ExtentX        =   1826
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Status FE"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.TextBox txtlinea 
         Height          =   285
         Left            =   12720
         TabIndex        =   77
         Text            =   "17"
         Top             =   750
         Width           =   1575
      End
      Begin VB.ListBox vlog 
         ForeColor       =   &H000000FF&
         Height          =   1815
         Left            =   540
         TabIndex        =   75
         Top             =   2250
         Width           =   14295
      End
      Begin XtremeSuiteControls.FlatEdit vreferencia 
         Height          =   315
         Left            =   8040
         TabIndex        =   73
         Top             =   1650
         Width           =   4365
         _Version        =   851968
         _ExtentX        =   7699
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.ProgressBar pb2 
         Height          =   315
         Left            =   540
         TabIndex        =   71
         Top             =   4830
         Width           =   14265
         _Version        =   851968
         _ExtentX        =   25162
         _ExtentY        =   556
         _StockProps     =   93
         Text            =   "Pb2"
      End
      Begin MSComCtl2.DTPicker vfecha 
         Height          =   345
         Left            =   8040
         TabIndex        =   69
         Top             =   1200
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   609
         _Version        =   393216
         Format          =   79560705
         CurrentDate     =   41598
      End
      Begin XtremeSuiteControls.FlatEdit txtPuntoVenta 
         Height          =   315
         Left            =   8040
         TabIndex        =   67
         Top             =   780
         Width           =   1635
         _Version        =   851968
         _ExtentX        =   2884
         _ExtentY        =   556
         _StockProps     =   77
         BackColor       =   -2147483643
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   435
         Left            =   555
         TabIndex        =   65
         Top             =   4200
         Width           =   12540
         _Version        =   851968
         _ExtentX        =   22119
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Ejecutar el proceso de confección de facturas automaticamente"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   5535
         Left            =   -69940
         TabIndex        =   2
         Top             =   510
         Visible         =   0   'False
         Width           =   15495
         Begin MSComctlLib.ProgressBar BarraExcel 
            Height          =   135
            Left            =   11370
            TabIndex        =   61
            Top             =   5220
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   238
            _Version        =   393216
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.TextBox txtCodigo1 
            Height          =   315
            Left            =   6030
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   1080
            Width           =   675
         End
         Begin VB.TextBox porcentaje 
            Height          =   315
            Left            =   6030
            TabIndex        =   35
            Top             =   2400
            Width           =   675
         End
         Begin VB.TextBox colprecio 
            Height          =   315
            Left            =   6030
            TabIndex        =   34
            Top             =   1740
            Width           =   675
         End
         Begin VB.TextBox col4 
            Height          =   315
            Left            =   8100
            TabIndex        =   33
            Top             =   1410
            Width           =   675
         End
         Begin VB.TextBox col3 
            Height          =   315
            Left            =   7410
            TabIndex        =   32
            Top             =   1410
            Width           =   675
         End
         Begin VB.TextBox col2 
            Height          =   315
            Left            =   6720
            TabIndex        =   31
            Top             =   1410
            Width           =   675
         End
         Begin VB.TextBox col1 
            Height          =   315
            Left            =   6030
            ScrollBars      =   2  'Vertical
            TabIndex        =   30
            Top             =   1410
            Width           =   675
         End
         Begin VB.OptionButton o1 
            Caption         =   "Dolar u$s"
            Height          =   225
            Left            =   7890
            TabIndex        =   29
            Top             =   1800
            Width           =   1095
         End
         Begin VB.OptionButton o2 
            Caption         =   "Pesos $"
            Height          =   225
            Left            =   6810
            TabIndex        =   28
            Top             =   1800
            Value           =   -1  'True
            Width           =   1005
         End
         Begin VB.TextBox vdescuento 
            Height          =   315
            Left            =   6030
            TabIndex        =   27
            Top             =   2730
            Width           =   675
         End
         Begin VB.CheckBox chkIva 
            Caption         =   "Este lista de precio contiene el IVA"
            Height          =   315
            Left            =   180
            TabIndex        =   26
            Top             =   4110
            Value           =   1  'Checked
            Width           =   2885
         End
         Begin VB.TextBox vdescuento2 
            Height          =   315
            Left            =   6720
            TabIndex        =   25
            Top             =   2730
            Width           =   675
         End
         Begin VB.TextBox vdescuento3 
            Height          =   315
            Left            =   7410
            TabIndex        =   24
            Top             =   2730
            Width           =   675
         End
         Begin VB.TextBox vdescuento4 
            Height          =   315
            Left            =   8100
            TabIndex        =   23
            Top             =   2730
            Width           =   675
         End
         Begin VB.TextBox txtCodigo2 
            Height          =   315
            Left            =   6720
            ScrollBars      =   2  'Vertical
            TabIndex        =   22
            Top             =   1080
            Width           =   675
         End
         Begin VB.TextBox txtDecimales 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   6030
            TabIndex        =   21
            Text            =   "2"
            Top             =   3090
            Width           =   1455
         End
         Begin VB.ComboBox cboProveedor 
            Height          =   315
            Left            =   6030
            TabIndex        =   20
            Top             =   2085
            Width           =   3885
         End
         Begin VB.ListBox LOG 
            Height          =   1620
            Index           =   0
            Left            =   10590
            TabIndex        =   19
            Top             =   2370
            Width           =   4425
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Borrar toda la lista actual"
            Height          =   585
            Left            =   11520
            TabIndex        =   18
            Top             =   1080
            Width           =   3405
         End
         Begin VB.ListBox vdisplay 
            Height          =   255
            Left            =   10650
            TabIndex        =   17
            Top             =   4080
            Width           =   4335
         End
         Begin VB.TextBox txtRubroCodigo 
            Height          =   315
            Left            =   6000
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   3420
            Width           =   675
         End
         Begin VB.TextBox txtRubroDescripcion 
            Height          =   315
            Left            =   7200
            ScrollBars      =   2  'Vertical
            TabIndex        =   15
            Top             =   3420
            Width           =   3105
         End
         Begin VB.TextBox txtSrubroCodigo 
            Height          =   315
            Left            =   6000
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   3780
            Width           =   675
         End
         Begin VB.TextBox txtSrubroDescripcion 
            Height          =   315
            Left            =   7200
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   3780
            Width           =   3105
         End
         Begin VB.TextBox vivaDescipcion 
            Height          =   315
            Left            =   7200
            ScrollBars      =   2  'Vertical
            TabIndex        =   12
            Top             =   4260
            Width           =   3105
         End
         Begin VB.TextBox vivaNro 
            Height          =   315
            Left            =   6000
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   4260
            Width           =   675
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   345
            Left            =   10680
            TabIndex        =   3
            Top             =   4440
            Width           =   2235
            _Version        =   851968
            _ExtentX        =   3942
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "PushButton1"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox2 
            Height          =   645
            Left            =   30
            TabIndex        =   4
            Top             =   4860
            Width           =   15435
            _Version        =   851968
            _ExtentX        =   27226
            _ExtentY        =   1138
            _StockProps     =   79
            UseVisualStyle  =   -1  'True
            Begin VB.CommandButton cmdTraspaso 
               Caption         =   "Ejecutar el traspaso de la lista"
               Height          =   375
               Left            =   12720
               Style           =   1  'Graphical
               TabIndex        =   9
               Top             =   180
               Width           =   2625
            End
            Begin VB.CommandButton cmdVerificar 
               Caption         =   "Verificar"
               Height          =   375
               Left            =   1650
               TabIndex        =   8
               Top             =   180
               Width           =   1425
            End
            Begin VB.CommandButton cmdCargarLista 
               Caption         =   "Cargar Lista"
               Height          =   375
               Left            =   60
               TabIndex        =   7
               Top             =   180
               Visible         =   0   'False
               Width           =   1605
            End
            Begin VB.CommandButton Command3 
               Caption         =   "Cargar una nueva lista"
               Height          =   375
               Left            =   3090
               TabIndex        =   6
               Top             =   180
               Width           =   2415
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Actualizar datos"
               Height          =   375
               Left            =   5520
               TabIndex        =   5
               Top             =   180
               Width           =   1755
            End
            Begin MSComctlLib.ProgressBar BarraArticulos 
               Height          =   165
               Left            =   7620
               TabIndex        =   10
               Top             =   180
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   291
               _Version        =   393216
               Appearance      =   0
               Scrolling       =   1
            End
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   1
            Left            =   6780
            TabIndex        =   37
            Tag             =   "SubRubro"
            Top             =   3450
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   2
            Left            =   6780
            TabIndex        =   38
            Tag             =   "Rubro"
            Top             =   3810
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton pbCarga 
            Height          =   315
            Index           =   0
            Left            =   6780
            TabIndex        =   39
            Tag             =   "Rubro"
            Top             =   4260
            Width           =   315
            _Version        =   851968
            _ExtentX        =   556
            _ExtentY        =   556
            _StockProps     =   79
            Caption         =   "..."
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.GroupBox GroupBox1 
            Height          =   975
            Left            =   0
            TabIndex        =   40
            Top             =   0
            Width           =   14925
            _Version        =   851968
            _ExtentX        =   26326
            _ExtentY        =   1720
            _StockProps     =   79
            Caption         =   "Filtros:"
            UseVisualStyle  =   -1  'True
            Begin VB.CommandButton cmdFiltrar 
               Caption         =   "Filtrar Datos"
               Height          =   345
               Left            =   10830
               TabIndex        =   44
               Top             =   570
               Width           =   4035
            End
            Begin VB.CheckBox negar 
               BackColor       =   &H8000000B&
               Caption         =   "Filtrar los datos que no cumplen con la condición"
               Height          =   345
               Left            =   150
               TabIndex        =   43
               Top             =   570
               Width           =   4155
            End
            Begin VB.ComboBox cboColumnas 
               Height          =   315
               ItemData        =   "frmArticuloProveedor.frx":0000
               Left            =   3180
               List            =   "frmArticuloProveedor.frx":003D
               TabIndex        =   42
               Text            =   "C1"
               Top             =   180
               Width           =   4185
            End
            Begin VB.TextBox varticulo 
               Height          =   315
               Left            =   10830
               TabIndex        =   41
               Top             =   180
               Width           =   4035
            End
            Begin VB.Label Label10 
               BackColor       =   &H00808080&
               BackStyle       =   0  'Transparent
               Caption         =   "> Ingresar nombre de la columna a Filtar :"
               Height          =   315
               Left            =   90
               TabIndex        =   46
               Top             =   270
               Width           =   3075
            End
            Begin VB.Label Label11 
               BackColor       =   &H00808080&
               BackStyle       =   0  'Transparent
               Caption         =   "> Ingresar columna a Filtar :"
               Height          =   225
               Left            =   8640
               TabIndex        =   45
               Top             =   240
               Width           =   2085
            End
         End
         Begin VB.Label display 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   90
            TabIndex        =   79
            Top             =   4590
            Width           =   10245
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Ingresar la columna donde se encuentra el Código del producto :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   60
            Top             =   1110
            Width           =   5385
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Ingresar las columnas que desea incorporar en la descripción :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   59
            Top             =   1470
            Width           =   5385
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Ingresar porcentaje que desea incrementrar al precio del producto."
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   58
            Top             =   2430
            Width           =   5865
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Ingresar el nombre del proveedor :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   57
            Top             =   2115
            Width           =   5385
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Ingresar las columnas donde se encuentra el precio del producto :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   56
            Top             =   1770
            Width           =   5385
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Ingresar descuento que realiza el proveedor :"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   55
            Top             =   2790
            Width           =   5385
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Ingrese cantidad de decimales que desea correr"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   90
            TabIndex        =   54
            Top             =   3120
            Width           =   5385
         End
         Begin VB.Label Label1 
            Caption         =   "Ingrese el nro de columna"
            Height          =   225
            Left            =   7470
            TabIndex        =   53
            Top             =   1140
            Width           =   1905
         End
         Begin VB.Label lblCódigoDe 
            Caption         =   "Código de artículos actualizados:"
            Height          =   225
            Left            =   10560
            TabIndex        =   52
            Top             =   2070
            Width           =   4545
         End
         Begin VB.Label vcantidadpasada 
            Alignment       =   2  'Center
            Caption         =   "0"
            ForeColor       =   &H00800000&
            Height          =   345
            Left            =   9720
            TabIndex        =   51
            Top             =   2610
            Width           =   795
         End
         Begin VB.Label vnopasado 
            Alignment       =   1  'Right Justify
            Caption         =   "0"
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   9360
            TabIndex        =   50
            Top             =   3060
            Width           =   795
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Rubro:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   49
            Top             =   3450
            Width           =   5385
         End
         Begin VB.Label lblActualizacion 
            Caption         =   "> Sub - Rubro:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   48
            Top             =   3810
            Width           =   5385
         End
         Begin VB.Label lblActualizacion 
            Caption         =   ">Tipo de IVA:"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   9
            Left            =   3810
            TabIndex        =   47
            Top             =   4140
            Width           =   2025
         End
      End
      Begin XtremeSuiteControls.CheckBox CheckBox2 
         Height          =   285
         Left            =   150
         TabIndex        =   62
         Top             =   780
         Width           =   3945
         _Version        =   851968
         _ExtentX        =   6959
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Paso 1. Cargar la planilla de pedidos"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox CheckBox1 
         Height          =   285
         Left            =   150
         TabIndex        =   63
         Top             =   1170
         Width           =   3945
         _Version        =   851968
         _ExtentX        =   6959
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Paso 2. Modificar los datos cargados."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.CheckBox CheckBox3 
         Height          =   285
         Left            =   150
         TabIndex        =   64
         Top             =   1590
         Width           =   3945
         _Version        =   851968
         _ExtentX        =   6959
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Paso 3. Confirmar los datos."
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   315
         Left            =   12540
         TabIndex        =   74
         Top             =   1650
         Width           =   1815
         _Version        =   851968
         _ExtentX        =   3201
         _ExtentY        =   556
         _StockProps     =   79
         Caption         =   "Ver  importaciones"
         UseVisualStyle  =   -1  'True
      End
      Begin VB.Label Label6 
         Caption         =   "Cantidad de lineas por Documento:"
         Height          =   255
         Left            =   10080
         TabIndex        =   78
         Top             =   780
         Width           =   2745
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   285
         Left            =   600
         TabIndex        =   76
         Top             =   2010
         Width           =   9705
         _Version        =   851968
         _ExtentX        =   17119
         _ExtentY        =   503
         _StockProps     =   79
         Caption         =   "Mensajes:"
         ForeColor       =   255
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Ingrese una referencia para esta operación:"
         Height          =   255
         Left            =   3990
         TabIndex        =   72
         Top             =   1740
         Width           =   3675
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha para los documentos a realizar:"
         Height          =   255
         Left            =   5010
         TabIndex        =   70
         Top             =   1230
         Width           =   2715
      End
      Begin VB.Label Label2 
         Caption         =   "Punto de ventas:"
         Height          =   255
         Left            =   6450
         TabIndex        =   68
         Top             =   840
         Width           =   1275
      End
   End
   Begin MSComDlg.CommonDialog file 
      Left            =   2370
      Top             =   5820
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid dgExcel 
      Bindings        =   "frmArticuloProveedor.frx":0097
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   1508
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
            LCID            =   1033
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
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmArticuloProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim descuento, vcodigo, cprecio As String
Dim vPrecio, vPrecioVenta As Double
Dim strFile, strWkShtName, strSQL, strError, strDir As String
Public XlDB As Data
Public XlRS As Recordset
Dim rsExcel As New ADODB.Recordset
Dim rsArticulos As New ADODB.Recordset
Dim totalesiva21, totalesiva105, totalesiva27 As Double
Dim vlistaPrecio As Integer
Dim vvnroFA, vvnroFB As Integer


Private Sub cboProveedor_GotFocus()
On Error Resume Next

    Call CargarCombo("Proveedores", "Nombre", cboProveedor, True)

If Err Then GrabarLog "cboProveedor_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cboProveedor_LostFocus()
On Error Resume Next
    
    cboProveedor.Tag = TraerDato("Proveedores", "Nombre = '" & Trim(cboProveedor.Text) & "'", "Codigo")

If Err Then GrabarLog "cboProveedor_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub cmdVerificar_Click()
On Error Resume Next

    Dim des, vcodigo, cprecio  As String
    Dim vPrecio, vPrecioVenta As Double
    
    If Val(txtCodigo1) > 0 Then vcodigo = rsExcel.Fields(Val(txtCodigo1) - 1).Value
    If Val(txtCodigo2) > 0 Then vcodigo = vcodigo + rsExcel.Fields(Val(txtCodigo2) - 1)
    If Val(col1) > 0 Then des = des + " " + Format((rsExcel.Fields(Val(col1) - 1)), "##################################################")
    If Val(col2) > 0 Then des = des + " " + Format((rsExcel.Fields(Val(col2) - 1)), "##################################################")
    If Val(col3) > 0 Then des = des + " " + Format((rsExcel.Fields(Val(col3) - 1)), "##################################################")
    If Val(col4) > 0 Then des = des + " " + Format((rsExcel.Fields(Val(col4) - 1)), "##################################################")
    ' MsgBox Val(b.Recordset(Val(colprecio) - 1))
    
    If Val(colprecio) > 0 Then
        cprecio = (inulo(rsExcel.Fields(Val(colprecio) - 1).Value))

        If Len(cprecio) > 0 Then
            If Val(txtDecimales.Text) >= 0 Then
                vPrecio = (PonerPunto(cprecio))
               ' vPrecio = Left(cprecio, Len(cprecio) - Val(txtDecimales.Text)) & "." & Right(cprecio, Val(txtDecimales.Text))
            Else
                vPrecio = (PonerPunto(cprecio))
            End If
        End If
    
    End If
    
    If chkIva.Value = 0 Then vPrecio = vPrecio * 1.21
    
    vPrecio = vPrecio - (vPrecio * Val(vdescuento) / 100)
    vPrecio = vPrecio - (vPrecio * Val(vdescuento2) / 100)
    vPrecio = vPrecio - (vPrecio * Val(vdescuento3) / 100)
    vPrecio = vPrecio - (vPrecio * Val(vdescuento4) / 100)
    
    vPrecioVenta = (vPrecio * Val(porcentaje) / 100) + vPrecio
    
    display.Caption = "- Codigo: " & EsNulo(vcodigo) & " - Descripción : " & EsNulo(des) & " - Costo: " & Format(vPrecio, "######0.000") & "     - Venta: " & Format(vPrecioVenta, "######0.000")
   
If Err Then GrabarLog "cmdVerificar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdTraspaso_Click()
    On Error Resume Next

Dim vsql As String

    If Not fvalidacion Then Exit Sub ' si validación es verdadero sigo

    MousePointer = vbHourglass

    HabilitarControles (False)
            
    rsExcel.Fields.Refresh
    BarraArticulos.Max = rsExcel.RecordCount
    
    vlog.Clear
    
    
    With rsArticulos
        If .State = 1 Then .Close
    
        .CursorLocation = adUseClient
        
        
        Call .Open("SELECT * FROM Articulos", ConnDDBB, adOpenStatic, adLockPessimistic)
        
        If Not .State = 1 Then
            MsgBox Err.Description
            Exit Sub
        End If
        
        .MoveFirst
    End With
    
    vdisplay.AddItem ("Cantidad a pasar: " + Str(rsExcel.RecordCount))
    
    Do Until rsExcel.EOF = True
        
        DoEvents
        vcodigo = ""
        descuento = ""
    
        Me.vlog.AddItem ((rsExcel.Fields(Val(txtCodigo1) - 1).Value))
    
        If Val(txtCodigo1) > 0 Then vcodigo = (rsExcel.Fields(Val(txtCodigo1) - 1).Value)
        If Val(txtCodigo2) > 0 Then vcodigo = vcodigo + rsExcel.Fields(Val(txtCodigo2) - 1).Value
        
        If Val(col1) > 0 Then descuento = descuento + " " + Format((rsExcel.Fields(Val(col1) - 1).Value), "##################################################")
        If Val(col2) > 0 Then descuento = descuento + " " + Format((rsExcel.Fields(Val(col2) - 1).Value), "##################################################")
        If Val(col3) > 0 Then descuento = descuento + " " + Format((rsExcel.Fields(Val(col3) - 1).Value), "##################################################")
        If Val(col4) > 0 Then descuento = descuento + " " + Format((rsExcel.Fields(Val(col4) - 1).Value), "##################################################")
    



        If Val(colprecio) > 0 Then
            vPrecio = (inulo(rsExcel.Fields(Val(colprecio) - 1)))

            If Len(cprecio) > 0 Then
                If Val(txtDecimales.Text) >= 0 Then
                    vPrecio = Left(cprecio, Len(cprecio) - Val(txtDecimales.Text)) & "." & Right(cprecio, Val(txtDecimales.Text))
                Else
                    vPrecio = (PonerPunto(cprecio))
                End If
            End If
    
        End If
        
        vPrecio = Round(vPrecio - (vPrecio * Val(vdescuento) / 100), Val(txtDecimales.Text))
        vPrecio = Round(vPrecio - (vPrecio * Val(vdescuento2) / 100), Val(txtDecimales.Text))
        vPrecio = Round(vPrecio - (vPrecio * Val(vdescuento3) / 100), Val(txtDecimales.Text))
        vPrecio = Round(vPrecio - (vPrecio * Val(vdescuento4) / 100), Val(txtDecimales.Text))
    
        vPrecioVenta = Round((vPrecio * Val(porcentaje) / 100) + vPrecio, Val(txtDecimales.Text))
        
        Dim sqlFiltro As String
        
        sqlFiltro = "SELECT * FROM articulos WHERE (codigo = '" & EsNulo(vcodigo) & "') AND (idproveedor = '" & Trim(cboProveedor.Tag) & "')"
        
        With rsArticulos
            If .State = 1 Then .Close
            
            Call .Open(sqlFiltro, ConnDDBB, adOpenStatic, adLockPessimistic)
            
            If .EOF = True Then
                Call GuardarArticulo(True)
            Else
                'log.AddItem (vCodigo)
                Call GuardarArticulo(False)
            End If
        End With

    
        'Call GuardarArticulo    'Graba en la base de datos.
    
        BarraArticulos.Value = BarraArticulos.Value + 1
    
        rsArticulos.Update

        rsExcel.MoveNext
    Loop

vsql = "update articulos set pventa5=pcosto, pventa2=pcosto, pventa3=pcosto, pventa4=pcosto"
Call EjecutarScript(vsql, pathDBMySQL)

    MsgBox "Trabajo Finalizado Correctamente"
    
    HabilitarControles (True)
    
    MousePointer = vbDefault

If Err Then GrabarLog "cmdTraspaso_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub HabilitarControles(vHabilita As Boolean)

    With Me
        .cmdCargarLista.Enabled = vHabilita
        .cmdFiltrar.Enabled = vHabilita
        .cmdTraspaso.Enabled = vHabilita
        .cmdVerificar.Enabled = vHabilita
    End With
    
End Sub
Private Sub cmdFiltrar_Click()
    On Error Resume Next
    
    MousePointer = vbHourglass
    
    Dim sqlExcel As String
    
    If negar.Value = 0 Then
        sqlExcel = "SELECT * FROM [" & strWkShtName & "$] WHERE [" & Trim(cboColumnas.Text) & "] LIKE '%" + Trim(varticulo.Text) + "%'"
    Else
        sqlExcel = "SELECT * FROM [" & strWkShtName & "$] WHERE not [" & Trim(cboColumnas.Text) & "] LIKE '%" + varticulo.Text + "%'"
    End If
    
    With rsExcel
        If .State = 1 Then .Close
        .CursorLocation = adUseClient
        
        Call .Open(sqlExcel, pathDBExcel(strFile), adOpenStatic, adLockReadOnly)
    
        If Not .State = 1 Then
            MsgBox Err.Description
            Exit Sub
        End If
    End With
    
    MousePointer = vbDefault

If Err Then GrabarLog "cmdFiltrar_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cmdCargarLista_Click()
    On Error Resume Next
    
    file.ShowOpen

    strDir = App.Path
    strFile = file.FileName
    
    If Not strFile = "" Then
    

        strWkShtName = TraerPropiedadExcel(strFile, "")
        strWkShtName = InputBox("Ingresar nombre de la Hoja de Excel :", "Mensaje ...", strWkShtName)
        
        Dim xlwbook As Excel.Workbook
        Dim xl As New Excel.Application
        Dim xlSheet As Excel.Worksheet

                    
        Set xlwbook = xl.Workbooks.Open(strFile)
        Set xlSheet = xlwbook.Sheets.Item(strWkShtName)

        ExcelToGrid xlSheet, xlwbook


    End If
    
    If Err Then GrabarLog "cmdCargarLista_Click", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub cboColumnas_GotFocus()
On Error Resume Next
    
    Dim i As Integer
    
    cboColumnas.Clear
    
    With dgExcel
        For i = 0 To .Columns.Count - 1
            cboColumnas.AddItem (.Columns(i).Caption)
        Next
    End With

If Err Then GrabarLog "cboColumnas_GotFocus", Err.Number & " " & Err.Description, Me.Caption
End Sub

Private Sub Command1_Click()

Dim vsql As String
vsql = "delete from articulos"

If MsgBox("Está seguro de borrar toda la lista de precio ?", vbYesNo, "Cuidado.") = vbYes Then
    Call EjecutarScript(vsql, pathDBMySQL)
End If

End Sub

Private Sub Command2_Click()
On Error Resume Next
With rsExcel
            
            If .State = 1 Then .Close
            .CursorLocation = adUseClient
            
            ' Call .Open("select * from excelexport", pathDBMySQL, adOpenStatic, adLockPessimistic)
         
           
           
           Call .Open("select * from excelexport", ConnDDBB, adOpenStatic, adLockPessimistic)
         '   Call .Open("SELECT * FROM articulos", ConnDDBB, adOpenStatic, adLockPessimistic)

            If Not .State = 1 Then
                MsgBox Err.Description
                Exit Sub
            Else
                Set dgExcel.DataSource = rsExcel
                Set grilla.Recordset = rsExcel
                
            End If
End With
If Err Then Exit Sub
End Sub

Private Sub Command3_Click()

'If LeerXml("UEmpresa") = "wgestionpons" Then
    'Call Shell("java -jar Listas-Pons.jar", 1)
'Else
    If UCase(LeerXml("UEmpresa")) = "WGESTION" Then
        Call Shell("ListasWGestion.bat", 1)
    Else
        'Call Shell("java -jar Listas.jar", 1)
        Call Shell("cargalista.bat", 1)
    
    End If
'End If

End Sub

Private Sub CheckBox2_Click()
Call Command3_Click
End Sub

Private Sub Form_Load()
On Error Resume Next

    Dim rsExcelExport As New ADODB.Recordset
       
    Dim cmdExcelExport As New ADODB.Command
    
    cmdExcelExport.ActiveConnection = ConnDDBB
    
    Dim sqlExcelExport, sqlBorrar As String


    dgExcel.ClearFields
    
   ' With Me
   '     .Show
   '     .Height = 11520
   '     .Width = 20490
   ' End With
    Me.Top = 0
    
    BarraExcel.Value = 0
    
    init

If Err Then GrabarLog "Form_Load", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub init()
    Call Command2_Click
    vfecha.Value = Date
    Me.txtPuntoVenta.Text = traerDatos2("select * from configuracion", "SucursalDocVenta", PathDBConfig)
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    Call BorrarBase("FormActivos WHERE (idUsuarios = " & vConfigGral.vIdUsuario & ") AND (idFormularios = " & Val(Me.Tag) & ")", PathDBConfig)

If Err Then GrabarLog "Form_Unload", Err.Number & " " & Err.Description, Me.Caption
End Sub
Private Sub GuardarArticulo(vnuevo As Boolean)
    On Error Resume Next
    
    Dim vadicionales As String
    
    
    If vcodigo = "" Then
        Me.vnopasado.Caption = Str(Val(Me.vnopasado.Caption) + 1)
        Exit Sub
    End If
    
    Dim vsql, vcampos, vvalores As String

    vcampos = "(Codigo,Descrip,Pcosto,PVenta1,idProveedor,idPorcentajeIva,idRubros,idSubRubros)"
    
    vvalores = "('" + EsNulo(vcodigo) + "','" + Trim(Left(EsNulo(descuento), 255)) + "'," + Str(inulo(vPrecio)) + "," + Str((vPrecio * Val(porcentaje.Text) / 100) + vPrecio) + ",'" + cboProveedor.Tag + "'," + Me.vivaNro + ",'" + Me.txtRubroCodigo.Tag + "','" + Me.txtSrubroCodigo.Tag + "')"
    
    If vnuevo Then
        vsql = "insert into articulos " + vcampos + " values " + vvalores
    Else
        
        vadicionales = ""
        If Not Me.txtSrubroCodigo = "" Then vadicionales = vadicionales + ",idSubRubros=" + Me.txtSrubroCodigo.Text
        If Not Me.txtRubroCodigo = "" Then vadicionales = vadicionales + ",idRubros=" + Me.txtRubroCodigo.Text
        If Not Me.vivaNro = "" Then vadicionales = vadicionales + ",idPorcentajeIva=" + Me.vivaNro.Text
        
        
        
        vsql = "update articulos set  Descrip='" + Trim(Left(EsNulo(descuento), 255)) + "'," + "Pcosto=" + Str(inulo(vPrecio)) + "," + "PVenta1=" + Str((vPrecio * Val(porcentaje.Text) / 100) + vPrecio) + vadicionales + " where codigo='" + EsNulo(vcodigo) + "'"
    
    End If
    Me.vcantidadpasada.Caption = Str(Val(Me.vcantidadpasada.Caption) + 1)
    
   If Not Trim(vvalores) = "" Then Call EjecutarScript(vsql, pathDBMySQL)
    
    
    ' se actualiza los precios de costo y los otros precios en proporción
    Call actualizarPreciosArticulo(vcodigo, vPrecio)
    
    Exit Sub
    
    
    With rsArticulos
        
        .Fields("Codigo").Value = EsNulo(vcodigo)
        .Fields("Descrip").Value = Trim(Left(EsNulo(descuento), 255))
        .Fields("Pcosto").Value = inulo(vPrecio)
        
        .Fields("idRubros").Value = Me.txtRubroCodigo.Text
        .Fields("idSubRubros").Value = Me.txtSrubroCodigo.Text
        
        If o1.Value = True Then
            .Fields("PDolar").Value = (vPrecio * Val(porcentaje) / 100) + vPrecio
        Else
            .Fields("PVenta1").Value = (vPrecio * Val(porcentaje.Text) / 100) + vPrecio
        End If

        .Fields("idProveedor").Value = Val(cboProveedor.Tag)
        
        .Fields("idPorcentajeIva").Value = Val(Me.vivaNro) ' este contiene el id del iva (2-iva21)
        
        Me.vcantidadpasada.Caption = Str(Val(Me.vcantidadpasada.Caption) + 1)
        
        .Update
        
    End With
    
If Err Then
'Me.log.AddItem ("Problema con el código: " + Str(vCodigo) + " - " + Trim(Left(EsNulo(descuento), 255)))
'GrabarLog "GuardarArticulo", Err.Number & " " & Err.Description, Me.Caption
End If
End Sub

'Actualizado 09/05/2010
Function ExcelToGrid(xls As Worksheet, wxls As Excel.Workbook) As Worksheet

    Dim nxlsheet As New Excel.Worksheet

    'Set xlwbook = xl.Workbooks.Open(strFile)
    Set nxlsheet = wxls.Sheets.Add()

    Dim fila, i, j As Long
    fila = xls.Range("B65000").End(xlUp).Row

    Dim rsExcelExport As New ADODB.Recordset

    Dim cmdExcelExport As New ADODB.Command

    cmdExcelExport.ActiveConnection = ConnDDBB
    Dim sqlExcelExport, sqlBorrar As String

    sqlExcelExport = "Select * from ExcelExport"
    sqlBorrar = "Delete from ExcelExport"

    If rsExcelExport.State = 0 Then
        rsExcelExport.Open sqlBorrar, ConnDDBB, adOpenKeyset, adLockOptimistic
        rsExcelExport.Open sqlExcelExport, ConnDDBB, adOpenKeyset, adLockOptimistic
    End If

    If Not rsExcelExport.BOF Then
        rsExcelExport.MoveFirst
    End If

    BarraExcel.Max = fila
    BarraExcel.Min = 0
    BarraExcel.Value = 0

    vdisplay.AddItem ("Cantidad de artículos: " + Str(fila))

    For i = 1 To fila
        On Error Resume Next
        If (IsNumeric(Trim(xls.Range("A" & i).Value)) Or IsNumeric(Trim(xls.Range("B" & i).Value)) Or IsNumeric(Trim(xls.Range("C" & i).Value)) Or IsNumeric(Trim(xls.Range("D" & i).Value)) Or IsNumeric(Trim(xls.Range("E" & i).Value)) Or IsNumeric(Trim(xls.Range("F" & i).Value)) Or IsNumeric(Trim(xls.Range("G" & i).Value)) Or IsNumeric(Trim(xls.Range("H" & i).Value)) Or IsNumeric(Trim(xls.Range("I" & i).Value)) Or IsNumeric(Trim(xls.Range("J" & i).Value)) Or IsNumeric(Trim(xls.Range("K" & i).Value)) Or IsNumeric(Trim(xls.Range("L" & i).Value)) Or IsNumeric(Trim(xls.Range("M" & i).Value)) Or IsNumeric(Trim(xls.Range("N" & i).Value)) Or IsNumeric(Trim(xls.Range("O" & i).Value)) Or IsNumeric(Trim(xls.Range("P" & i).Value)) Or IsNumeric(Trim(xls.Range("Q" & i).Value)) Or IsNumeric(Trim(xls.Range("R" & i).Value)) Or IsNumeric(Trim(xls.Range("S" & i).Value)) Or IsNumeric(Trim(xls.Range("T" & i).Value)) Or IsNumeric(Trim(xls.Range("U" & i).Value)) Or IsNumeric(Trim(xls.Range("V" & i).Value))) Then

            rsExcelExport.AddNew

            For j = 1 To 35
                rsExcelExport(j) = xls.Cells(i, j)
            Next j

            rsExcelExport.Update
            DoEvents
        End If
        BarraExcel.Value = Me.BarraExcel.Value + 1
     Next i

      With rsExcel
            If .State = 1 Then .Close
            .CursorLocation = adUseClient

            Call .Open("SELECT * FROM ExcelExport", ConnDDBB, adOpenStatic, adLockPessimistic)

            If Not .State = 1 Then
                MsgBox Err.Description
                Exit Function
            Else
                Set dgExcel.DataSource = rsExcel
                Set grilla.Recordset = rsExcel
            End If
      End With


End Function

Sub toXML()

    Dim strNombreArchivo, strRuta, strArchivoTexto As String, Nombre As String
    Dim f As Integer, x, fila As Long, i As Integer
    strNombreArchivo = "articulos.xml"
    strRuta = "C:\"
    strArchivoTexto = strRuta & strNombreArchivo
    
    'Abrimos el archivo para escribir
    f = FreeFile
    Open strArchivoTexto For Output As #f
    
    'Encabezado del Archivo
    Print #f, "<?xml version='1.0' encoding='UTF-8'?>"
    
    'escribimos al archivo
    Print #f, "<Articulos>"
        Nombre = ActiveSheet.Name
        Set x = Worksheets(Nombre)
        With x
            fila = Range("B65000").End(xlUp).Row
            For i = 1 To fila
                On Error Resume Next
                If IsNumeric(Trim(Range("B" & i).Value)) Then
                                         
                    Print #f, "<Item>"
                        Print #f, "<codigoAntiguo>" & StripString(Range("B" & i).Text) & "</codigoAntiguo>"
                        Print #f, "<descripcion>" & StripString(Trim(Range("C" & i).Text)) & "</descripcion>"
                        Print #f, "<codigo2010>" & StripString(Range("D" & i).Text) & "</codigo2010>"
                        Print #f, "<cantPorCaja>" & StripString(Range("E" & i).Text) & "</cantPorCaja>"
                        Print #f, "<base>" & StripString(Range("F" & i).Text) & "</base>"
                        Print #f, "<vidaMedia>" & StripString(Range("G" & i).Text) & "</vidaMedia>"
                        Print #f, "<precio>" & Range("K" & i).Text & "</precio>"
                                        
                    Print #f, "</Item>"
                    DoEvents
                End If
            Next i
        End With
    
    'Footer del Archivo
    Print #f, "</Articulos>"
    
    'cerramos el archivo de texto
    Close f
End Sub
      
      
      
' --------------------------------------------------
' Function StripString()
'
' Returns a string minus a set of specified chars.
' --------------------------------------------------
Function StripString(MyStr As Variant) As Variant
   On Error GoTo StripStringError

   Dim strChar As String, strHoldString As String
   Dim i As Integer

   ' Exit if the passed value is null.
   If IsNull(MyStr) Then Exit Function

   ' Exit if the passed value is not a string.
   If VarType(MyStr) <> 8 Then Exit Function

   ' Check each value for invalid characters.
   For i = 1 To Len(MyStr)
      strChar = Mid$(MyStr, i, 1)
      If IsDigit(strChar) Then
            strHoldString = strHoldString & strChar
      End If
   Next i

   ' Pass back corrected string.
   StripString = strHoldString

StripStringEnd:
         Exit Function

StripStringError:
         MsgBox Error$
         Resume StripStringEnd
      End Function
                    
Function IsDigit(ByVal cS As String) As Boolean
      ' Returns True if first character of cString is digit,
      ' otherwise False.
      Dim cTemp As String
      Dim vTemp As Integer

      'cTemp = cS.Substring(0, 1) ' get first character first
      ' cant pass Asc() an empty string
      If Len(Trim(cS)) = 0 Then
         IsDigit = False
      End If

      vTemp = Asc(cS)
      IsDigit = (vTemp < 128)
End Function
Private Sub DepurarExcel()
On Error Resume Next

    Dim rsImp As New ADODB.Recordset, rs As New ADODB.Recordset
    Dim cmdImp As New ADODB.Command, j As Integer
    
    cmdImp.ActiveConnection = ConnDDBB
    Dim sqlImp As String, sqlBorrar As String
    
    sqlImp = "Select * from ImpFacturaAutomatica"
    sqlBorrar = "Delete from ImpFacturaAutomatica"
    
    If rsImp.State = 0 Then
        rsImp.Open sqlBorrar, ConnDDBB, adOpenKeyset, adLockOptimistic
        rsImp.Open sqlImp, ConnDDBB, adOpenKeyset, adLockOptimistic
    End If
    
    If Not rs.BOF Then
        rs.MoveFirst
    End If
    'Insertar los registos para la impresion
    
    Do While Not rs.EOF
        rsImp.AddNew
        For j = 0 To rs.Fields.Count - 1
            rsImp(j) = rs(j)
        Next
        rsImp.Update
        rs.MoveNext
        
    Loop
        
    
    Unload Mantenimiento
    Wait (250)
    Load Mantenimiento
    
    'drGanancias.Show

    If Err Then GrabarLog "frmGanancia", Left(Err.Number & " " & Err.Description, 99), Me.Name
End Sub

Function fvalidacion() As Boolean
Dim vmensaje As String
vmensaje = ""
' --- proveedores
If Me.cboProveedor.Tag = "" Then vmensaje = Chr(13) + "- Debe ingresar un proveedor seleccionado de la lista"
'--------------

' ---- hay datos en el grid para pasar ------
If Not (rsExcel.BOF = False) Then vmensaje = Chr(13) + "- Debe cargar una lista de precio"
'---------------

' ------------- tiene que ingresar un iva ------------------------
If Not (Val(Me.vivaNro) >= 1) Then
    vmensaje = Chr(13) + "- Debe ingresar un tipo de iva"
End If
'------------------------------------------------------------------

If Not vmensaje = "" Then
    MsgBox vmensaje, vbCritical, "Error de validación de datos"
    fvalidacion = False
Else
    fvalidacion = True
End If

End Function

Private Sub pbCarga_Click(Index As Integer)
If Index = 0 Then Call fbuscarGrilla("PorcentajeIva", "Descripcion", "idPorcentajeIva", Me.vivaDescipcion.Name, Me)
If Index = 2 Then Call fbuscarGrilla("SubRubros", "SubRubro", "idSubRubros", Me.txtSrubroDescripcion.Name, Me)
If Index = 1 Then Call fbuscarGrilla("Rubros", "Rubro", "idRubros", Me.txtRubroDescripcion.Name, Me)
End Sub



Private Sub PushButton2_Click()
    'Call validarPlanilla

    Call finPlanilla
    Call importPedidos
    
    vreferencia.Text = ""
    
    grilla.Clear
    
    
End Sub
Private Sub finPlanilla()
Dim vsql As String
Dim valor As String
Dim ult As Integer
grilla.Rows = grilla.Rows + 1
ult = grilla.Rows - 1

Me.grilla.TextMatrix(ult, 2) = "Fin"

'vsql = "select * from excelexport here c2='fin'"
'valor = traerDatos2(vsql, "c2", pathDBMySQL)

'If valor = "" Then
'    vsql = "insert into excelexport (c2) values ('fin')"
'    Call EjecutarScript(vsql, pathDBMySQL)
'End If


End Sub

Private Sub PushButton4_Click()
frmFeStatus.Show
End Sub

Private Sub PushButton5_Click()
    If vvnroFA = 0 And vvnroFB = 0 Then
        MsgBox "Todavía no se fijaron parámetros"
    Else
        MsgBox "Fact B : " + Str(vvnroFA) + Chr(13) + "Fact B : " + Str(vvnroFB)
    End If
End Sub

Private Sub PushButton6_Click()
vvnroFA = 0
vvnroFB = 0
Call getStatusAfip2(vvnroFA, vvnroFB)
End Sub

Private Sub PusLimpiarGrilla_Click()
grilla.Clear
End Sub

Private Sub txtRubroDescripcion_Change()
Me.txtRubroCodigo.Text = Me.txtRubroDescripcion.Tag
End Sub

Private Sub txtSrubroDescripcion_Change()
Me.txtSrubroCodigo.Text = Me.txtSrubroDescripcion.Tag
End Sub

Private Sub vivaDescipcion_Change()
Me.vivaNro = Me.vivaDescipcion.Tag
End Sub


Private Sub importPedidos()

Dim i, vpos, j, vlinea  As Integer
Dim vcodigoCli, vcodigoArt, vCodigoAuxi, vTipoIva As String
Dim vtotal, vTotal210, vTotal105 As Double
Dim vnrointerno, vnroremito As Long

Dim lineas As Integer


'---init-------------------------------

lineas = Val(Me.txtlinea) - 1

i = 1
j = 0
vpos = 1
vtotal = 0
'vnroremito = UltimoRemito("factura") ' ver si ya existe esta función

vnroremito = NroRemitoNuevo
vnrointerno = UltimoNroInterno2

pb2.Value = 0
pb2.Max = Me.grilla.Rows - 1
    
    
'------------------------------------
    
    j = getPrimerCliente  ' se posiciona en el primer código  de cliente
    
    If j = 0 Then
        MsgBox "No se puedo detectar clientes en la planilla fuente", vbCritical
        Exit Sub
    End If
    
    vcodigoCli = Trim(Str(Val(grilla.TextMatrix(j, 3))))   ' poner una función que si es nro lo deje sin decimales, caso contrario igual
    
   ' If vcodigoCli = "15" Then MsgBox "15"
  
  vTipoIva = setTipoIva(vcodigoCli)
    
    pb2.Value = 0
    pb2.Max = Me.grilla.Rows - 1
    
    vcodigoArt = 0

totalesiva21 = 0
totalesiva27 = 0
totalesiva105 = 0
vlinea = 0


vlistaPrecio = Val(traerDatos2("select idlistas from clientes where codigo ='" + (vcodigoCli) + "'", "idlistas", pathDBMySQL))


For i = j To Me.grilla.Rows - 1
  
'If Not Trim(vcodigoArt) = "" Then
  
pb2.Value = i

    
    'If esCodigocliente(i) Or Me.grilla.TextMatrix(i, 2) = "Fin" Or vlinea > lineas Then
   
        If ((Me.grilla.TextMatrix(i, 2) = "Fin" Or esOtroCodigo(i, vcodigoCli) Or vlinea > lineas Or i = Me.grilla.Rows - 1)) Then      ' i>19 tiene que hacer otra factura porque no alcanzan las filas
            vTipoIva = setTipoIva(vcodigoCli)
           ' vnroremito = UltimoRemito("factura")
            
            Call cerrarFacura(vnroremito, vcodigoCli, vtotal, vnrointerno, vTipoIva, Trim(Me.vreferencia.Text))
           ' vnroremito = UltimoRemito("factura")
           
            vnroremito = NroRemitoNuevo
            vnrointerno = UltimoNroInterno2
            
            If vlinea > lineas And Trim(Me.grilla.TextMatrix(i, 3)) = "" Then
                vcodigoCli = vcodigoCli
                vlistaPrecio = Val(traerDatos2("select idlistas from clientes where codigo ='" + vcodigoCli + "'", "idlistas", pathDBMySQL))
                vlinea = 0
                Debug.Print " **************** separa en 2 facturas *********************"
               
            Else
                
                vcodigoCli = Trim(Str(Val(grilla.TextMatrix(i, 3))))  ' actualizo el código
                vlistaPrecio = Val(traerDatos2("select idlistas from clientes where codigo ='" + vcodigoCli + "'", "idlistas", pathDBMySQL))
    
            End If
            
            vtotal = 0
            vlinea = 0
            
            ' 07052014
            
            vTipoIva = setTipoIva(vcodigoCli)  ' fijo nuevamente el iva del codigo del cliente
            
            totalesiva21 = 0
            totalesiva27 = 0
            totalesiva105 = 0
            Debug.Print "-------------------------------------------------------------"
            
        End If
        
    'End If
    
    If esDetalle(i) And vlinea <= lineas Then
        
        'Set vtotales = vtotales + grabarFDetalle(vCodigo, vnroremito, i, vnroremito, vTipoIva, 21)
        'vTotal210 = vTotal210 + grabarFDetalle(vCodigo, vnroremito, i, vnroremito, vTipoIva, 21)
        'vTotal105 = vTotal105 + grabarFDetalle(vCodigo, vnroremito, i, vnroremito, vTipoIva, 10.5)
        vlinea = vlinea + 1
        vcodigoArt = Trim(grilla.TextMatrix(i, 4))
        'vTipoIva = setTipoIva(vcodigoArt)
        vtotal = vtotal + grabarFDetalle(vlinea, vcodigoArt, vnroremito, i, vnroremito, vTipoIva)
    End If
'Else
    
   ' vcodigoArt = Trim(grilla.TextMatrix(i + 1, 4)) ' poner una función que si es nro lo deje sin decimales, caso contrario igual
   ' vTipoIva = setTipoIva(vcodigoArt)

'End If
Next

MsgBox "Trabajo finalizado", vbExclamation

End Sub
Function esCodigocliente(ByVal i As Integer) As Boolean

If Trim(Me.grilla.TextMatrix(i, 3)) = "" Then
    esCodigocliente = False
Else
    esCodigocliente = True
End If

End Function


Function setTipoIvaCondi(ByVal vcodigo As String) As String
On Error Resume Next

Dim vIdTipoIva As String


vIdTipoIva = traerDatos2("select * from clientes where codigo = '" + vcodigo + "'", "idtipoiva", pathDBMySQL)

Select Case vIdTipoIva
    Case "001"
        setTipoIvaCondi = "Iva Responsable Inscripto"
    Case "002"
        setTipoIvaCondi = "Iva Responsable No Inscripto"
    Case "003"
        setTipoIvaCondi = "Responsable Monotributo"
    Case "004"
        setTipoIvaCondi = "Iva Exento"
    Case "005"
        setTipoIvaCondi = "Consumidor Final"
    Case "006"
        setTipoIvaCondi = "Iva No Responsable"
    Case Else
        setTipoIvaCondi = "Documento"
End Select


If Err < 0 Then

    vlog.AddItem "IVA. Problema al grabar el cliente con código: " + Str(vcodigo)

End If
End Function





Function setTipoIva(ByVal vcodigo As String) As String
On Error Resume Next

Dim vIdTipoIva As String


vIdTipoIva = traerDatos2("select * from clientes where codigo = '" + vcodigo + "'", "idtipoiva", pathDBMySQL)

Select Case vIdTipoIva
    Case "001"
        setTipoIva = "Fact A"
    Case "004"
        setTipoIva = "Fact B"
    Case "003"
        setTipoIva = "Fact B"
    Case "005"
        setTipoIva = "Documento"
    Case Else
        setTipoIva = "Documento"
End Select


If Err < 0 Then

    vlog.AddItem "IVA. Problema al grabar el cliente con código: " + Str(vcodigo)

End If
End Function

Function getPrimerCliente() As Integer
Dim i As Integer

getPrimerCliente = 0

For i = 1 To Me.grilla.Rows - 1

    If Val(grilla.TextMatrix(i, 3)) > 0 Then
        
        getPrimerCliente = i
        i = Me.grilla.Rows - 1
        
    End If

Next


End Function
Function esOtroCodigo(ByVal i As Integer, ByVal vcanterior As String) As Boolean
Dim vcodigo As String
        Debug.Print Str(i) + Str(vcanterior)
     If Not (Trim(vcanterior) = Trim(Str(Val(grilla.TextMatrix(i, 3))))) And (Not Trim(grilla.TextMatrix(i, 3)) = "") Then
     
        esOtroCodigo = True
     Else

        esOtroCodigo = False
        
     End If
    
    
End Function

Function esDetalle(ByVal i As Integer) As Boolean

If Not Trim(grilla.TextMatrix(i, 4)) = "" Then
'If (Not grilla.TextMatrix(i, 5) = "") And (Not Trim((grilla.TextMatrix(i, 6) + grilla.TextMatrix(i, 7))) = "") Then
    esDetalle = True
Else
    esDetalle = False
End If
End Function
Function fivaArticulo(vcodigo As String) As Double
Dim vsql As String

vsql = "SELECT * " + _
"From " + _
"  `articulos` " + _
"  INNER JOIN `porcentajeiva` ON (`articulos`.`idPorcentajeIva` = `porcentajeiva`.`idPorcentajeIva`)  where codigo = '" + vcodigo + "'"


fivaArticulo = Val(traerDatos2(vsql, "porcentaje", pathDBMySQL))

'If fivaArticulo = 0 Then fivaArticulo = 21

End Function

Function grabarFDetalle(ByVal vlinea As Integer, ByVal vcodigo As String, ByVal vnroremito As Long, ByVal i As Integer, vnrointerno As Long, vTipoIva As String) As Double

Dim vIva, vdescuento, vpreciounitario  As Double
Dim vIDFDetalle As Long
Dim vmensaje As String

'If Trim(vcodigo) = "15" Then MsgBox "15"


 ' vTipoIva = setTipoIva(vcodigo)

'If vTipoIva = "Fact A" Then

'    vIva = 1

'Else

'    vIva = fivaArticulo(vcodigo) ' 1.21
    
'End If



Dim vcampos, vvalores, vsql, vDetalle  As String

vcampos = "remito,codigo,detalle,cantidad,precio,total,descuento,tiva"

With Me.grilla

    vcodigo = .TextMatrix(i, 4)  ' codigo producto
    
    vmensaje = ""
    
    vIva = fivaArticulo(vcodigo)
    
    If Not vTipoIva = "Fact A" Then
         vmensaje = vTipoIva + "  " + " > Articulo: " + vcodigo + "  Con iva: " + Str((getprecio(vcodigo) * (1 + vIva / 100)))
         vlog.AddItem (vmensaje)
        vpreciounitario = (getprecio(vcodigo) * (1 + vIva / 100))
        vPrecio = getcantidad(i) * (getprecio(vcodigo) * (1 + vIva / 100))  ' total
    
    Else
        
        vPrecio = getcantidad(i) * getprecio(vcodigo)   ' total
        vpreciounitario = (getprecio(vcodigo))
    End If
    
    
    vdescuento = Val(.TextMatrix(i, 7))
    vPrecio = vPrecio - (vdescuento * vPrecio / 100)
    vDetalle = getDetalle(vcodigo)

    vvalores = Str(vnroremito) + ",'" + Trim(vcodigo) + "','" + Trim(vDetalle) + "'," + Str(getcantidad(i)) + "," + Str(vpreciounitario) + "," + Str(vPrecio) + "," + Str(vdescuento) + ",'" + Str(vIva) + "'"
    
    vsql = "insert into fdetalle (" + vcampos + ")" + " values " + "(" + vvalores + ")"
    
    If Trim(vcodigo) = "" Or Trim(vDetalle) = "" Then
        grabarFDetalle = 0
        Exit Function
    End If
    
    
    
    vmensaje = vTipoIva + "  " + " > Articulo: " + vcodigo + "  Precio: " + Str(vPrecio)
    vlog.AddItem (vmensaje)
    
    vlog.AddItem (vsql)
    
    
    Call EjecutarScript(vsql, pathDBMySQL)
    Debug.Print "(" + Str(vlinea) + ")  " + vsql
    
    
    vsql = "select max(idfdetalle) as c from fdetalle "
    vIDFDetalle = traerDatos2(vsql, "c", pathDBMySQL)
    
    Call GuardarEnStock("Automatico", vcodigo, Date, getcantidad(i), "Automático", vIDFDetalle, 0)
    

End With

        
If vTipoIva = "Fact A" Then
        
        Debug.Print "***********************************************************************"
        
        If vIva = 21 Then totalesiva21 = totalesiva21 + vPrecio * vIva / 100
        
        If vIva = 10.5 Then totalesiva105 = totalesiva105 + vPrecio * vIva / 100
    
        If vIva = 27 Then totalesiva27 = totalesiva27 + vPrecio * vIva
            
End If


grabarFDetalle = vPrecio
        
End Function

Function getDetalle(vcodigo As String)
On Error Resume Next
Dim vsql As String

vsql = "select * from articulos where codigo = '" + vcodigo + "'"
getDetalle = traerDatos2(vsql, "descrip", pathDBMySQL)


If Err Then
    
    getDetalle = ""
End If
End Function

Private Sub cerrarFacura(ByVal vnroremito As Long, ByVal vcodigo As String, ByVal vSubTotal As Double, ByVal vnrointerno As Long, ByVal vTipoIva As String, ByVal vref As String)
Dim vcampo, vvalores, vIdTipoIva, vnombre, vsql, vnrocomprobante, vLetra, vcomentario, iva As String
Dim vtotal, vIva210, vivaAgregado, vtotaliva   As Double


vTipoIva = setTipoIva(vcodigo)

If vcodigo = "0" Or vSubTotal = 0 Then Exit Sub
' --- init
vtotal = vSubTotal
vIva210 = 0


vtotaliva = totalesiva21 + totalesiva105 + totalesiva27


If vTipoIva = "Fact A" Then
    'vtotal = vSubTotal * getIvaValor(vCodigo)
    vtotal = vSubTotal + vtotaliva
   ' vIva210 = vSubTotal * (getIvaValor(vCodigo) - 1)
     
    vLetra = "A"
    iva = "Iva Responsable Inscripto"
End If

If Not vTipoIva = "Fact A" Then

    vivaAgregado = getIvaValor(vcodigo)
   ' iva = "Monotributo"
    vLetra = "B"
    iva = setTipoIvaCondi(vcodigo) ' fija la condición del iva según la tabla
    
End If


'-------------------------------------------------------------------------------
vnrocomprobante = Str(NroComprobanteNuevo(vTipoIva, vLetra, Me.txtPuntoVenta))

If vTipoIva = "Fact A" And vvnroFA > 0 Then
    vnrocomprobante = vvnroFA
    vvnroFA = 0
End If

If vTipoIva = "Fact B" And vvnroFB > 0 Then
    vnrocomprobante = vvnroFB
    vvnroFB = 0
End If
' ------------------------------------------------------------------------------



vcampo = "refexportpedidos,ncomprobante,letra,tipo,fecha,nrointerno,remito,codigo,nombre,PuntoDeVenta,subtotal,total,cuit,iva"

With Me.grilla

vnombre = getNombre(vcodigo)


vvalores = "'" + Trim(vref) + "','" + Trim(vnrocomprobante) + "','" + vLetra + "','" + vTipoIva + "'," + "'" + strfechaMySQL(vfecha) + "'," + Str(vnrointerno) + "," + Str(vnroremito) + ",'" + Trim(vcodigo) + "','" + Trim(vnombre) + "','" + Trim(txtPuntoVenta.Text) + "'," + Str(vSubTotal) + "," + Str(vtotal) + ",'" + getCuit(vcodigo) + "','" + Trim(iva) + "'"


End With

vsql = "insert into factura (" + vcampo + ") values (" + vvalores + ")"
Call EjecutarScript(vsql, pathDBMySQL)

'iva
Debug.Print vsql

Call setIvaVenta(vnrointerno, vnroremito, totalesiva21, totalesiva105, totalesiva27)


vcomentario = " Nro.Comp:" + Trim(vnrocomprobante) + " Tipo: " + vTipoIva + " PtoVenta: " + Trim(txtPuntoVenta.Text)


Call setsaldo(vcodigo)

Call setCtacte(vcodigo, vfecha, vtotal, vcomentario, vnrointerno)
Debug.Print vsql

        totalesiva21 = 0
        totalesiva27 = 0
        totalesiva105 = 0

End Sub


Private Sub setsaldo(vcodigo As String)
On Error Resume Next
Dim vsql, vsaldo  As String

vsaldo = Str(getSaldoCliente2(vcodigo))

vsql = "update factura set saldos = " + vsaldo + " where codigo ='" + vcodigo + "'"

Call EjecutarScript(vsql, pathDBMySQL)

If Err Then Exit Sub
End Sub


Function getNombre(vcodigo As String) As String

getNombre = ""
getNombre = traerDatos2("select * from clientes where codigo='" + Trim(vcodigo) + "'", "nombre", pathDBMySQL)

If getNombre = "" Then
    vlog.AddItem "El cliente " + vcodigo + " no se puedo encontrar en la base de clientes."
End If

End Function

Function getCuit(vcodigo As String)
Dim vsql As String
getCuit = ""
vsql = "select * from clientes where codigo='" + vcodigo + "'"
getCuit = traerDatos2(vsql, "cuit", pathDBMySQL)

If getCuit = "" Then
    vlog.AddItem "El cliente " + vcodigo + " no tiene un CUIT disponible. Verifique la factura antes de imprimir"
End If

End Function
Function getIvaValor(vcodigo As String) As Double
getIvaValor = 1.21
End Function

Private Sub setCtacte(ByVal vcodigo As String, ByVal vfecha As Date, ByVal vtotal As Double, ByVal vcomentario As String, ByVal vnrointerno As Long)
On Error Resume Next
Dim vcampos, vvalores, vsql  As String

vcampos = "codigo, fecha,comentario, nrointerno, debito,credito"
vvalores = "'" + Trim(vcodigo) + "','" + strfechaMySQL(vfecha) + "','" + Trim(vcomentario) + "'," + Str(vnrointerno) + "," + Str(vtotal) + ",0"

vsql = "insert into cuentascorrientes (" + vcampos + ") values (" + vvalores + ")"
    
Call EjecutarScript(vsql, pathDBMySQL)


If vcodigo = "329" Then
'MsgBox ""
End If
    
If Err < 0 Then
    vlog.AddItem "CtaCte. Problema al grabar el cliente con código: " + Str(vcodigo)
    Exit Sub
End If

End Sub
Function getIvaaValor(vcodigo As String) As Double
    getIvaaValor = 1.21
End Function

Private Sub setIvaVenta(ByVal vnrointerno As Long, ByVal vremito As Long, ByVal vIva210 As Double, ByVal viva105 As Double, ByVal viva27 As Double)
Dim vcampos, vvalores, vsql  As String

vcampos = "nrointerno,remito,iva210,iva105,iva270"
vvalores = Str(vnrointerno) + "," + Str(vremito) + "," + Str(vIva210) + "," + Str(viva105) + "," + Str(viva27)

vsql = "insert into ivafacturaventa (" + vcampos + ") values  (" + vvalores + ")"
If vIva210 > 0 Then Debug.Print "!!!!!!!!!!!!!!!!!!!!!!!!!! iva 20 !!!!!!!!!!!!!!!"
Debug.Print vsql
Call EjecutarScript(vsql, pathDBMySQL)

End Sub

Function getprecio(ByVal vcodigo As String) As Double
    getprecio = Val(traerDatos2("select * from articulos where codigo = '" + Trim(vcodigo) + "'", "pventa" + Trim(Str(vlistaPrecio)), pathDBMySQL))
End Function
Function getcantidad(ByVal i As Integer) As Double

If Val(Me.grilla.TextMatrix(i, 6)) > 0 Then
    getcantidad = Val(Me.grilla.TextMatrix(i, 6))
Else
    getcantidad = Val(Me.grilla.TextMatrix(i, 5))
End If

End Function
