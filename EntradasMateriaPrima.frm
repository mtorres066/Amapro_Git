VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form EntradasMateriaPrima 
   BackColor       =   &H00008000&
   Caption         =   $"EntradasMateriaPrima.frx":0000
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   Icon            =   "EntradasMateriaPrima.frx":00A3
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
      Caption         =   "Busqueda De Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11415
      Begin VB.Frame FrameTipos 
         Caption         =   "Tipos De Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   5640
         TabIndex        =   6
         Top             =   240
         Width           =   4335
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   2280
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   10440
         Picture         =   "EntradasMateriaPrima.frx":096D
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataBuscar 
         Caption         =   "Productos"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGridBuscar 
         Bindings        =   "EntradasMateriaPrima.frx":29DF
         Height          =   7455
         Left            =   120
         OleObjectBlob   =   "EntradasMateriaPrima.frx":29F8
         TabIndex        =   4
         ToolTipText     =   "Doble Click o Esc Para Seleccionar"
         Top             =   960
         Width           =   11055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4935
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      Connect         =   "pwd=metal"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin TabDlg.SSTab TabEntradas 
      Height          =   7935
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Encabezado"
      TabPicture(0)   =   "EntradasMateriaPrima.frx":33D0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameEncabezado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "EntradasMateriaPrima.frx":3822
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameDetalle"
      Tab(1).Control(1)=   "DBGridDetallePedidos"
      Tab(1).ControlCount=   2
      Begin MSDBGrid.DBGrid DBGridDetallePedidos 
         Bindings        =   "EntradasMateriaPrima.frx":3B3C
         Height          =   4695
         Left            =   -74760
         OleObjectBlob   =   "EntradasMateriaPrima.frx":3B5E
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   2400
         Width           =   11175
      End
      Begin VB.Frame FrameDetalle 
         Caption         =   "Detalle Materia Prima"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   7095
         Left            =   -75000
         TabIndex        =   53
         Top             =   720
         Width           =   11445
         Begin VB.Frame FrameDetalleCompras 
            Enabled         =   0   'False
            Height          =   1455
            Left            =   120
            TabIndex        =   54
            Top             =   240
            Width           =   11175
            Begin VB.TextBox TxtDesPro 
               Appearance      =   0  'Flat
               BackColor       =   &H000080FF&
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
               Left            =   3480
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   69
               TabStop         =   0   'False
               Top             =   480
               Width           =   3732
            End
            Begin VB.TextBox TxtCodPro 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               DataField       =   "Codigo"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   1800
               MaxLength       =   15
               TabIndex        =   56
               ToolTipText     =   "signo + o doble click para ayuda"
               Top             =   480
               Width           =   1635
            End
            Begin VB.TextBox TxtDocDet 
               Appearance      =   0  'Flat
               DataField       =   "Documento"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   3120
               MaxLength       =   15
               TabIndex        =   67
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox TxtBoleta 
               Appearance      =   0  'Flat
               DataField       =   "NumeroUnicoSerieBoleta"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Index           =   0
               Left            =   120
               MaxLength       =   20
               TabIndex        =   59
               Top             =   1080
               Width           =   1500
            End
            Begin VB.TextBox TxtBoleta 
               Appearance      =   0  'Flat
               DataField       =   "OrdenBoleta"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Index           =   1
               Left            =   1680
               MaxLength       =   20
               TabIndex        =   60
               Top             =   1080
               Width           =   1500
            End
            Begin VB.TextBox TxtBoleta 
               Appearance      =   0  'Flat
               DataField       =   "BultoBoleta"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Index           =   2
               Left            =   3240
               MaxLength       =   20
               TabIndex        =   61
               Top             =   1080
               Width           =   1500
            End
            Begin VB.TextBox TxtBoleta 
               Appearance      =   0  'Flat
               DataField       =   "BobinaBoleta"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Index           =   3
               Left            =   4800
               MaxLength       =   20
               TabIndex        =   62
               Top             =   1080
               Width           =   1500
            End
            Begin VB.TextBox TxtBoleta 
               Appearance      =   0  'Flat
               BackColor       =   &H000080FF&
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
               Index           =   4
               Left            =   8880
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   65
               TabStop         =   0   'False
               Top             =   480
               Width           =   2115
            End
            Begin VB.TextBox TxtBoleta 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               DataField       =   "PesoEntrada"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Index           =   5
               Left            =   7680
               TabIndex        =   64
               Top             =   1080
               Width           =   1155
            End
            Begin VB.CheckBox ChkMultiplica 
               BackColor       =   &H000080FF&
               Height          =   255
               Left            =   7320
               TabIndex        =   57
               TabStop         =   0   'False
               Top             =   480
               Width           =   255
            End
            Begin VB.TextBox TxtBoleta 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0E0FF&
               DataField       =   "OrdenProduccion"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Index           =   6
               Left            =   120
               MaxLength       =   15
               TabIndex        =   55
               Top             =   480
               Width           =   1635
            End
            Begin MSMask.MaskEdBox MskFecBol 
               DataField       =   "FechaBoleta"
               DataSource      =   "DataDetalleEntradas"
               Height          =   288
               Left            =   6360
               TabIndex        =   63
               Top             =   1080
               Width           =   1188
               _ExtentX        =   2090
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12640511
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TxtCanPro 
               DataField       =   "Cantidad"
               DataSource      =   "DataDetalleEntradas"
               Height          =   288
               Left            =   7680
               TabIndex        =   58
               Top             =   480
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   12640511
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               BackColor       =   &H000080FF&
               Caption         =   "Codigo"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   0
               Left            =   1800
               TabIndex        =   87
               Top             =   240
               Width           =   1452
            End
            Begin VB.Label Label2 
               BackColor       =   &H000080FF&
               Caption         =   "Descripcion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   0
               Left            =   3480
               TabIndex        =   86
               Top             =   240
               Width           =   1572
            End
            Begin VB.Label Label3 
               BackColor       =   &H000080FF&
               Caption         =   "Cantidad"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   0
               Left            =   7680
               TabIndex        =   85
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H000080FF&
               Caption         =   "# De Serie"
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
               Index           =   7
               Left            =   120
               TabIndex        =   84
               Top             =   840
               Width           =   930
            End
            Begin VB.Label Label6 
               BackColor       =   &H000080FF&
               Caption         =   "Orden/Proveedor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   1680
               TabIndex        =   83
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label6 
               BackColor       =   &H000080FF&
               Caption         =   "Bulto/Paleta"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   9
               Left            =   3240
               TabIndex        =   82
               Top             =   840
               Width           =   1212
            End
            Begin VB.Label Label6 
               BackColor       =   &H000080FF&
               Caption         =   "Fecha"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   10
               Left            =   6360
               TabIndex        =   81
               Top             =   840
               Width           =   972
            End
            Begin VB.Label Label6 
               BackColor       =   &H000080FF&
               Caption         =   "Bobina"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   252
               Index           =   11
               Left            =   4800
               TabIndex        =   80
               Top             =   840
               Width           =   972
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H000080FF&
               Caption         =   "Unidad Medida Bulto"
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
               Index           =   16
               Left            =   8880
               TabIndex        =   79
               Top             =   240
               Width           =   1785
            End
            Begin VB.Label Label6 
               BackColor       =   &H000080FF&
               Caption         =   "Peso"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   17
               Left            =   7680
               TabIndex        =   78
               Top             =   840
               Width           =   615
            End
            Begin VB.Label LblPeso 
               Appearance      =   0  'Flat
               BackColor       =   &H000080FF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   8880
               TabIndex        =   77
               Top             =   1080
               Width           =   2175
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackColor       =   &H000080FF&
               Caption         =   "Unidad Medida Peso"
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
               Index           =   1
               Left            =   8880
               TabIndex        =   75
               Top             =   840
               Width           =   1770
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H000080FF&
               Caption         =   "laminas x cuerpos"
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
               Index           =   8
               Left            =   6000
               TabIndex        =   73
               Top             =   240
               Width           =   1530
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H000080FF&
               Caption         =   "Orden Produccion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   192
               Index           =   19
               Left            =   120
               TabIndex        =   71
               Top             =   240
               Width           =   1548
            End
            Begin VB.Shape Shape1 
               BackColor       =   &H000080FF&
               BackStyle       =   1  'Opaque
               Height          =   1335
               Left            =   0
               Top             =   120
               Width           =   11175
            End
         End
         Begin VB.CommandButton CmdAgregar2 
            Caption         =   "A&gregar"
            Height          =   460
            Left            =   120
            Picture         =   "EntradasMateriaPrima.frx":6566
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   6480
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdGrabar2 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   460
            Left            =   3720
            Picture         =   "EntradasMateriaPrima.frx":6A98
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   6480
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdTerminar 
            Caption         =   "&Terminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   460
            Left            =   9120
            Picture         =   "EntradasMateriaPrima.frx":6FCA
            TabIndex        =   76
            Top             =   6480
            Visible         =   0   'False
            Width           =   2040
         End
         Begin VB.CommandButton CmdCancelar2 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   460
            Left            =   5520
            Picture         =   "EntradasMateriaPrima.frx":74FC
            Style           =   1  'Graphical
            TabIndex        =   72
            Top             =   6480
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdBorrar2 
            Caption         =   "B&orrar"
            Height          =   460
            Left            =   7320
            Picture         =   "EntradasMateriaPrima.frx":7A2E
            Style           =   1  'Graphical
            TabIndex        =   74
            Top             =   6480
            Visible         =   0   'False
            Width           =   1800
         End
         Begin VB.CommandButton CmdEditar2 
            Caption         =   "Editar"
            Height          =   460
            Left            =   1920
            Picture         =   "EntradasMateriaPrima.frx":7F60
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   6480
            Visible         =   0   'False
            Width           =   1800
         End
      End
      Begin VB.Frame FrameEncabezado 
         Caption         =   "Encabezado Entradas De Materia Prima"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   5415
         Left            =   0
         TabIndex        =   9
         Top             =   1560
         Width           =   11535
         Begin VB.Frame FrameCompras 
            Enabled         =   0   'False
            Height          =   4335
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   11295
            Begin VB.TextBox TxtDocIng 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               DataField       =   "Documento"
               DataSource      =   "DataEntradas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   4440
               Locked          =   -1  'True
               TabIndex        =   22
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TxtLib 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Liberado"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   9600
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtEncabezado 
               Appearance      =   0  'Flat
               DataField       =   "Observaciones"
               DataSource      =   "DataEntradas"
               Height          =   285
               Index           =   3
               Left            =   1560
               MaxLength       =   150
               TabIndex        =   28
               ToolTipText     =   "Maximo 150 Caracteres"
               Top             =   2400
               Width           =   6735
            End
            Begin VB.TextBox TxtBodega 
               Appearance      =   0  'Flat
               DataField       =   "Bodega"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   25
               ToolTipText     =   "signo + o doble click para ayuda"
               Top             =   1320
               Width           =   1575
            End
            Begin VB.TextBox TxtReq 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Requerido"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   9600
               MaxLength       =   10
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TxtEncabezado 
               Appearance      =   0  'Flat
               DataField       =   "Transportista"
               DataSource      =   "DataEntradas"
               Height          =   285
               Index           =   0
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   29
               Top             =   2760
               Width           =   1575
            End
            Begin VB.TextBox TxtEncabezado 
               Appearance      =   0  'Flat
               DataField       =   "NombreDePiloto"
               DataSource      =   "DataEntradas"
               Height          =   285
               Index           =   1
               Left            =   1560
               MaxLength       =   30
               TabIndex        =   30
               Top             =   3120
               Width           =   1575
            End
            Begin VB.TextBox TxtEncabezado 
               Appearance      =   0  'Flat
               DataField       =   "PlacasCamion"
               DataSource      =   "DataEntradas"
               Height          =   285
               Index           =   2
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   31
               Top             =   3480
               Width           =   1575
            End
            Begin VB.TextBox TxtEncabezado 
               Appearance      =   0  'Flat
               DataField       =   "Placas Furgon"
               DataSource      =   "DataEntradas"
               Height          =   285
               Index           =   4
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   32
               Top             =   3840
               Width           =   1575
            End
            Begin VB.TextBox TxtEncabezado 
               Appearance      =   0  'Flat
               DataField       =   "NumeroDocumento"
               DataSource      =   "DataEntradas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   285
               Index           =   5
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   23
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TxtTipDoc 
               Appearance      =   0  'Flat
               DataField       =   "TipoDeDocumento"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   24
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtPro 
               Appearance      =   0  'Flat
               DataField       =   "Proveedor"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   27
               ToolTipText     =   "signo + o doble click para ayuda"
               Top             =   2040
               Width           =   1575
            End
            Begin VB.TextBox TxtEncabezado 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               DataField       =   "Estado"
               DataSource      =   "DataEntradas"
               Height          =   285
               Index           =   7
               Left            =   9600
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TxtTipEnt 
               Appearance      =   0  'Flat
               DataField       =   "TipoEntrada"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   26
               ToolTipText     =   "signo + o doble click para ayuda"
               Top             =   1680
               Width           =   1575
            End
            Begin MSMask.MaskEdBox MskFec 
               DataField       =   "FechaEntrada"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   1560
               TabIndex        =   21
               Top             =   240
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Requerido"
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
               Index           =   21
               Left            =   8400
               TabIndex        =   90
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Liberado"
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
               Index           =   20
               Left            =   8400
               TabIndex        =   89
               Top             =   960
               Width           =   750
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Entrada"
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
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   240
               Width           =   1260
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Transaccion"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   195
               Left            =   3240
               TabIndex        =   51
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label6 
               Caption         =   "Observaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   50
               Top             =   2400
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Bodega"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   49
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label LblBodega 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   48
               Top             =   1320
               Width           =   5055
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Transportista"
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
               Index           =   2
               Left            =   120
               TabIndex        =   47
               Top             =   2760
               Width           =   1125
            End
            Begin VB.Label Label6 
               Caption         =   "Conductor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   120
               TabIndex        =   46
               Top             =   3120
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Placa Camion"
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
               Index           =   5
               Left            =   120
               TabIndex        =   45
               Top             =   3480
               Width           =   1170
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Placa Furgon"
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
               Index           =   6
               Left            =   120
               TabIndex        =   44
               Top             =   3840
               Width           =   1140
            End
            Begin VB.Label Label6 
               Caption         =   "Proveedor"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   13
               Left            =   120
               TabIndex        =   43
               Top             =   2040
               Width           =   975
            End
            Begin VB.Label LblProveedor 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   42
               Top             =   2040
               Width           =   5055
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "# Documento"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   41
               Top             =   600
               Width           =   1155
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Documento"
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
               Index           =   15
               Left            =   120
               TabIndex        =   40
               Top             =   960
               Width           =   1410
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Clasificacion"
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
               Index           =   12
               Left            =   8400
               TabIndex        =   39
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Entrada"
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
               Index           =   18
               Left            =   120
               TabIndex        =   38
               Top             =   1680
               Width           =   1110
            End
            Begin VB.Label LblTipoEntrada 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   37
               Top             =   1680
               Width           =   5055
            End
            Begin VB.Label LblTipDoc 
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3240
               TabIndex        =   36
               Top             =   960
               Width           =   5055
            End
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar"
            Height          =   600
            Left            =   120
            Picture         =   "EntradasMateriaPrima.frx":8492
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4680
            Width           =   1200
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "&Grabar"
            Enabled         =   0   'False
            Height          =   600
            Left            =   2520
            Picture         =   "EntradasMateriaPrima.frx":89C4
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   4680
            Width           =   1200
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   600
            Left            =   3720
            Picture         =   "EntradasMateriaPrima.frx":8EF6
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   4680
            Width           =   1200
         End
         Begin VB.CommandButton CmdBorrar 
            Caption         =   "&Borrar"
            Height          =   600
            Left            =   4920
            Picture         =   "EntradasMateriaPrima.frx":9428
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   4680
            Width           =   1200
         End
         Begin VB.CommandButton CmdSalida 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   10800
            Picture         =   "EntradasMateriaPrima.frx":995A
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Salida"
            Top             =   4680
            Width           =   600
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "B&uscar Documento"
            Height          =   600
            Left            =   6120
            Picture         =   "EntradasMateriaPrima.frx":B9CC
            TabIndex        =   15
            Top             =   4680
            Width           =   1200
         End
         Begin VB.CommandButton CmdEditar 
            Caption         =   "&Editar"
            Height          =   600
            Left            =   1320
            Picture         =   "EntradasMateriaPrima.frx":BEFE
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   4680
            Width           =   1200
         End
         Begin VB.CommandButton CmdImprimir 
            Caption         =   "&Imprimir"
            Height          =   600
            Left            =   9600
            Picture         =   "EntradasMateriaPrima.frx":C430
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   4680
            Width           =   1200
         End
         Begin VB.CommandButton CmdImprimirCedula 
            Caption         =   "Cedulas"
            Height          =   600
            Left            =   8400
            Picture         =   "EntradasMateriaPrima.frx":C962
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   4680
            Width           =   1200
         End
         Begin VB.CommandButton CmdBuscarSiguiente 
            Caption         =   "Siguiente Documento"
            Height          =   600
            Left            =   7320
            TabIndex        =   16
            Top             =   4680
            Width           =   1095
         End
      End
   End
   Begin VB.Data DataDetalleEntradas 
      Caption         =   "Detalle Entradas Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleEntradasMateriaPrima"
      Top             =   8040
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Data DataEntradas 
      Caption         =   "Entradas De Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Erick\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EncabezadoEntradasMateriaPrima"
      Top             =   7920
      Width           =   11295
   End
End
Attribute VB_Name = "EntradasMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As Long
Dim VDocumentoDetalle As Long
Dim VCantidad As Double

Dim VDiasDeAtraso As Long
Dim VCodigoProducto As String
Dim VCantidadProducto As Double
Dim VBodega As String
Dim VEstado As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim BBodega As Boolean
Dim BBodegaDetalle As Boolean
Dim BMateriaPrima As Boolean
Dim BProveedor As Boolean
Dim BTipoEntrada As Boolean
Dim BTipoDocumento As Boolean
Dim BEditar As Boolean

Dim RBuscaMateriaPrima As Recordset
Dim RBuscaSigDoc As Recordset
Dim RBuscaBodega As Recordset
Dim RBuscaProveedor As Recordset
Dim RBuscaDetalle As Recordset
Dim RBuscaEncabezado As Recordset
Dim RBuscaCuerpos As Recordset
Dim RBuscaTipoEntrada As Recordset
Dim RBuscaTipoDocumento As Recordset
Dim RBuscaDocumento As Recordset
Dim RBuscaFichaOrden As Recordset

Dim VCodigoFepsa As String
Dim VUltimaCantidad As Single
Dim VSerie As String
Dim VOrden As String
Dim VBulto As String
Dim VFecha As String
Dim VBobina As String
Dim VPeso As Single
Dim VOrdenProduccion As String
Dim VMultiplica As Boolean

Sub Botones1()
    If Bandera = True Then
         FrameCompras.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdBuscarSiguiente.Enabled = False
         CmdImprimirCedula.Enabled = False
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         DataEntradas.Visible = False
    Else
         FrameCompras.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdBuscarSiguiente.Enabled = True
         CmdImprimirCedula.Enabled = True
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         DataEntradas.Visible = True
         
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleCompras.Enabled = True
         CmdAgregar2.Enabled = False
         CmdEditar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalleCompras.Enabled = False
         CmdAgregar2.Enabled = True
         CmdEditar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
    End If

End Sub

Sub BotonesVisiblesDetalle()
    If Bandera3 = True Then
         CmdAgregar2.Visible = True
         CmdEditar2.Visible = True
         CmdGrabar2.Visible = True
         CmdTerminar.Visible = True
         CmdBorrar2.Visible = True
         CmdCancelar2.Visible = True
    Else
         CmdAgregar2.Visible = False
         CmdEditar2.Visible = False
         CmdGrabar2.Visible = False
         CmdTerminar.Visible = False
         CmdBorrar2.Visible = False
         CmdCancelar2.Visible = False
    End If

End Sub

Sub BotonesVisiblesEncabezado()
    If Bandera4 = True Then
         CmdAgregar.Visible = True
         CmdEditar.Visible = True
         CmdGrabar.Visible = True
         CmdCancelar.Visible = True
         CmdBorrar.Visible = True
         CmdBuscar.Visible = True
         CmdBuscarSiguiente.Visible = True
         CmdImprimirCedula.Visible = True
         CmdImprimir.Visible = True
         CmdSalida.Visible = True
    Else
         CmdAgregar.Visible = False
         CmdEditar.Visible = False
         CmdGrabar.Visible = False
         CmdCancelar.Visible = False
         CmdBorrar.Visible = False
         CmdBuscar.Visible = False
         CmdBuscarSiguiente.Visible = False
         CmdImprimirCedula.Visible = False
         CmdImprimir.Visible = False
         CmdSalida.Visible = False
    End If

End Sub


Private Sub CmdBuscarSiguiente_Click()
On Error Resume Next
    'mensaje = InputBox("Documento a Buscar")
    If mensaje = "" Then
    Else
        DataEntradas.Recordset.FindNext ("NumeroDocumento = '" & mensaje & "'")
    End If
    If Err <> 0 Then
       'MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       'Exit Sub
    End If

End Sub

Private Sub ChkMultiplica_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
End Sub

Private Sub CmdAgregar2_Click()
On Error Resume Next
    'AGREGA DATOS
    DataDetalleEntradas.Recordset.AddNew
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
    Bandera2 = True
    Botones2
    DBGridDetallePedidos.Enabled = False
    TxtDocDet.Text = VDocumento
        
    'ASIGNA LA CALIDAD A EL BULTO
    DataDetalleEntradas.Recordset!Calidad = "A"
    
    TxtDesPro.Text = ""
    
    'ASIGNA LOS VALORES GUARDADOS DE ULTIMO
    TxtCodPro.Text = VCodigoFepsa
    TxtCanPro.Text = VUltimaCantidad
    TxtBoleta.Item(0).Text = VSerie
    TxtBoleta.Item(1).Text = VOrden
    TxtBoleta.Item(2).Text = VBulto
    MskFecBol.Text = VFecha
    TxtBoleta.Item(3).Text = VBobina
    TxtBoleta.Item(5).Text = VPeso
    TxtBoleta.Item(6).Text = VOrdenProduccion
    ChkMultiplica.Value = VMultiplica
    'PARA GUATEMALA
    'TxtCodPro.SetFocus
    'PARA FEPSA
    TxtBoleta.Item(6).SetFocus

    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            If GBorrar = True Then
                'NO HACE NADA PORQUE SI TIENE ACCESO
            ElseIf TxtEncabezado.Item(7).Text = "LIBERADO" Then
                'VERIFICA SI YA FUE LIBERADA LA ENTRADA
                    MsgBox "Esta Recepcion No Se Puede BORRAR Porque Ya Fue Liberada", vbOKOnly + vbExclamation, "Informacion"
                    Exit Sub
            End If
            VDocumento = TxtDocIng.Text
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                        'BORRA EL ENCABEZADO DE EL PEDIDO
                        DataEntradas.Recordset.Delete
                        If Err <> 0 Then
                            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        DataEntradas.Recordset.MoveLast
                MousePointer = 0
            End If
            If DataEntradas.Recordset.EOF Then
                DataEntradas.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
            
            'ASIGANMOS A UNA VARIABLE EL DOCUMENTO DETALLE
            VDocumentoDetalle = TxtDocDet.Text
    
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                                        
                   'BORRA EL DETALLE DE LA ENTRADA
                    DataDetalleEntradas.Recordset.Delete
                    
                    If Err <> 0 Then
                       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                       Exit Sub
                    End If
                    'SELECCIONA TODOS LOS DETALLES DE LA ENTRADAS
                    DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasMateriaPrima where documento = " & VDocumentoDetalle & " order By Codigo")
                    DataDetalleEntradas.Refresh
                    DBGridDetallePedidos.Refresh
                MousePointer = 0
            End If
  
            If DataEntradas.Recordset.EOF Then
                DataEntradas.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
    mensaje = InputBox("Documento a Buscar")
    If mensaje = "" Then
    Else
          DataEntradas.Recordset.FindFirst ("NumeroDocumento = '" & mensaje & "'")
    End If
    If Err <> 0 Then
       'MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       'Exit Sub
    End If
    
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
    'CANCELA LOS CAMBIOS
    DataEntradas.Recordset.CancelUpdate
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error "
        Err.Clear
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    FrameDetalle.Visible = True
    DBGridDetallePedidos.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    'CANCELA LOS DATOS CAMBIADOS Y GRABA LOS DATOS COMO ESTABAN
    DataDetalleEntradas.Recordset.CancelUpdate
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    DBGridDetallePedidos.Enabled = True
    Bandera2 = False
    Botones2

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
    BEditar = True
    
    If GEditar = True Then
        'NO HACE NADA PORQUE SI TIENE ACCESO
    ElseIf TxtEncabezado.Item(7).Text = "LIBERADO" Then
        'VERIFICA SI YA FUE LIBERADA LA ENTRADA
        MsgBox "Esta Recepcion No Se Puede EDITAR Porque Ya Fue Liberada", vbOKOnly + vbExclamation, "Informacion"
        Exit Sub
    End If
    
    'EDITA EL REGISTRO
    DataEntradas.Recordset.Edit
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    MskFec.SetFocus
        
    FrameDetalle.Visible = False
    DBGridDetallePedidos.Visible = False
    
End Sub


Private Sub CmdEditar2_Click()
On Error Resume Next
    'AGREGA DATOS
    DataDetalleEntradas.Recordset.Edit
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
    Bandera2 = True
    Botones2
    DBGridDetallePedidos.Enabled = False
    'PARA GUATEMALA
    'TxtCodPro.SetFocus
    'PARA MEXICO
    TxtBoleta.Item(6).SetFocus
    TxtDesPro.Text = ""

End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
               
    'REVISAMOS DATOS
    If Not IsNumeric(TxtCanPro.Text) Then
       MsgBox "Cantidad de Materia Prima Incorrecta", vbOKOnly + vbCritical, "Error"
       TxtCanPro.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS PESO
    If Not IsNumeric(TxtBoleta.Item(5).Text) Then
       MsgBox "Peso Del Bulto Incorrecto", vbOKOnly + vbCritical, "Error"
       TxtBoleta.Item(5).SetFocus
       Exit Sub
    End If
    
    'REVISAMOS UNIDAD MEDIDA DEL BULTO
    If TxtBoleta.Item(4).Text = "" Then
       MsgBox "Unidad Medida No Puede Estar Vacio", vbOKOnly + vbCritical, "Error"
       TxtBoleta.Item(4).SetFocus
       Exit Sub
    End If
    
    'REVISAMOS DATOS
    If Not IsDate(MskFecBol.Text) Then
       MsgBox "Fecha De Boleta Incorrecta", vbOKOnly + vbCritical, "Error"
       MskFecBol.SetFocus
       Exit Sub
    End If
    
    'REVISA SI EXISTE LA MATERIA PRIMA EN LOS CORRELATIVOS MAXIMOS
    Set RBuscaMateriaPrima = Db.OpenRecordset("Select * from CorrelativosMateriaPrima where CodigoMateriaPrima = '" & TxtCodPro.Text & "'")
    If RBuscaMateriaPrima.RecordCount > 0 Then
    Else
        MsgBox "Codigo De Materia Prima No Existe En Correlativos Maximos", vbOKOnly + vbInformation, "Informacion"
        TxtCodPro.SetFocus
        Exit Sub
    End If
    
     VCantidad = TxtCanPro.Text
    
    'ASIGNAMOS A LA CANTIDAD DE TRASLADO LA CANTIDAD QUE ESTA ENTRANDO
    DataDetalleEntradas.Recordset!CantidadTraslado = VCantidad
    
    'ASIGNAMOS A LA CANTIDAD DE SALDO LA CANTIDAD QUE ESTA ENTRANDO
    DataDetalleEntradas.Recordset!SaldoDisponibilidad = VCantidad
    
    'ASIGNA LA BODEGA DONDE VA A ESTAR UBICADO EL BULTO POR EL MOMENTO
    DataDetalleEntradas.Recordset!BodegaDisponibilidad = VBodega
    
    'ASIGNA EL MISMO PESO INGRESADO PARA PESO DE ENTRADA
    DataDetalleEntradas.Recordset!PESO = TxtBoleta.Item(5).Text
    
    'ASIGNA COMO NO INSPECIONADO EL ESTADO DEL BULTO
    DataDetalleEntradas.Recordset!Estado = "N"
    
    'GUARDA VARIABLES TEMPORALES PARA AGREGAR EL ULTIMO DATO INGRESADO
    
    VCodigoFepsa = TxtCodPro.Text
    VUltimaCantidad = TxtCanPro.Text
    VSerie = TxtBoleta.Item(0).Text
    VOrden = TxtBoleta.Item(1).Text
    VBulto = TxtBoleta.Item(2).Text
    VFecha = MskFecBol.Text
    VBobina = TxtBoleta.Item(3).Text
    VPeso = TxtBoleta.Item(5).Text
    VOrdenProduccion = TxtBoleta.Item(6).Text
    VMultiplica = ChkMultiplica.Value
    
   
        
    'GRABA DATOS
    DataDetalleEntradas.Recordset.Update
        
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
        
    Bandera2 = False
    Botones2
    
    
         
    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE LA FACTURA QUE SE ESTA GRABANDO
    DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasMateriaPrima where Documento = " & VDocumento & " Order by Codigo")
    DataDetalleEntradas.Refresh
    DBGridDetallePedidos.Refresh
           
    DBGridDetallePedidos.Enabled = True
    TxtDesPro.Text = ""
    CmdAgregar2.SetFocus
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
    BEditar = False
    
    DataEntradas.Recordset.AddNew
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    
    'ASIGNA EL USUARIO
    TxtReq.Text = GUsuario
    MskFec.Text = Format(Date, "dd/mm/yyyy")
    MskFec.SetFocus
    TxtEncabezado.Item(7).Text = "NO LIBERADA"
    
    'BUSCA LA TRANSACCION MAXIMA Y LE SUMA 1
    Set RBuscaSigDoc = Db.OpenRecordset("Select Max(documento) from EncabezadoEntradasMateriaPrima")
        If RBuscaSigDoc.RecordCount > 0 Then
            If IsNull(RBuscaSigDoc(0)) Then
                TxtDocIng.Text = "1"
            Else
                TxtDocIng.Text = Val(RBuscaSigDoc(0)) + 1
            End If
        End If
    
    FrameDetalle.Visible = False
    DBGridDetallePedidos.Visible = False

End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    VDocumento = TxtDocIng.Text
    VBodega = TxtBodega.Text
    
    'OSEA QUE SI ESTA AGREGANDO UN REGISTRO
    If BEditar = False Then
            'BUSCA SI YA EXISTE EL NUMERO DE DOCUMENTO PARA ESTE TIPO DE DOCUMENTO
            Set RBuscaDocumento = Db.OpenRecordset("Select * From EncabezadoEntradasMateriaPrima Where TipoDeDocumento = '" & TxtTipDoc.Text & "' And NumeroDocumento = '" & TxtEncabezado.Item(5).Text & "'")
                    If RBuscaDocumento.RecordCount > 0 Then
                        MsgBox "Numero Documento Para Este Tipo De Documento Ya Existe", vbOKOnly + vbInformation, "Informacion"
                        TxtTipDoc.SetFocus
                        Exit Sub
                    End If
    End If
            
    'GRABA DATOS
    DataEntradas.Recordset.Update
    
    If Err = 3022 Then
        MsgBox "Transaccion Ya Existe ", vbOKOnly + vbCritical, "Informacion"
        TxtDocIng.SetFocus
        Exit Sub
    ElseIf Err <> 0 And Err <> 3022 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
    DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasMateriaPrima where Documento = " & VDocumento & " Order by Codigo")
    DataDetalleEntradas.Refresh
    DBGridDetallePedidos.Refresh
            
    'MUEVE EL RECORDSET A EL DOCUMENTO ACTUAL PARA QUE SE ACTUALIZEN LOS CAMBIOS
    DataEntradas.Recordset.FindFirst ("Documento = " & VDocumento)
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = False
    BotonesVisiblesEncabezado
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'VCodigoFepsa = ""
    VUltimaCantidad = 0
    VSerie = ""
    VOrden = ""
    VBulto = ""
    VFecha = ""
    VBobina = ""
    VPeso = 0
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = True
    FrameDetalle.Visible = True
    FrameEncabezado.Enabled = False
    DBGridDetallePedidos.Visible = True
    DBGridDetallePedidos.AllowDelete = True
    DBGridDetallePedidos.AllowUpdate = True
    
    'ESCONDE EL DATA
    DataEntradas.Visible = False
            
    TabEntradas.Tab = 1
    CmdAgregar2.SetFocus
    
    
End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDocIng.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
        
                CrReportes.SelectionFormula = "{EncabezadoEntradasMateriaPrima.Documento} = " & TxtDocIng.Text
                CrReportes.ReportFileName = App.Path & "\FormatoEntradasMateriaPrima.rpt"
                CrReportes.Action = 1
                
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description
                    Exit Sub
                End If
MousePointer = 0

End Sub

Private Sub CmdImprimirCedula_Click()
On Error Resume Next
MousePointer = 11
                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Documento} = " & TxtDocIng.Text
                CrReportes.ReportFileName = App.Path & "\CedulaMateriaPrima.rpt"
                CrReportes.Action = 1
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description
                    Exit Sub
                End If
                
MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdTerminar_Click()
If CmdCancelar2.Enabled = True Then
     CmdCancelar2_Click
End If
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesVisiblesEncabezado
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle
        
    'VISUALIZA EL DATA
    DataEntradas.Visible = True
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = False
    'FrameDetalle.Visible = False
    FrameEncabezado.Enabled = True
    
    TabEntradas.Tab = 0

End Sub

Private Sub Command1_Click()
    Framebuscar.Visible = False
End Sub



Private Sub DataDetalleEntradas_Validate(Action As Integer, Save As Integer)
                 Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida, UnidadMedidaPeso From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCodPro.Text & "'")
                 If RBuscaMateriaPrima.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaMateriaPrima!Descripcion
                            If Not IsNull(RBuscaMateriaPrima!UnidadMedidaPeso) Then
                                LblPeso.Caption = RBuscaMateriaPrima!UnidadMedidaPeso
                            End If
                 Else
                        TxtDesPro.Text = ""
                        LblPeso.Caption = ""
                 End If

End Sub

Private Sub DataEntradas_Error(DataErr As Integer, Response As Integer)
    On Error Resume Next
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
        End If
End Sub

Private Sub DataEntradas_Reposition()
    If IsNumeric(TxtDocIng.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
        DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasMateriaPrima where Documento = " & TxtDocIng.Text & " Order by Codigo")
        DataDetalleEntradas.Refresh
        DBGridDetallePedidos.Refresh
    End If

End Sub

Private Sub DBGridBuscar_DblClick()
    'BODEGA
    If BBodega = True Then
        TxtBodega.Text = DBGridBuscar.Columns(0)
        TxtBodega.SetFocus
    'MATERIA PRIMA
    ElseIf BMateriaPrima = True Then
        TxtCodPro.Text = DBGridBuscar.Columns(0)
        TxtCodPro.SetFocus
    'PROVEEDOR
    ElseIf BProveedor = True Then
        TxtPro.Text = DBGridBuscar.Columns(0)
        TxtPro.SetFocus
    'TIPO DE ENTRADA
    ElseIf BTipoEntrada = True Then
        TxtTipEnt.Text = DBGridBuscar.Columns(0)
        TxtTipEnt.SetFocus
    'TIPO DE DOCUMENTO
    ElseIf BTipoDocumento = True Then
        TxtTipDoc.Text = DBGridBuscar.Columns(0)
        TxtTipDoc.SetFocus
    End If
        TxtBuscar.Text = ""
        Framebuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
    'BODEGA
    If BBodega = True Then
        TxtBodega.Text = DBGridBuscar.Columns(0)
        TxtBodega.SetFocus
    'MATERIA PRIMA
    ElseIf BMateriaPrima = True Then
        TxtCodPro.Text = DBGridBuscar.Columns(0)
        TxtCodPro.SetFocus
    'PROVEEDOR
    ElseIf BProveedor = True Then
        TxtPro.Text = DBGridBuscar.Columns(0)
        TxtPro.SetFocus
    'TIPO DE ENTRADA
    ElseIf BTipoEntrada = True Then
        TxtTipEnt.Text = DBGridBuscar.Columns(0)
        TxtTipEnt.SetFocus
    'TIPO DE DOCUMENTO
    ElseIf BTipoDocumento = True Then
        TxtTipDoc.Text = DBGridBuscar.Columns(0)
        TxtTipDoc.SetFocus
    End If
        TxtBuscar.Text = ""
        Framebuscar.Visible = False
End If

End Sub
Private Sub Form_Activate()
    If IsNumeric(TxtDocIng.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
        DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasMateriaPrima where Documento = " & TxtDocIng.Text & " Order by Codigo")
        DataDetalleEntradas.Refresh
        DBGridDetallePedidos.Refresh
    End If
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
End Sub

Private Sub Form_Load()
    DataEntradas.ConnectionString = GTipoProveedor
    DataDetalleEntradas.ConnectionString = GTipoProveedor
    DataBuscar.ConnectionString = GTipoProveedor
    
    DataEntradas.Refresh
    DataDetalleEntradas.Refresh
    DataBuscar.Refresh
End Sub

Private Sub MskFec_GotFocus()
    MskFec.SelStart = 0
    MskFec.SelLength = Len(MskFec.Text)
End Sub
Private Sub MskFec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub MskFecBol_GotFocus()
    MskFecBol.SelStart = 0
    MskFecBol.SelLength = Len(MskFecBol.Text)
End Sub
Private Sub MskFecBol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtBodega_Change()
    Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtBodega.Text & "'")
    If RBuscaBodega.RecordCount > 0 Then
        LblBodega.Caption = RBuscaBodega!Descripcion
    Else
        LblBodega.Caption = ""
    End If
End Sub
Private Sub TxtBodega_DblClick()
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        FrameTipos.Visible = True
        BBodega = True
        BBodegaDetalle = False
        BMateriaPrima = False
        BProveedor = False
        BTipoEntrada = False
        BTipoDocumento = False
        Framebuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select * from BodegasMateriaPrima Order by CodigoBodega")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"

End Sub
Private Sub TxtBodega_GotFocus()
    TxtBodega.SelStart = 0
    TxtBodega.SelLength = Len(TxtBodega.Text)
End Sub

Private Sub TxtBodega_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
        
    If KeyAscii = 43 Then
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        FrameTipos.Visible = True
        BBodega = True
        BBodegaDetalle = False
        BMateriaPrima = False
        BProveedor = False
        BTipoEntrada = False
        BTipoDocumento = False
        Framebuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select * from BodegasMateriaPrima Order by CodigoBodega")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
    End If
End Sub

Private Sub TxtBoleta_GotFocus(Index As Integer)
    TxtBoleta.Item(Index).SelStart = 0
    TxtBoleta.Item(Index).SelLength = Len(TxtBoleta.Item(Index).Text)
End Sub
Private Sub TxtBoleta_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtBoleta_LostFocus(Index As Integer)
'ORDEN EN DETALLE DE PRODUCCION

'SOLO PARA FEPSA
            If Index = 6 Then
                Set RBuscaFichaOrden = Db.OpenRecordset("Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtBoleta.Item(6).Text & "'")
                    If RBuscaFichaOrden.RecordCount > 0 Then
                        TxtCodPro.Text = RBuscaFichaOrden!FichaTecnica
                    Else
                        TxtCodPro.Text = ""
                    End If
            End If
            
End Sub

Private Sub Txtbuscar_Change()
    'BODEGA
    If BBodega = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from BodegasMateriaPrima Where CodigoBodega Like '" & TxtBuscar.Text & "*' Order by CodigoBodega")
            Else
                    DataBuscar.RecordSource = ("Select * from BodegasMateriaPrima Where CodigoBodega Like '*" & TxtBuscar.Text & "*' Order by CodigoBodega")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from BodegasMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*' Order by CodigoBodega")
            Else
                    DataBuscar.RecordSource = ("Select * from BodegasMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*' Order by CodigoBodega")
            End If
        End If

    'MATERIA PRIMA
    ElseIf BMateriaPrima = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
            Else
                    DataBuscar.RecordSource = ("Select * From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * From CorrelativosMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
            Else
                    DataBuscar.RecordSource = ("Select * From CorrelativosMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
            End If
        End If
    'PROVEEDOR
    ElseIf BProveedor = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select CodigoProveedor, Proveedor from Proveedores Where CodigoProveedor Like '" & TxtBuscar.Text & "*'")
            Else
                    DataBuscar.RecordSource = ("Select CodigoProveedor, Proveedor from Proveedores Where CodigoProveedor Like '*" & TxtBuscar.Text & "*'")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select CodigoProveedor, Proveedor from Proveedores Where Proveedor Like '" & TxtBuscar.Text & "*'")
            Else
                    DataBuscar.RecordSource = ("Select CodigoProveedor, Proveedor from Proveedores Where Proveedor Like '*" & TxtBuscar.Text & "*'")
            End If
        End If
    'TIPO DE ENTRADA
    ElseIf BTipoEntrada = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from TiposEntradasMateriaPrima Where Codigo Like '" & TxtBuscar.Text & "*'")
            Else
                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from TiposEntradasMateriaPrima Where Codigo Like '*" & TxtBuscar.Text & "*'")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from TiposEntradasMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*'")
            Else
                    DataBuscar.RecordSource = ("Select Codigo, Descripcion from TiposEntradasMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*'")
            End If
        End If
    'DOCUMENTOS
    ElseIf BTipoDocumento = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select CodigoDocumento, Descripcion from Documentos Where CodigoDocumento Like '" & TxtBuscar.Text & "*'")
            Else
                    DataBuscar.RecordSource = ("Select CodigoDocumento, Descripcion from Documentos Where CodigoDocumento Like '*" & TxtBuscar.Text & "*'")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select CodigoDocumento, Descripcion from Documentos Where Descripcion Like '" & TxtBuscar.Text & "*'")
            Else
                    DataBuscar.RecordSource = ("Select CodigoDocumento, Descripcion from Documentos Where Descripcion Like '*" & TxtBuscar.Text & "*'")
            End If
        End If
    End If
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
End Sub
Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub
Private Sub TxtCanPro_GotFocus()
    TxtCanPro.SelStart = 0
    TxtCanPro.SelLength = Len(TxtCanPro.Text)
End Sub
Private Sub TxtCanPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtCanPro_LostFocus()
    'SI DESEA MULTIPLICAR LAS LAMINAS QUE VIENEN, BUSCA POR EL CODIGO CUANTAS LAMINA TIENE CADA CODIGO Y LAS MULTIPLICA
    If ChkMultiplica.Value = 1 Then
        If IsNumeric(TxtCanPro.Text) Then
            Set RBuscaCuerpos = Db.OpenRecordset("Select CuerposPorLamina From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCodPro.Text & "'")
                If RBuscaCuerpos.RecordCount > 0 Then
                    TxtCanPro.Text = TxtCanPro.Text * RBuscaCuerpos!CuerposPorLamina
                End If
        End If
    End If
End Sub

Private Sub TxtCodPro_Change()
                 Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida, UnidadMedidaPeso From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCodPro.Text & "'")
                 If RBuscaMateriaPrima.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaMateriaPrima!Descripcion
                        If Not IsNull(RBuscaMateriaPrima!UnidadMedida) Then
                            TxtBoleta.Item(4).Text = RBuscaMateriaPrima!UnidadMedida
                            If Not IsNull(RBuscaMateriaPrima!UnidadMedidaPeso) Then
                                LblPeso.Caption = RBuscaMateriaPrima!UnidadMedidaPeso
                            End If
                        End If
                 Else
                        TxtDesPro.Text = ""
                        TxtBoleta.Item(4).Text = ""
                        LblPeso.Caption = ""
                 End If
End Sub

Private Sub TxtCodPro_DblClick()
            TxtBuscar.Visible = True
            OptDescripcion.Visible = True
            OptCodigo.Visible = True
            FrameTipos.Visible = True
            BBodega = False
            BBodegaDetalle = False
            BMateriaPrima = True
            BProveedor = False
            BTipoEntrada = False
            BTipoDocumento = False
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
            DataBuscar.RecordSource = ("Select * from CorrelativosMateriaPrima Order by CodigoMateriaPrima")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"
End Sub

Private Sub TxtCodPro_GotFocus()
    TxtCodPro.SelStart = 0
    TxtCodPro.SelLength = Len(TxtCodPro.Text)
End Sub

Private Sub TxtCodPro_KeyPress(KeyAscii As Integer)
    'SI PRECIONA ENTER
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    'SI PRECIONA LA TECLA DE SIGNO +
    If KeyAscii = 43 Then
       TxtBuscar.Visible = True
       OptDescripcion.Visible = True
       OptCodigo.Visible = True
       FrameTipos.Visible = True
       BBodega = False
       BBodegaDetalle = False
       BMateriaPrima = True
       BProveedor = False
       BTipoEntrada = False
       BTipoDocumento = False
       Framebuscar.Visible = True
       TxtBuscar.SetFocus
       DataBuscar.RecordSource = ("Select * from CorrelativosMateriaPrima Order by CodigoMateriaPrima")
       DataBuscar.Refresh
       DBGridBuscar.Refresh
       DBGridBuscar.Columns(1).Width = "4000"
    End If
End Sub

Private Sub TxtCodPro_LostFocus()
Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida, UnidadMedidaPeso From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCodPro.Text & "'")
                 If RBuscaMateriaPrima.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaMateriaPrima!Descripcion
                            If Not IsNull(RBuscaMateriaPrima!UnidadMedidaPeso) Then
                                LblPeso.Caption = RBuscaMateriaPrima!UnidadMedidaPeso
                            End If
                 Else
                        TxtDesPro.Text = ""
                        LblPeso.Caption = ""
                 End If
End Sub

Private Sub TxtDesPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub
Private Sub TxtDocing_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtEncabezado_GotFocus(Index As Integer)
        TxtEncabezado.Item(Index).SelStart = 0
        TxtEncabezado.Item(Index).SelLength = Len(TxtEncabezado.Item(Index).Text)
End Sub
Private Sub TxtEncabezado_KeyPress(Index As Integer, KeyAscii As Integer)
        'SI PRECIONA ENTER
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtPro_Change()
    'RProveedor.Open "Select * From Produccion", GTipoProveedor, adOpenDynamic, adLockBatchOptimistic
    Set RBuscaProveedor = Db.OpenRecordset("Select Proveedor From Proveedores Where CodigoProveedor = '" & TxtPro.Text & "'")
    If RBuscaProveedor.RecordCount > 0 Then
    'If Not (RProveedor.EOF) And Not (RProveedor.BOF) Then
        LblProveedor.Caption = RBuscaProveedor!Proveedor
     '   LblProveedor.Caption = RProveedor(0)
    Else
        LblProveedor.Caption = ""
    End If
    'RProveedor.Close
End Sub

Private Sub TxtPro_DblClick()
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        FrameTipos.Visible = True
        BBodega = False
        BBodegaDetalle = False
        BMateriaPrima = False
        BProveedor = True
        BTipoEntrada = False
        BTipoDocumento = False
        Framebuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select CodigoProveedor, Proveedor from Proveedores")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"

End Sub

Private Sub TxtPro_GotFocus()
        TxtPro.SelStart = 0
        TxtPro.SelLength = Len(TxtPro.Text)
End Sub

Private Sub TxtPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
        
    If KeyAscii = 43 Then
            TxtBuscar.Visible = True
            OptDescripcion.Visible = True
            OptCodigo.Visible = True
            FrameTipos.Visible = True
            BBodega = False
            BBodegaDetalle = False
            BMateriaPrima = False
            BProveedor = True
            BTipoEntrada = False
            BTipoDocumento = False
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
            DataBuscar.RecordSource = ("Select CodigoProveedor, Proveedor from Proveedores")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"
    End If
End Sub

Private Sub TxtTipDoc_Change()
            Set RBuscaTipoDocumento = Db.OpenRecordset("Select Descripcion from Documentos Where CodigoDocumento = '" & TxtTipDoc.Text & "'")
                If RBuscaTipoDocumento.RecordCount > 0 Then
                    LblTipDoc.Caption = RBuscaTipoDocumento!Descripcion
                Else
                    LblTipDoc.Caption = ""
                End If
End Sub

Private Sub TxtTipDoc_DblClick()
            TxtBuscar.Visible = True
            OptDescripcion.Visible = True
            OptCodigo.Visible = True
            FrameTipos.Visible = True
            BBodega = False
            BBodegaDetalle = False
            BMateriaPrima = False
            BProveedor = False
            BTipoEntrada = False
            BTipoDocumento = True
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
            DataBuscar.RecordSource = ("Select * from Documentos Order by CodigoDocumento")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"

End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If
            
            If KeyAscii = 43 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                FrameTipos.Visible = True
                BBodega = False
                BBodegaDetalle = False
                BMateriaPrima = False
                BProveedor = False
                BTipoEntrada = False
                BTipoDocumento = True
                Framebuscar.Visible = True
                TxtBuscar.SetFocus
                DataBuscar.RecordSource = ("Select * from Documentos Order by CodigoDocumento")
                DataBuscar.Refresh
                DBGridBuscar.Refresh
                DBGridBuscar.Columns(1).Width = "4000"
            End If

End Sub

Private Sub TxtTipEnt_Change()
            Set RBuscaTipoEntrada = Db.OpenRecordset("Select Descripcion From TiposEntradasMateriaPrima Where Codigo = '" & TxtTipEnt.Text & "'")
                If RBuscaTipoEntrada.RecordCount > 0 Then
                    LblTipoEntrada.Caption = RBuscaTipoEntrada!Descripcion
                Else
                    LblTipoEntrada.Caption = ""
                End If
End Sub

Private Sub TxtTipEnt_DblClick()
            TxtBuscar.Visible = True
            OptDescripcion.Visible = True
            OptCodigo.Visible = True
            FrameTipos.Visible = True
            BBodega = False
            BBodegaDetalle = False
            BMateriaPrima = False
            BProveedor = False
            BTipoEntrada = True
            BTipoDocumento = False
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
            DataBuscar.RecordSource = ("Select * from TiposEntradasMateriaPrima Order by Codigo")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"

End Sub

Private Sub TxtTipEnt_GotFocus()
        TxtTipEnt.SelStart = 0
        TxtTipEnt.SelLength = Len(TxtTipEnt.Text)
End Sub

Private Sub TxtTipEnt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                FrameTipos.Visible = True
                BBodega = False
                BBodegaDetalle = False
                BMateriaPrima = False
                BProveedor = False
                BTipoEntrada = True
                BTipoDocumento = False
                Framebuscar.Visible = True
                TxtBuscar.SetFocus
                DataBuscar.RecordSource = ("Select * from TiposEntradasMateriaPrima Order by Codigo")
                DataBuscar.Refresh
                DBGridBuscar.Refresh
                DBGridBuscar.Columns(1).Width = "4000"
        End If

End Sub
