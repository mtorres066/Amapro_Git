VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "dbgrid32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form SupervisionMateriasPrimas 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inspeccion De Materias Primas"
   ClientHeight    =   8595
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "SupervisionMateriasPrimas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   7455
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Data DataConsultas 
         Caption         =   "Defectos"
         Connect         =   "Access"
         DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Defectos"
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "SupervisionMateriasPrimas.frx":030A
         Height          =   6255
         Left            =   120
         OleObjectBlob   =   "SupervisionMateriasPrimas.frx":0326
         TabIndex        =   18
         Top             =   1080
         Width           =   11415
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   7440
         TabIndex        =   31
         Top             =   120
         Width           =   3135
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   1440
            TabIndex        =   20
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   17
         Top             =   720
         Width           =   6135
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   10680
         Picture         =   "SupervisionMateriasPrimas.frx":0D01
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   420
      ItemData        =   "SupervisionMateriasPrimas.frx":1143
      Left            =   7560
      List            =   "SupervisionMateriasPrimas.frx":114D
      TabIndex        =   50
      Top             =   3240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Defectos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   6
      Left            =   6840
      MouseIcon       =   "SupervisionMateriasPrimas.frx":1170
      Picture         =   "SupervisionMateriasPrimas.frx":15B2
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1200
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   615
      ItemData        =   "SupervisionMateriasPrimas.frx":1AE4
      Left            =   6360
      List            =   "SupervisionMateriasPrimas.frx":1AF1
      TabIndex        =   49
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Buscar Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   9480
      MouseIcon       =   "SupervisionMateriasPrimas.frx":1B16
      Picture         =   "SupervisionMateriasPrimas.frx":1F58
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   0
      Left            =   8160
      MouseIcon       =   "SupervisionMateriasPrimas.frx":248A
      Picture         =   "SupervisionMateriasPrimas.frx":28CC
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1200
   End
   Begin VB.TextBox TxtNumRec 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   7800
      Width           =   1455
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   9960
      Top             =   8160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Connect         =   "pwd=metal"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton CmdBotones 
      Height          =   465
      Index           =   5
      Left            =   10800
      MouseIcon       =   "SupervisionMateriasPrimas.frx":2DFE
      Picture         =   "SupervisionMateriasPrimas.frx":3240
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salida"
      Top             =   7680
      Width           =   840
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   2880
      MouseIcon       =   "SupervisionMateriasPrimas.frx":52B2
      Picture         =   "SupervisionMateriasPrimas.frx":56F4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1200
   End
   Begin TabDlg.SSTab TabMateriasPrimas 
      Height          =   7575
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   882
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "SupervisionMateriasPrimas.frx":5C26
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DBGridDetalleMateriasPrimas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameMateriasPrimas"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame FrameMateriasPrimas 
         Enabled         =   0   'False
         Height          =   2775
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   11535
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "FechaBoleta"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   3
            Left            =   7800
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   53
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Estado"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   1
            Left            =   3960
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   960
            Width           =   375
         End
         Begin VB.ComboBox CboCalidad 
            DataField       =   "Calidad"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   315
            ItemData        =   "SupervisionMateriasPrimas.frx":5C42
            Left            =   10080
            List            =   "SupervisionMateriasPrimas.frx":5C4F
            TabIndex        =   47
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "NumeroUnicoSerieBoleta"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   15
            Left            =   840
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   41
            Top             =   2400
            Width           =   1455
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "OrdenBoleta"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   14
            Left            =   3360
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   40
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "FechaBoleta"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   10
            Left            =   7320
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   39
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BobinaBoleta"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   8
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   38
            Top             =   2400
            Width           =   1695
         End
         Begin MSMask.MaskEdBox MskCanEnt 
            DataField       =   "Cantidad"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Left            =   2160
            TabIndex        =   11
            Top             =   1320
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BultoBoleta"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   13
            Left            =   5520
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   36
            Top             =   2400
            Width           =   975
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "NumeroIngreso"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   12
            Left            =   2160
            MaxLength       =   10
            TabIndex        =   13
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Observaciones"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   11
            Left            =   2160
            MaxLength       =   50
            TabIndex        =   14
            Top             =   2040
            Width           =   9255
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Codigo"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
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
            Index           =   2
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Documento"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
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
            Index           =   0
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "BodegaDisponibilidad"
            DataSource      =   "DataDetalleEntradaMateriaPrima"
            Height          =   285
            Index           =   4
            Left            =   2160
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label LblEstado 
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
            Left            =   4440
            TabIndex        =   52
            Top             =   960
            Width           =   2055
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
            Left            =   3960
            TabIndex        =   46
            Top             =   1680
            Width           =   7455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "# Serie"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   45
            Top             =   2400
            Width           =   510
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "# Orden"
            Height          =   195
            Index           =   6
            Left            =   2520
            TabIndex        =   44
            Top             =   2400
            Width           =   585
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   5
            Left            =   6720
            TabIndex        =   43
            Top             =   2400
            Width           =   450
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "# Bobina"
            Height          =   195
            Index           =   2
            Left            =   8760
            TabIndex        =   42
            Top             =   2400
            Width           =   645
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "# Bulto"
            Height          =   195
            Index           =   4
            Left            =   4800
            TabIndex        =   37
            Top             =   2400
            Width           =   510
         End
         Begin VB.Label lblLabels 
            Caption         =   "Numero Ingreso"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   1815
         End
         Begin VB.Image ImgAdvertencia 
            Height          =   480
            Left            =   10920
            Picture         =   "SupervisionMateriasPrimas.frx":5C5C
            Top             =   960
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image ImgNoConforme 
            Height          =   480
            Left            =   10920
            Picture         =   "SupervisionMateriasPrimas.frx":609E
            Top             =   960
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Image ImgAceptada 
            Height          =   480
            Left            =   10920
            Picture         =   "SupervisionMateriasPrimas.frx":64E0
            Top             =   960
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label LblUniMed 
            BackStyle       =   0  'Transparent
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
            Left            =   3960
            TabIndex        =   34
            Top             =   1320
            Width           =   2535
         End
         Begin VB.Label LblMateriaPrima 
            BackStyle       =   0  'Transparent
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
            Left            =   3960
            TabIndex        =   33
            Top             =   600
            Width           =   7455
         End
         Begin VB.Label lblLabels 
            Caption         =   "Observaciones"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Codigo Materia Prima"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No. Transaccion"
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
            TabIndex        =   28
            Top             =   240
            Width           =   1425
         End
         Begin VB.Label lblLabels 
            Caption         =   "Cantidad Entrada"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   27
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Bodega"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   26
            Top             =   1680
            Width           =   555
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Calidad"
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
            Index           =   22
            Left            =   9360
            TabIndex        =   25
            Top             =   1080
            Width           =   645
         End
      End
      Begin MSDBGrid.DBGrid DBGridDetalleMateriasPrimas 
         Bindings        =   "SupervisionMateriasPrimas.frx":6922
         Height          =   4455
         Left            =   120
         OleObjectBlob   =   "SupervisionMateriasPrimas.frx":694F
         TabIndex        =   8
         Top             =   3000
         Width           =   11535
      End
   End
   Begin VB.Data DataDetalleEntradaMateriaPrima 
      Caption         =   "Inspeccion De Materias Primas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleEntradasMateriaPrima"
      Top             =   8160
      Width           =   11760
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Inspeccionar"
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
      Height          =   465
      Index           =   3
      Left            =   4200
      MouseIcon       =   "SupervisionMateriasPrimas.frx":840E
      Picture         =   "SupervisionMateriasPrimas.frx":8850
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   4
      Left            =   5520
      MouseIcon       =   "SupervisionMateriasPrimas.frx":8D82
      Picture         =   "SupervisionMateriasPrimas.frx":91C4
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   1200
   End
   Begin VB.Label LblBusqueda 
      BackColor       =   &H00008000&
      Caption         =   "No. Transaccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   48
      Top             =   7680
      Width           =   1095
   End
End
Attribute VB_Name = "SupervisionMateriasPrimas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BBodega As Boolean
Dim BMateriaPrima As Boolean
Dim BModificar As Boolean

Dim RSupervisaMateriaPrima As Recordset
Dim RBuscaBodega As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaPedido As Recordset
Dim RBuscaCorrelativo As Recordset
Dim RBuscaFechaEntrada As Recordset
Dim RBuscaCorrelativoEnDetalle As Recordset
Dim RBuscaEncabezadoEntradas As Recordset


Dim R1 As Recordset
Dim R2 As Recordset

Dim VDiasDeAtraso As Long
Dim VCantidadEntrada As Double
Dim VNumeroPedido As Double
Dim VMateriaPrima As String
Dim VBodega As String
Dim VFechaEntrada As Date
Dim VCorrelativo As Double
Dim VRecepcion As String

Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         CmdBotones.Item(6).Enabled = False
         TxtNumRec.Visible = False
         LblBusqueda.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         CmdBotones.Item(5).Enabled = True
         TxtNumRec.Visible = True
         LblBusqueda.Visible = True
         
    End If
End Sub





Private Sub CboCalidad_Change()
            'CALIDAD
            If CboCalidad.Text = "A" Then
                ImgAceptada.Visible = True
                ImgAdvertencia.Visible = False
                ImgNoConforme.Visible = False
            ElseIf CboCalidad.Text = "P" Then
                ImgAceptada.Visible = False
                ImgAdvertencia.Visible = True
                ImgNoConforme.Visible = False
            ElseIf CboCalidad.Text = "R" Then
                ImgAceptada.Visible = False
                ImgAdvertencia.Visible = False
                ImgNoConforme.Visible = True
            End If
End Sub

Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
                'BUSQUEDA DE DATOS
                If Index = 0 Then
                                DataDetalleEntradaMateriaPrima.RecordSource = "Select * from DetalleEntradasMateriaPrima Where Documento = " & TxtNumRec.Text
                                DataDetalleEntradaMateriaPrima.Refresh
                                DBGridDetalleMateriasPrimas.Refresh
                'ACTUALIZAR
                ElseIf Index = 1 Then
                                DataDetalleEntradaMateriaPrima.RecordSource = "Select * from DetalleEntradasMateriaPrima"
                                DataDetalleEntradaMateriaPrima.Refresh
                                DBGridDetalleMateriasPrimas.Refresh
                'EDITAR
                ElseIf Index = 2 Then
                                    'ASIGNA EL NUMERO DE RECEPCION A UNA VARIABLE
                                    VRecepcion = TxtNumRec.Text
                                                                        
                                    'BUSCA EL DOCUMENTO DE ENTRADA
                                    Set RBuscaEncabezadoEntradas = Db.OpenRecordset("Select Estado From EncabezadoEntradasMateriaPrima Where Documento = " & VRecepcion)
                                    If RBuscaEncabezadoEntradas!Estado = "LIBERADO" Then
                                        MsgBox "Esta Recepcion No Se Puede EDITAR Porque ya fue LIBERADA", vbOKOnly + vbExclamation, "Informacion"
                                        TxtNumRec.SetFocus
                                        Exit Sub
                                    End If
                                    
                                                                        
                                    'SELECCIONA LOS REGISTROS QUE SEAN DEL DOCUMENTO INGRESADO
                                    Set RSupervisaMateriaPrima = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where Documento = " & VRecepcion)
                                    
                                    'SI LO ENCUENTRA
                                    If RSupervisaMateriaPrima.RecordCount > 0 Then
                                        'REVISA TODO EL DETALLE Y VERIFICA SI TIENE NUMERO DE INGRESOS LA LA MATERIA PRIMA
                                        'EN LA BASE DE CORRELATIVOS DE MATERIA PRIMA
                                        Do Until RSupervisaMateriaPrima.EOF
                                                'BUSCA EL CODIGO DE MATERIA PRIMA EN LA BASE DE CORRELATIVOS
                                                Set RBuscaCorrelativo = Db.OpenRecordset("Select * from CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & RSupervisaMateriaPrima!Codigo & "'")
                                                    If RBuscaCorrelativo.RecordCount > 0 Then
                                                        VCorrelativo = RBuscaCorrelativo!Correlativo + 1
                                                        'BUSCA NUMERO INGRESO MAXIMO EN EL DETALLE DE LAS ENTRADAS
                                                        Set RBuscaCorrelativoEnDetalle = Db.OpenRecordset("Select Max(NumeroIngreso) From DetalleEntradasMateriaPrima Where Codigo = '" & RBuscaCorrelativo!CodigoMateriaPrima & "'")
                                                            If RBuscaCorrelativoEnDetalle.RecordCount > 0 Then
                                                                If RBuscaCorrelativoEnDetalle(0) >= VCorrelativo Then
                                                                    MsgBox "El Correlativo Maximo Para La Materia Prima Ya Existe, El Correlativo Correcto Es " & (RBuscaCorrelativoEnDetalle(0) + 1), vbOKOnly + vbExclamation, "Verifique"
                                                                    Exit Sub
                                                                End If
                                                            End If
                                                    Else
                                                        MsgBox "Un Codigo De Materia Prima No Tiene Asignado Correlativo De Numero Ingreso", vbOKOnly + vbCritical, "Verifique"
                                                        Exit Sub
                                                    End If
                                            RSupervisaMateriaPrima.MoveNext
                                        Loop
                                    
                                    Else
                                        'SI NO LO ENCUENTRA
                                        MsgBox "La Recepcion " & TxtNumRec.Text & " No Existe", vbOKOnly + vbInformation, "Informacion"
                                        Exit Sub
                                    End If
                                    
                                    'SELECCIONA LOS REGISTROS QUE SEAN DEL DOCUMENTO INGRESADO
                                    DataDetalleEntradaMateriaPrima.RecordSource = ("Select * From DetalleEntradasMateriaPrima Where Documento = " & VRecepcion & " Order By Codigo")
                                    DataDetalleEntradaMateriaPrima.Refresh
                                    DBGridDetalleMateriasPrimas.Refresh
                                                                        
                                    If Err <> 0 Then
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                       Exit Sub
                                    End If
                                                                        
                                    'HABILITA EL GRID PARA MODIFICACIONES
                                    DBGridDetalleMateriasPrimas.AllowUpdate = True
                                    
                                    DBGridDetalleMateriasPrimas.SetFocus
                                    
                                    Bandera = True
                                    botones
                                    
                'GRABAR
                ElseIf Index = 3 Then
                                               
                            Db.Execute ("update DetalleEntradasMateriaPrima Set Estado = 'I' Where Documento = " & VRecepcion)
                            
                           'DESABILITA EL GRID PARA QUE YA NO PUEDAN EDITAR
                            DBGridDetalleMateriasPrimas.AllowUpdate = False
                            
                            DataDetalleEntradaMateriaPrima.RecordSource = "Select * from DetalleEntradasMateriaPrima Where Documento = " & TxtNumRec.Text
                            DataDetalleEntradaMateriaPrima.Refresh
                            DBGridDetalleMateriasPrimas.Refresh
                         
                            'CAMBIO BOTONES
                            Bandera = False
                            botones
                'IMPRIMIR
                ElseIf Index = 4 Then
                            MousePointer = 11
                                            CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Documento} = " & TxtTexto.Item(0).Text
                                            CrReportes.ReportFileName = App.Path & "\FormatoInspeccionMP.rpt"
                                            CrReportes.Action = 1
                                            
                                            If Err <> 0 Then
                                                MsgBox "Error " & Err.Number & " " & Err.Description
                                                MousePointer = 0
                                                Exit Sub
                                            End If
                            MousePointer = 0
                'SALIDA
                ElseIf Index = 5 Then
                                        Unload Me
                'DEFECTOS
                ElseIf Index = 6 Then
                    DefectosMateriaPrima.Show 1
                End If
    
    
End Sub


Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub DataDetalleEntradaMateriaPrima_Error(DataErr As Integer, Response As Integer)
    On Error Resume Next
    If Err <> 0 Then
        MsgBox "Error En Data " & Err.Description, vbOKOnly + vbInformation, "Error"
    End If
End Sub

    
Private Sub DataDetalleEntradaMateriaPrima_Reposition()
    'MATERIA PRIMA
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(2).Text & "'")
            If RBuscaMateriaPrima.RecordCount > 0 Then
                LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                If Not IsNull(RBuscaMateriaPrima!UnidadMedida) Then
                    LblUniMed.Caption = RBuscaMateriaPrima!UnidadMedida
                End If
            Else
                LblMateriaPrima.Caption = ""
                LblUniMed.Caption = ""
            End If
    
End Sub

Private Sub DataDetalleEntradaMateriaPrima_Validate(Action As Integer, Save As Integer)
    'MATERIA PRIMA
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(2).Text & "'")
            If RBuscaMateriaPrima.RecordCount > 0 Then
                LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                If Not IsNull(RBuscaMateriaPrima!UnidadMedida) Then
                    LblUniMed.Caption = RBuscaMateriaPrima!UnidadMedida
                End If
            Else
                LblMateriaPrima.Caption = ""
                LblUniMed.Caption = ""
            End If
    
    

End Sub



Private Sub DBGridConsultas_DblClick()
                    DBGridDetalleMateriasPrimas.Columns(8).Text = DBGridConsultas.Columns(0).Text
                    DBGridDetalleMateriasPrimas.SetFocus
                    FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
                If KeyAscii = 43 Then
                    DBGridDetalleMateriasPrimas.Columns(8).Text = DBGridConsultas.Columns(0).Text
                    DBGridDetalleMateriasPrimas.SetFocus
                    FrameConsultas.Visible = False
                End If
End Sub

Private Sub DBGridDetalleMateriasPrimas_HeadClick(ByVal ColIndex As Integer)
    DataDetalleEntradaMateriaPrima.RecordSource = ("Select * from DetalleEntradasMateriaPrima order by " & DBGridDetalleMateriasPrimas.Columns(ColIndex).DataField)
    DataDetalleEntradaMateriaPrima.Refresh
    DBGridDetalleMateriasPrimas.Refresh
End Sub

Private Sub DBGridDetalleMateriasPrimas_BeforeColEdit(ByVal ColIndex As Integer, ByVal KeyAscii As Integer, Cancel As Integer)

        '// Use este evento para que cuando el usuario teclee un caracter sobrela celda
        '// se despliegue la lista. Es decir, se obliga al usuario a usar un ítem de la lista.
        '// En caso de dar al usuario libertad de escribir, elimine las siguientes líneas (If-End If),
        '// o precedales con un comentario
            'If (ColIndex = 12 Or ColIndex = 13 Or ColIndex = 14 Or ColIndex = 15) Then
            If (ColIndex = 5) Then
               '// Se obliga a seleccionar de la lista:
               Cancel = True
               DBGridDetalleMateriasPrimas_ButtonClick (ColIndex)
            ElseIf (ColIndex = 6) Then
               '// Se obliga a seleccionar de la lista:
               Cancel = True
               DBGridDetalleMateriasPrimas_ButtonClick (ColIndex)
            End If
        
     
End Sub

Private Sub DBGridDetalleMateriasPrimas_ButtonClick(ByVal ColIndex As Integer)

'SI EL GRID ESTA HABILITADO PARA PODER GRABAR DATOS
If DBGridDetalleMateriasPrimas.AllowUpdate = True Then
        
        'SI PRECIONA EL BUTON DE LA COLUMNA DE CALIDAD
        If ColIndex = 5 Then
            Dim C As Column
            Set C = DBGridDetalleMateriasPrimas.Columns(ColIndex)
            With List1
                '// Despliegue de la lista al lado de la celda.
                '// Elimine los comentarios de las dos siguientes líneas
                '// y coloque comentarios a las tres posteriores. A su gusto
                
                .Left = DBGridDetalleMateriasPrimas.Left + C.Left + C.Width
                .Top = DBGridDetalleMateriasPrimas.Top + DBGridDetalleMateriasPrimas.RowTop(DBGridDetalleMateriasPrimas.Row)

                '// Lista debajo de la celda, al estilo ComboBox (3 líneas)
                .Left = DBGridDetalleMateriasPrimas.Left + C.Left
                .Top = DBGridDetalleMateriasPrimas.Top + DBGridDetalleMateriasPrimas.RowTop(DBGridDetalleMateriasPrimas.Row) + DBGridDetalleMateriasPrimas.RowHeight
                .Width = C.Width + 15

                .ListIndex = 0
                .Visible = True
                .ZOrder 0
                .SetFocus
            End With
        'SI PRECIONA EL BUTON DE LA COLUMNA DE ESTADO
        ElseIf ColIndex = 6 Then
            Dim C2 As Column
            Set C2 = DBGridDetalleMateriasPrimas.Columns(ColIndex)
            With List2
                '// Despliegue de la lista al lado de la celda.
                '// Elimine los comentarios de las dos siguientes líneas
                '// y coloque comentarios a las tres posteriores. A su gusto
                
                .Left = DBGridDetalleMateriasPrimas.Left + C2.Left + C2.Width
                .Top = DBGridDetalleMateriasPrimas.Top + DBGridDetalleMateriasPrimas.RowTop(DBGridDetalleMateriasPrimas.Row)

                '// Lista debajo de la celda, al estilo ComboBox (3 líneas)
                .Left = DBGridDetalleMateriasPrimas.Left + C2.Left
                .Top = DBGridDetalleMateriasPrimas.Top + DBGridDetalleMateriasPrimas.RowTop(DBGridDetalleMateriasPrimas.Row) + DBGridDetalleMateriasPrimas.RowHeight
                .Width = C2.Width + 15

                .ListIndex = 0
                .Visible = True
                .ZOrder 0
                .SetFocus
            End With
        'BODEGAS DE MATERIA PRIMA
        ElseIf ColIndex = 8 Then
            DataConsultas.RecordSource = "Select * From BodegasMateriaPrima"
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            TxtBuscar.SetFocus
        End If
End If
         
End Sub

Private Sub DBGridDetalleMateriasPrimas_DblClick()
    If DBGridDetalleMateriasPrimas.AllowUpdate = True Then
            
            'BODEGAS DE MATERIA PRIMA
            If DBGridDetalleMateriasPrimas.Col = 8 Then
                DataConsultas.RecordSource = "Select * From BodegasMateriaPrima"
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                FrameConsultas.Visible = True
                TxtBuscar.SetFocus
            End If
                DBGridConsultas.Columns(0).Width = "1500"
                DBGridConsultas.Columns(1).Width = "4000"
    End If
                
End Sub

Private Sub DBGridDetalleMateriasPrimas_KeyPress(KeyAscii As Integer)
        
    If DBGridDetalleMateriasPrimas.AllowUpdate = True Then
            'SI EL INDICE DE COLUMNA PERTENECE A LOS 3 DEFECTOS
            If KeyAscii = 43 Then
                    DataConsultas.RecordSource = "Select * From Defectos"
                    DataConsultas.Refresh
                    DBGridConsultas.Refresh
                'BODEGAS DE MATERIA PRIMA
                    If DBGridDetalleMateriasPrimas.Col = 8 Then
                        DataConsultas.RecordSource = "Select * From BodegasMateriaPrima"
                        DataConsultas.Refresh
                        DBGridConsultas.Refresh
                        FrameConsultas.Visible = True
                        TxtBuscar.SetFocus
                    End If
            End If
                    DBGridConsultas.Columns(0).Width = "1500"
                    DBGridConsultas.Columns(1).Width = "4000"
    End If
End Sub

Private Sub DBGridDetalleMateriasPrimas_Scroll(Cancel As Integer)
         '//Oculta la lista si hace Scroll
         List1.Visible = False
         List2.Visible = False
End Sub


Private Sub Form_Load()
        DataDetalleEntradaMateriaPrima.ConnectionString = GTipoProveedor
        DataConsultas.ConnectionString = GTipoProveedor
        
        DataDetalleEntradaMateriaPrima.Refresh
        DataConsultas.Refresh
End Sub


Private Sub List2_Click()
          DBGridDetalleMateriasPrimas.Columns(6).Text = Mid(List2.Text, 1, 1)
          List2.Visible = False
          DBGridDetalleMateriasPrimas.SetFocus

End Sub

Private Sub MskCanEnt_GotFocus()
    MskCanEnt.SelStart = 0
    MskCanEnt.SelLength = Len(MskCanEnt.Text)
End Sub

Private Sub MskCanEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub




Private Sub OptCod_Click()
    LblBuscar.Caption = "Codigo"
End Sub

Private Sub OptDes_Click()
    LblBuscar.Caption = "Descripcion"
End Sub

Private Sub Txtbuscar_Change()
                'DESCRIPCION
                    If OptDes.Value = True Then
                        If OptPalIni.Value = True Then
                            DataConsultas.RecordSource = "Select * from BodegasMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*'"
                        Else
                            DataConsultas.RecordSource = "Select * from BodegasMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*'"
                        End If
                    'CODIGO
                    ElseIf OptCod.Value = True Then
                        If OptPalIni.Value = True Then
                            DataConsultas.RecordSource = "Select * from BodegasMateriaPrima Where CodigoBodega Like '" & TxtBuscar.Text & "*'"
                        Else
                            DataConsultas.RecordSource = "Select * from BodegasMateriaPrima Where CodigoBodega Like '*" & TxtBuscar.Text & "*'"
                        End If
                    End If
                    DataConsultas.Refresh
                    DBGridConsultas.Refresh

End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub



Private Sub TxtNumRec_DblClick()
On Error Resume Next
    'Set R1 = Db.OpenRecordset("Select * From DefectosMateriaPrima")
    
    'Set R2 = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima")
    
    'Do Until R2.EOF
    '        R1.AddNew
    '            R1!CodigoMateriaPrima = R2!Codigo
    '            R1!NumeroIngreso = R2!NumeroIngreso
    '            If IsNull(R2!Defecto2) Then
    '            Else
    '                R1!Defecto = R2!Defecto2
    '            End If
    '        R1.Update
    '        If Err <> 0 Then
    '            'MsgBox Err.Description
    '        End If
    '    R2.MoveNext
    'Loop
    
    'MsgBox "ya"
    
End Sub

Private Sub TxtNumRec_GotFocus()
    TxtNumRec.SelStart = 0
    TxtNumRec.SelLength = Len(TxtNumRec.Text)
End Sub

Private Sub TxtNumRec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
    
    If Index = 1 Then
            If TxtTexto.Item(1).Text = "I" Then
                LblEstado.Caption = "INSPECCIONADO"
            ElseIf TxtTexto.Item(1).Text = "N" Then
                LblEstado.Caption = "NO INSPECCIONADO"
            End If
            
    'MATERIA PRIMA
    ElseIf Index = 2 Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(2).Text & "'")
            If RBuscaMateriaPrima.RecordCount > 0 Then
                LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                If Not IsNull(RBuscaMateriaPrima!UnidadMedida) Then
                    LblUniMed.Caption = RBuscaMateriaPrima!UnidadMedida
                End If
            Else
                LblMateriaPrima.Caption = ""
                LblUniMed.Caption = ""
            End If
    
    'BODEGA DIS0PONIBLE
    ElseIf Index = 4 Then
            Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtTexto.Item(4).Text & "'")
                If RBuscaBodega.RecordCount > 0 Then
                    LblBodega.Caption = RBuscaBodega!Descripcion
                Else
                    LblBodega.Caption = ""
                End If
    
    End If
    
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    'BODEGA
    If Index = 4 Then
                    BBodega = True
                    BMateriaPrima = False
                    DataConsultas.RecordSource = "BodegasMateriaPrima"
                    DataConsultas.Refresh
                    FrameConsultas.Visible = True
                    DBGridConsultas.SetFocus
                    DBGridConsultas.Columns(1).Width = "4000"
                    TxtBuscar.SetFocus
    'MATERIA PRIMA
    ElseIf Index = 2 Then
                    BBodega = False
                    BMateriaPrima = True
                    DataConsultas.RecordSource = "Select * From CorrelativosMateriaPrima"
                    DataConsultas.Refresh
                    FrameConsultas.Visible = True
                    DBGridConsultas.SetFocus
                    DBGridConsultas.Columns(1).Width = "4000"
                    TxtBuscar.SetFocus
    
    End If
End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
    TxtTexto.Item(Index).SelStart = 0
    TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
                'BODEGA
                If Index = 4 Then
                                BBodega = True
                                BMateriaPrima = False
                                DataConsultas.RecordSource = "BodegasMateriaPrima"
                                DataConsultas.Refresh
                                FrameConsultas.Visible = True
                                DBGridConsultas.SetFocus
                                DBGridConsultas.Columns(1).Width = "4000"
                                TxtBuscar.SetFocus
                'MATERIA PRIMA
                ElseIf Index = 2 Then
                                BBodega = False
                                BMateriaPrima = True
                                DataConsultas.RecordSource = "Select * From CorrelativosMateriaPrima"
                                DataConsultas.Refresh
                                FrameConsultas.Visible = True
                                DBGridConsultas.SetFocus
                                DBGridConsultas.Columns(1).Width = "4000"
                                TxtBuscar.SetFocus
                End If
    End If
End Sub

Private Sub List1_DblClick()
          DBGridDetalleMateriasPrimas.Columns(5).Text = Mid(List1.Text, 1, 1)
          List1.Visible = False
          DBGridDetalleMateriasPrimas.SetFocus

End Sub

Private Sub List1_KeyPress(KeyAscii As Integer)
            
      If KeyAscii = 43 Then
          DBGridDetalleMateriasPrimas.Columns(6).Text = Mid(List1.Text, 1, 1)
          List1.Visible = False
          DBGridDetalleMateriasPrimas.SetFocus
      End If
End Sub

Private Sub List1_LostFocus()
          List1.Visible = False
End Sub

