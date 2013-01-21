VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form DevolucionesMateriaPrimaProceso 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"DevolucionesMateriaPrimaProcesos.frx":0000
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11490
   Icon            =   "DevolucionesMateriaPrimaProcesos.frx":009D
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   11490
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
      Height          =   7695
      Left            =   120
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
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
         TabIndex        =   37
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
            TabIndex        =   38
            Top             =   360
            Value           =   -1  'True
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
            TabIndex        =   39
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   10440
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":0967
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataBuscar 
         Caption         =   "Productos"
         Connect         =   "Access"
         DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
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
         TabIndex        =   33
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1560
         TabIndex        =   34
         Top             =   360
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGridBuscar 
         Bindings        =   "DevolucionesMateriaPrimaProcesos.frx":0DA9
         Height          =   6615
         Left            =   120
         OleObjectBlob   =   "DevolucionesMateriaPrimaProcesos.frx":0DC2
         TabIndex        =   36
         ToolTipText     =   "Doble Click o Esc Para Seleccionar"
         Top             =   960
         Width           =   11055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   4935
      End
   End
   Begin MSDBGrid.DBGrid DBGridDetalleDevoluciones 
      Bindings        =   "DevolucionesMateriaPrimaProcesos.frx":179A
      Height          =   2775
      Left            =   240
      OleObjectBlob   =   "DevolucionesMateriaPrimaProcesos.frx":17C0
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   4200
      Width           =   11055
   End
   Begin VB.Data DataDetalleDevoluciones 
      Caption         =   "Detalle Traslados Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleDevolucionesMateriaPrima"
      Top             =   7680
      Visible         =   0   'False
      Width           =   4785
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.Frame FrameEncabezado 
      Caption         =   "Encabezado De Devouciones Materia Prima"
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
      Height          =   2295
      Left            =   120
      TabIndex        =   44
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   480
         Left            =   8280
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":312D
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   1300
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&EDITAR"
         Height          =   480
         Left            =   1680
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":365F
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   1300
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&USCAR"
         Height          =   480
         Left            =   6960
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":3B91
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   1300
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   9600
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":40C3
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salida"
         Top             =   1680
         Width           =   1300
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&BORRAR"
         Height          =   480
         Left            =   5640
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":6135
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   1300
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&CANCELAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   4320
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":6667
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   1300
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&GRABAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   3000
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":6B99
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   1300
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&AGREGAR"
         Height          =   480
         Left            =   360
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":70CB
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   1300
      End
      Begin VB.Frame FrameCompras 
         Enabled         =   0   'False
         Height          =   1335
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   10935
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            DataField       =   "Estado"
            DataSource      =   "DataEncabezadoDevoluciones"
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
            Left            =   9240
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            DataField       =   "Requerido"
            DataSource      =   "DataEncabezadoDevoluciones"
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
            Index           =   1
            Left            =   1680
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   600
            Width           =   1455
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataEncabezadoDevoluciones"
            Height          =   285
            Left            =   1680
            TabIndex        =   0
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataEncabezadoDevoluciones"
            Height          =   285
            Index           =   3
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   5
            Top             =   960
            Width           =   5655
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            DataField       =   "Liberado"
            DataSource      =   "DataEncabezadoDevoluciones"
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
            Left            =   5880
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TxtDocTra 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataEncabezadoDevoluciones"
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
            Left            =   5880
            MaxLength       =   15
            TabIndex        =   1
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Liberado Por"
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
            Left            =   4440
            TabIndex        =   54
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label6 
            Caption         =   "Requerido Por"
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
            Index           =   2
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   1455
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
            Left            =   8040
            TabIndex        =   52
            Top             =   240
            Width           =   1095
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
            TabIndex        =   51
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "No. Devolucion"
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
            Left            =   4440
            TabIndex        =   50
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Devolucion"
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
            TabIndex        =   49
            Top             =   240
            Width           =   1560
         End
      End
   End
   Begin VB.Frame FrameDetalle 
      Caption         =   "Detalle Devoluciones De Materia Prima"
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
      Height          =   5295
      Left            =   120
      TabIndex        =   40
      Top             =   2400
      Width           =   11325
      Begin VB.CommandButton CmdEditar2 
         Caption         =   "Editar"
         Height          =   495
         Left            =   2040
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":75FD
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   4680
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.CommandButton CmdBorrar2 
         Caption         =   "B&orrar"
         Height          =   495
         Left            =   7440
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":7B2F
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   4680
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.CommandButton CmdCancelar2 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5640
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":8061
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   4680
         Visible         =   0   'False
         Width           =   1750
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
         Height          =   495
         Left            =   9240
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":8593
         TabIndex        =   31
         Top             =   4680
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.CommandButton CmdGrabar2 
         Caption         =   "G&rabar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3840
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":8AC5
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4680
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.CommandButton CmdAgregar2 
         Caption         =   "A&gregar"
         Height          =   495
         Left            =   240
         Picture         =   "DevolucionesMateriaPrimaProcesos.frx":8FF7
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4680
         Visible         =   0   'False
         Width           =   1750
      End
      Begin VB.Frame FrameDetalleCompras 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   11055
         Begin VB.TextBox TxtUniMedSal 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Left            =   8400
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   480
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MskDifReqCorMas 
            DataField       =   "DiferenciaReqCorMas"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   4440
            TabIndex        =   21
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanRea 
            DataField       =   "CantidadReal"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   9720
            TabIndex        =   25
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskDifReqCor 
            DataField       =   "DiferenciaReqCor"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   5760
            TabIndex        =   22
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanDesPro 
            DataField       =   "CantidadDesperdicioProveedor"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   8400
            TabIndex        =   24
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanDes 
            DataField       =   "CantidadDesperdicio"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   7080
            TabIndex        =   23
            Top             =   1080
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanSal 
            DataField       =   "CantidadSalida"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   9720
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   480
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBodEnt 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "BodegaEntrada"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   3360
            MaxLength       =   3
            TabIndex        =   20
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox TxtNumIng 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            DataField       =   "NumeroIngreso"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   1800
            TabIndex        =   15
            Top             =   480
            Width           =   975
         End
         Begin VB.TextBox TxtBodSal 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            DataField       =   "BodegaSalida"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   2880
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox TxtDocDet 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   6000
            MaxLength       =   15
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox TxtCodSal 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            DataField       =   "CodigoSalida"
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   14
            Top             =   480
            Width           =   1575
         End
         Begin VB.TextBox TxtDesSal 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            DataSource      =   "DataDetalleDevoluciones"
            Height          =   285
            Left            =   3360
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   480
            Width           =   4935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Unidad Medida"
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
            Index           =   4
            Left            =   8280
            TabIndex        =   64
            Top             =   240
            Width           =   1290
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. Mas"
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
            Index           =   9
            Left            =   4440
            TabIndex        =   63
            Top             =   840
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cant. Real"
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
            Left            =   9720
            TabIndex        =   62
            Top             =   840
            Width           =   915
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cant.Menos"
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
            Left            =   5760
            TabIndex        =   61
            Top             =   840
            Width           =   1020
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Desp. Provee."
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
            Index           =   4
            Left            =   8400
            TabIndex        =   60
            Top             =   840
            Width           =   1230
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Desp. Proceso"
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
            Index           =   3
            Left            =   7080
            TabIndex        =   59
            Top             =   840
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Bod. Ent."
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
            Left            =   3360
            TabIndex        =   58
            Top             =   840
            Width           =   810
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "# Ingreso"
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
            Left            =   1800
            TabIndex        =   56
            Top             =   240
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Bod."
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
            Left            =   2880
            TabIndex        =   55
            Top             =   240
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad Salida"
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
            Left            =   9600
            TabIndex        =   48
            Top             =   240
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Left            =   3360
            TabIndex        =   47
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label Label1 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   615
         End
      End
   End
   Begin VB.Data DataEncabezadoDevoluciones 
      Caption         =   "Devoluciones De Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EncabezadoDevolucionesMateriaPrima"
      Top             =   7680
      Width           =   11295
   End
End
Attribute VB_Name = "DevolucionesMateriaPrimaProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As String
Dim VDocumentoDetalle As String
Dim VSumaEgresos As Double

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim BBodegaEntrada As Boolean
Dim BCodigoSalida As Boolean
Dim BNumeroIngreso As Boolean

Dim RBuscaMateriaPrimaSalida As Recordset
Dim RBuscaMateriaPrimaEntrada As Recordset
Dim RBuscaSigDoc As Recordset
Dim RBuscaDetalle As Recordset
Dim RBuscaEncabezado As Recordset
Dim RBuscaNumeroIngreso As Recordset

Dim VUltimoCodigo As String

Sub Botones1()
    If Bandera = True Then
         FrameCompras.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdSalida.Enabled = False
         CmdImprimir.Enabled = False
         DataEncabezadoDevoluciones.Visible = False
    Else
         FrameCompras.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdSalida.Enabled = True
         CmdImprimir.Enabled = True
         DataEncabezadoDevoluciones.Visible = True
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

Sub BotonesDetalleVisibles()
    If Bandera3 = True Then
         CmdAgregar2.Visible = True
         CmdEditar2.Visible = True
         CmdGrabar2.Visible = True
         CmdCancelar2.Visible = True
         CmdBorrar2.Visible = True
         CmdTerminar.Visible = True
    Else
         CmdAgregar2.Visible = False
         CmdEditar2.Visible = False
         CmdGrabar2.Visible = False
         CmdCancelar2.Visible = False
         CmdBorrar2.Visible = False
         CmdTerminar.Visible = False
    
    End If

End Sub

Private Sub CmdAgregar2_Click()
On Error Resume Next
    'AGREGA DATOS
    DataDetalleDevoluciones.Recordset.AddNew
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
    Bandera2 = True
    Botones2
    
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE EN EL GRID
    DBGridDetalleDevoluciones.Enabled = False
    'ASIGNA EL DOCUMENTO DEL ENCABEZADO AL DETALLE
    TxtDocDet.Text = VDocumento
    TxtCodSal.SetFocus
    TxtCodSal.Text = VUltimoCodigo
    
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            'REVISA EL ESTADO DEL TRASLADO
            If TxtEncabezado.Item(0).Text = "LIBERADO" Then
                MsgBox "Esta Devolucion No Se Puede Borrar Porque Ya Fue Liberada", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
            End If

            VDocumento = TxtDocTra.Text
            mensaje = MsgBox("¿Está Seguro De Borrar La Devolucion?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                        'BORRA EL ENCABEZADO DE EL PEDIDO
                        DataEncabezadoDevoluciones.Recordset.Delete
                        If Err <> 0 Then
                            'MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            'Exit Sub
                        End If
                        DataEncabezadoDevoluciones.Recordset.MoveLast
                MousePointer = 0
            End If
            If DataEncabezadoDevoluciones.Recordset.EOF Then
                DataEncabezadoDevoluciones.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
            VDocumentoDetalle = TxtDocDet.Text
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
            'SI CONTESTA QUE SI QUIERE BORRAR
            
            If mensaje = vbOK Then
                MousePointer = 11
                   
                   'BORRA EL DETALLE DE LA ENTRADA
                    DataDetalleDevoluciones.Recordset.Delete
                    
                    If Err <> 0 Then
                       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                       Exit Sub
                    End If
                    'SELECCIONA TODOS LOS DETALLES DE LA ENTRADAS
                    DataDetalleDevoluciones.RecordSource = ("Select * from DetalleDevolucionesMateriaPrima where Documento = '" & VDocumentoDetalle & "' order By BodegaSalida")
                    DataDetalleDevoluciones.Refresh
                    DBGridDetalleDevoluciones.Refresh
                MousePointer = 0
            End If
  
            If DataEncabezadoDevoluciones.Recordset.EOF Then
                DataEncabezadoDevoluciones.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
End Sub

Private Sub CmdBuscar_Click()
    mensaje = InputBox("No. Devolucion a Buscar")
    If mensaje = "" Then
    Else
        DataEncabezadoDevoluciones.Recordset.FindFirst ("Documento = '" & mensaje & "'")
    End If
    If Err <> 0 Then
    End If
    
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
    'CANCELA LOS CAMBIOS
    DataEncabezadoDevoluciones.Recordset.CancelUpdate
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error "
        Err.Clear
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    FrameDetalle.Visible = True
    DBGridDetalleDevoluciones.Visible = True
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next

    DataDetalleDevoluciones.Recordset.CancelUpdate
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    DBGridDetalleDevoluciones.Enabled = True
    Bandera2 = False
    Botones2

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
    'REVISA EL ESTADO DEL TRASLADO
    If TxtEncabezado.Item(0).Text = "LIBERADO" Then
        MsgBox "Esta Devolucion No Se Puede Editar Porque Ya Fue Liberada", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    'MODIFICA EL REGISTRO
    DataEncabezadoDevoluciones.Recordset.Edit
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    MskFec.SetFocus
    'ASIGNA AL CAMPO DE REQUERIDO EL USUARIO QUE LO ESTA EDITANDO
    TxtEncabezado.Item(1).Text = GUsuario
    'NO VIZUALIZA EL DETALLE
    FrameDetalle.Visible = False
    DBGridDetalleDevoluciones.Visible = False
End Sub


Private Sub CmdEditar2_Click()
On Error Resume Next
    'AGREGA DATOS
    DataDetalleDevoluciones.Recordset.Edit
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE EN EL GRID
    DBGridDetalleDevoluciones.Enabled = False
    
    Bandera2 = True
    Botones2
    
    DBGridDetalleDevoluciones.Enabled = False
    TxtNumIng.SetFocus

End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next

    'GUARDA EL ULTIMO CODIGO INGRESADO
    VUltimoCodigo = TxtCodSal.Text
        
    'BUSCA EL NUMERO DE INGRESO EN ENTRADAS DE BODEGA
    Set RBuscaNumeroIngreso = Db.OpenRecordset("Select NumeroIngreso From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCodSal.Text & "'")
    If RBuscaNumeroIngreso.RecordCount > 0 Then
    Else
        MsgBox "NUMERO INGRESO No Existe, En las Entradas De Recepcion De Bodega", vbOKOnly + vbInformation, "Informacion"
        TxtNumIng.SetFocus
        Exit Sub
    End If
    
    'BUSCA EL NUMERO DE INGRESO EN ENTRADAS DE BODEGA SI YA FUE LIBERADO
    Set RBuscaNumeroIngreso = Db.OpenRecordset("Select NumeroIngreso From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCodSal.Text & "'")
    If RBuscaNumeroIngreso.RecordCount > 0 Then
    Else
        MsgBox "Bulto No Ha Sido Liberado Por Recepcion De Bodega", vbOKOnly + vbInformation, "Informacion"
        TxtNumIng.SetFocus
        Exit Sub
    End If
    
    'REVISAMOS LA CANTIDAD DE SALIDA
    If Not IsNumeric(MskCanSal.Text) Then
       MsgBox "Cantidad De SALIDA Incorrecta", vbOKOnly + vbCritical, "Error"
       MskCanSal.SetFocus
       Exit Sub
    End If
    
    'REVISA LA CANTIDAD REQUISADA DE MENOS
    If Not IsNumeric(MskDifReqCor.Text) Then
       MsgBox "Cantidad De Diferencia Req/Cor Incorrecta", vbOKOnly + vbCritical, "Error"
       MskDifReqCor.SetFocus
       Exit Sub
    End If
    
    'REVISA LA CANTIDAD REQUISADA DE MAS
    If Not IsNumeric(MskDifReqCorMas.Text) Then
       MsgBox "Cantidad De Diferencia Req/Cor Incorrecta", vbOKOnly + vbCritical, "Error"
       MskDifReqCor.SetFocus
       Exit Sub
    End If
       
    'REVISA LA CANTIDAD REAL A TRASLADAR
    If Not IsNumeric(MskCanRea.Text) Then
       MsgBox "Cantidad Real Incorrecta", vbOKOnly + vbCritical, "Error"
       MskDifReqCor.SetFocus
       Exit Sub
    End If
    
        
    'REVISAMOS LA CANTIDAD DE DESPERDICIO
    If Not IsNumeric(MskCanDes.Text) Then
       MsgBox "Cantidad De DESPERDICIO De Proceso Incorrecta", vbOKOnly + vbCritical, "Error"
       MskCanDes.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS LA CANTIDAD DE DESPERDICIO
    If Not IsNumeric(MskCanDesPro.Text) Then
       MsgBox "Cantidad De DESPERDICIO De Proveedor Incorrecta", vbOKOnly + vbCritical, "Error"
       MskCanDes.SetFocus
       Exit Sub
    End If
    
    'VERIFICA BODEGA SALIDA CON BODEGA DE ENTRADA
    If TxtBodSal.Text = TxtBodEnt.Text Then
        MsgBox "La Bodega De Salida No Puede Ser Igual A La Bodega De Entrada", vbOKOnly + vbExclamation, "Informacion"
        TxtBodEnt.SetFocus
        Exit Sub
    End If
    
           
        
    'GRABA DATOS
    DataDetalleDevoluciones.Recordset.Update
        
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
        
    Bandera2 = False
    Botones2
         
    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE LA FACTURA QUE SE ESTA GRABANDO
    DataDetalleDevoluciones.RecordSource = ("Select * from DetalleDevolucionesMateriaPrima where Documento = '" & VDocumento & "' Order by BodegaSalida")
    DataDetalleDevoluciones.Refresh
    DBGridDetalleDevoluciones.Refresh
           
    DBGridDetalleDevoluciones.Enabled = True
    CmdAgregar2_Click
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
    
    'AGREGA UN REGISTRO
    DataEncabezadoDevoluciones.Recordset.AddNew
    
    'SI HAY ERRORES
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    
    'ASIGNA EL USUARIO A EL CAMPO DE REQUERIDO
    TxtEncabezado.Item(1).Text = GUsuario
    MskFec.Text = Date
    MskFec.SetFocus
    'ASIGNA EL ESTADO DE EL TRASLADO
    TxtEncabezado.Item(0).Text = "NO LIBERADO"
        
    'BUSCA EL MAXIMO DE DOCUMENTO Y LE SUMA 1
    'Set RBuscaSigDoc = Db.OpenRecordset("Select Max(Documento) from EncabezadoTrasladosMateriaPrimaP")
    '    If RBuscaSigDoc.RecordCount > 0 Then
    '        If IsNull(RBuscaSigDoc(0)) Then
    '            TxtDocTra.Text = "1"
    '        Else
    '            TxtDocTra.Text = Val(RBuscaSigDoc(0)) + 1
    '        End If
    '    End If

    FrameDetalle.Visible = False
    DBGridDetalleDevoluciones.Visible = False

End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    VDocumento = TxtDocTra.Text
        
    'GRABA DATOS
    DataEncabezadoDevoluciones.Recordset.Update
    
    If Err = 3022 Then
        MsgBox "Documento De Devolucion Ya Existe ", vbOKOnly + vbCritical, "Informacion"
        TxtDocTra.SetFocus
        Exit Sub
    ElseIf Err <> 0 And Err <> 3022 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    'SELECCIONA TODOS LOS DETALLES DE EL TRASLADO
    DataDetalleDevoluciones.RecordSource = ("Select * from DetalleDevolucionesMateriaPrima where Documento = '" & VDocumento & "' Order by BodegaSalida")
    DataDetalleDevoluciones.Refresh
    DBGridDetalleDevoluciones.Refresh
            
    'MUEVE EL RECORDSET A EL DOCUMENTO ACTUAL PARA QUE SE ACTUALIZEN LOS CAMBIOS
    DataEncabezadoDevoluciones.Recordset.FindFirst ("Documento = '" & VDocumento & "'")
    
    'VIZUALIZA EL DETALLE DE TRASLADO
    FrameDetalle.Visible = True
    
    'VISUALIZA TODOS LOS BOTONES DE DETALLE
    Bandera3 = True
    BotonesDetalleVisibles
    
    'NO VISUALIZA EL DATA DE ENCABEZADO DE TRASLADOS
    DataEncabezadoDevoluciones.Visible = False
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = True
    FrameDetalle.Visible = True
    FrameEncabezado.Enabled = False
    DBGridDetalleDevoluciones.Visible = True
    CmdAgregar2_Click
End Sub

Private Sub CmdImprimir_Click()
MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDocTra.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        CrReportes.SelectionFormula = "{EncabezadoDevolucionesMateriaPrima.Documento} = '" & TxtDocTra.Text & "'"
        CrReportes.ReportFileName = App.Path & "\FormatoDevolucionesMateriaPrima.rpt"
        CrReportes.Action = 1
MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdTerminar_Click()
If CmdCancelar2.Enabled = True Then
     CmdCancelar2_Click
End If
    
     
    'VISUALIZA EL DATA DE ENCABEZADO DE TRASLADOS
    DataEncabezadoDevoluciones.Visible = True
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = False
    'FrameDetalle.Visible = False
    FrameEncabezado.Enabled = True
    
    'VISUALIZA TODOS LOS BOTONES DE DETALLE
    Bandera3 = False
    BotonesDetalleVisibles

End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub


Private Sub DataDetalleDevoluciones_Reposition()
        'BUSCA LA MATERIA PRIMA DE ACUERDO A LA BODEGA DE SALIDA
        Set RBuscaMateriaPrimaSalida = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCodSal.Text & "'")
        If RBuscaMateriaPrimaSalida.RecordCount > 0 Then
            TxtDesSal.Text = RBuscaMateriaPrimaSalida!Descripcion
        Else
            TxtDesSal.Text = ""
        End If
End Sub

Private Sub DataEncabezadoDevoluciones_Reposition()

        'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
        DataDetalleDevoluciones.RecordSource = ("Select * from DetalleDevolucionesMateriaPrima where Documento = '" & TxtDocTra.Text & "' Order by BodegaSalida")
        DataDetalleDevoluciones.Refresh
        DBGridDetalleDevoluciones.Refresh

End Sub

Private Sub DBGridBuscar_DblClick()
    'BODEGA ENTRADA
    If BBodegaEntrada = True Then
        TxtBodEnt.Text = DBGridBuscar.Columns(0)
        TxtBodEnt.SetFocus
    'MATERIA PRIMA SALIDA
    ElseIf BCodigoSalida = True Then
        TxtCodSal.Text = DBGridBuscar.Columns(0)
        TxtCodSal.SetFocus
    'NUMERO INGRESO
    ElseIf BNumeroIngreso = True Then
        TxtNumIng.Text = DBGridBuscar.Columns(1)
        TxtNumIng.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
    'BODEGA ENTRADA
    If BBodegaEntrada = True Then
        TxtBodEnt.Text = DBGridBuscar.Columns(0)
        TxtBodEnt.SetFocus
    'MATERIA PRIMA SALIDA
    ElseIf BCodigoSalida = True Then
        TxtCodSal.Text = DBGridBuscar.Columns(0)
        TxtCodSal.SetFocus
    'NUMERO INGRESO
    ElseIf BNumeroIngreso = True Then
        TxtNumIng.Text = DBGridBuscar.Columns(1)
        TxtNumIng.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End If

End Sub



Private Sub Form_Activate()
    
        'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
        DataDetalleDevoluciones.RecordSource = ("Select * from DetalleDevolucionesMateriaPrima where Documento = '" & TxtDocTra.Text & "' Order by BodegaSalida")
        DataDetalleDevoluciones.Refresh
        DBGridDetalleDevoluciones.Refresh
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
    

End Sub

Private Sub Form_Load()
        
        DataEncabezadoDevoluciones.Connect = GConnect
        DataDetalleDevoluciones.Connect = GConnect
        DataBuscar.Connect = GConnect
        
        DataEncabezadoDevoluciones.DatabaseName = BasedeDatos
        DataDetalleDevoluciones.DatabaseName = BasedeDatos
        DataBuscar.DatabaseName = BasedeDatos
    
    

        'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
        DataDetalleDevoluciones.RecordSource = ("Select * from DetalleDevolucionesMateriaPrima where Documento = '" & TxtDocTra.Text & "' Order by BodegaSalida")
        DataDetalleDevoluciones.Refresh
        DBGridDetalleDevoluciones.Refresh


   
End Sub



Private Sub MskCanDes_GotFocus()
    MskCanDes.SelStart = 0
    MskCanDes.SelLength = Len(MskCanDes.Text)
End Sub

Private Sub MskCanDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskCanDesPro_GotFocus()
    MskCanDesPro.SelStart = 0
    MskCanDesPro.SelLength = Len(MskCanDesPro.Text)
End Sub

Private Sub MskCanDesPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskCanDesPro_LostFocus()
    VSumaEgresos = Val(MskDifReqCor.Text) + Val(MskCanDes.Text) + Val(MskCanDesPro.Text)
    MskCanRea.Text = ((Val(MskCanSal.Text) + Val(MskDifReqCorMas.Text)) - VSumaEgresos)
End Sub



Private Sub MskCanRea_GotFocus()
    MskCanRea.SelStart = 0
    MskCanRea.SelLength = Len(MskCanRea.Text)
End Sub

Private Sub MskCanRea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskCanSal_GotFocus()
    MskCanSal.SelStart = 0
    MskCanSal.SelLength = Len(MskCanSal.Text)
End Sub

Private Sub MskCanSal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskDifReqCor_GotFocus()
    MskDifReqCor.SelStart = 0
    MskDifReqCor.SelLength = Len(MskDifReqCor.Text)
End Sub

Private Sub MskDifReqCor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskDifReqCorMas_GotFocus()
    MskDifReqCorMas.SelStart = 0
    MskDifReqCorMas.SelLength = Len(MskDifReqCorMas.Text)
End Sub

Private Sub MskDifReqCorMas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

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

Private Sub TxtBodEnt_DblClick()
        BBodegaEntrada = True
        BCodigoSalida = False
        BNumeroIngreso = False
        DataBuscar.RecordSource = "Select * From BodegasMateriaPrima"
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
End Sub

Private Sub TxtBodEnt_GotFocus()
    TxtBodEnt.SelStart = 0
    TxtBodEnt.SelLength = Len(TxtBodEnt.Text)
End Sub

Private Sub TxtBodEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        BBodegaEntrada = True
        BCodigoSalida = False
        BNumeroIngreso = False
        DataBuscar.RecordSource = "Select * From BodegasMateriaPrima"
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
    End If
End Sub
Private Sub TxtBuscar_Change()
    'BODEGA ENTRADA
    If BBodegaEntrada = True Then
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
        
    'CODIGO MATERIA PRIMA SALIDA
    ElseIf BCodigoSalida = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from CorrelativosMateriaPrima Where CodigoMateriaPrima Like '" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
            Else
                    DataBuscar.RecordSource = ("Select * from CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from CorrelativosMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
            Else
                    DataBuscar.RecordSource = ("Select * from CorrelativosMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
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




Private Sub TxtCodSal_Change()
        'BUSCA LA MATERIA PRIMA DE ACUERDO A LA BODEGA DE SALIDA
        Set RBuscaMateriaPrimaSalida = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCodSal.Text & "'")
        If RBuscaMateriaPrimaSalida.RecordCount > 0 Then
            TxtDesSal.Text = RBuscaMateriaPrimaSalida!Descripcion
            TxtUniMedSal.Text = RBuscaMateriaPrimaSalida!UnidadMedida
        Else
            TxtDesSal.Text = ""
            TxtUniMedSal.Text = ""
        End If
End Sub

Private Sub TxtCodSal_DblClick()
        BBodegaEntrada = False
        BCodigoSalida = True
        BNumeroIngreso = False
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataBuscar.RecordSource = "Select * From CorrelativosMateriaPrima"
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
End Sub

Private Sub TxtCodSal_GotFocus()
    TxtCodSal.SelStart = 0
    TxtCodSal.SelLength = Len(TxtCodSal.Text)
End Sub

Private Sub TxtCodSal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        BBodegaEntrada = False
        BCodigoSalida = True
        BNumeroIngreso = False
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataBuscar.RecordSource = "Select * FROM CorrelativosMateriaPrima"
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
    End If
End Sub

Private Sub TxtDocTra_GotFocus()
    TxtDocTra.SelStart = 0
    TxtDocTra.SelLength = Len(TxtDocTra.Text)
End Sub

Private Sub TxtDocTra_KeyPress(KeyAscii As Integer)
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

Private Sub TxtNumIng_DblClick()
        BBodegaEntrada = False
        BCodigoSalida = False
        BNumeroIngreso = True
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataBuscar.RecordSource = "Select BodegaDisponibilidad, NumeroIngreso, CantidadTraslado, CantidadSalida, SaldoDisponibilidad From DetalleEntradasMateriaPrima Where Codigo = '" & TxtCodSal.Text & "' And SaldoDisponibilidad > 0 Order By BodegaDisponibilidad, NumeroIngreso"
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        Columnas
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
End Sub

Private Sub TxtNumIng_GotFocus()
    TxtNumIng.SelStart = 0
    TxtNumIng.SelLength = Len(TxtNumIng.Text)
End Sub

Private Sub TxtNumIng_KeyPress(KeyAscii As Integer)
    'SI PRECIONA ENTER
    If KeyAscii = 13 Then
        If IsNumeric(TxtNumIng.Text) Then
            'BUSCA EL NUMERO DE INGRESO Y ASIGNA LA BODEGA, CODIGO Y CANTIDAD DE ACUERDO COMO ENTRO A LA BODEGA
            Set RBuscaNumeroIngreso = Db.OpenRecordset("Select SaldoDisponibilidad, Codigo, BodegaDisponibilidad From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCodSal.Text & "'")
            'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
            If RBuscaNumeroIngreso.RecordCount > 0 Then
                MskCanSal.Text = RBuscaNumeroIngreso!SaldoDisponibilidad
                TxtCodSal.Text = RBuscaNumeroIngreso!Codigo
                If Not IsNull(RBuscaNumeroIngreso!BodegaDisponibilidad) Then
                    TxtBodSal.Text = RBuscaNumeroIngreso!BodegaDisponibilidad
                Else
                    TxtBodSal.Text = ""
                End If
            'SI NO ENCUENTRA EL INGRESO DEJA EN BLANCO
            Else
                MskCanSal.Text = 0
                TxtCodSal.Text = ""
                TxtBodSal.Text = ""
                TxtBodEnt.Text = ""
            End If
        End If
       SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        BBodegaEntrada = False
        BCodigoSalida = False
        BNumeroIngreso = True
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataBuscar.RecordSource = "Select BodegaDisponibilidad, NumeroIngreso, CantidadTraslado, CantidadSalida, SaldoDisponibilidad From DetalleEntradasMateriaPrima Where Codigo = '" & TxtCodSal.Text & "' And SaldoDisponibilidad > 0 Order By BodegaDisponibilidad, NumeroIngreso"
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        Columnas
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
    End If
End Sub


Private Sub TxtUniMedSal_GotFocus()
    TxtUniMedSal.SelStart = 0
    TxtUniMedSal.SelLength = Len(TxtUniMedSal.Text)
End Sub

Private Sub TxtUniMedSal_KeyPress(KeyAscii As Integer)
    'SI PRECIONA ENTER
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub


Sub Columnas()
    DBGridBuscar.Columns(0).Caption = "Bodega"
    DBGridBuscar.Columns(0).Width = "500"
    DBGridBuscar.Columns(1).Caption = "# Ingreso"
    DBGridBuscar.Columns(1).Width = "1000"
    DBGridBuscar.Columns(2).Caption = "Inicio"
    DBGridBuscar.Columns(3).Caption = "Salidas"
    DBGridBuscar.Columns(4).Caption = "Existencia"
    
End Sub

