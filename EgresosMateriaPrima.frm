VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form EgresosMateriaPrima 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"EgresosMateriaPrima.frx":0000
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "EgresosMateriaPrima.frx":00A3
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
      Caption         =   "Buscar Datos"
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
      TabIndex        =   44
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
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
         Left            =   6480
         TabIndex        =   54
         Top             =   240
         Width           =   4335
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
            TabIndex        =   40
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
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
            TabIndex        =   39
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   10920
         Picture         =   "EgresosMateriaPrima.frx":04E5
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataBuscar 
         Caption         =   "Productos"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   37
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1680
         TabIndex        =   38
         Top             =   360
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGridBuscar 
         Bindings        =   "EgresosMateriaPrima.frx":2557
         Height          =   7455
         Left            =   120
         OleObjectBlob   =   "EgresosMateriaPrima.frx":2570
         TabIndex        =   36
         ToolTipText     =   "Doble Click o Esc Para Seleccionar"
         Top             =   960
         Width           =   11535
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   600
         Width           =   6255
      End
   End
   Begin MSDBGrid.DBGrid DBGridDetalleRequisiciones 
      Bindings        =   "EgresosMateriaPrima.frx":2F48
      Height          =   2175
      Left            =   120
      OleObjectBlob   =   "EgresosMateriaPrima.frx":2F69
      TabIndex        =   58
      Top             =   5160
      Width           =   11535
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Connect         =   "pwd=metal"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame FrameEncabezado 
      Caption         =   "Encabezado de Salidas"
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
      Height          =   3495
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton CmdBuscar2 
         Caption         =   "Siguiente Documento"
         Height          =   480
         Left            =   8040
         TabIndex        =   22
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   480
         Left            =   9360
         Picture         =   "EgresosMateriaPrima.frx":3E67
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   480
         Left            =   1440
         Picture         =   "EgresosMateriaPrima.frx":4399
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&uscar Documento"
         Height          =   480
         Left            =   6720
         Picture         =   "EgresosMateriaPrima.frx":48CB
         TabIndex        =   21
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   10680
         Picture         =   "EgresosMateriaPrima.frx":4DFD
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Salida"
         Top             =   2880
         Width           =   960
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&Borrar"
         Height          =   480
         Left            =   5400
         Picture         =   "EgresosMateriaPrima.frx":6E6F
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   480
         Left            =   4080
         Picture         =   "EgresosMateriaPrima.frx":73A1
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   480
         Left            =   2760
         Picture         =   "EgresosMateriaPrima.frx":78D3
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2880
         Width           =   1200
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   480
         Left            =   120
         Picture         =   "EgresosMateriaPrima.frx":7E05
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Frame FrameRequisiciones 
         Enabled         =   0   'False
         Height          =   2535
         Left            =   120
         TabIndex        =   48
         Top             =   240
         Width           =   11535
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "TipoDeDocumento"
            DataSource      =   "DataEgresos"
            Height          =   285
            Index           =   1
            Left            =   4440
            MaxLength       =   10
            TabIndex        =   3
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "PlacasFurgon"
            DataSource      =   "DataEgresos"
            Height          =   285
            Index           =   8
            Left            =   10080
            MaxLength       =   10
            TabIndex        =   12
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "PlacasCamion"
            DataSource      =   "DataEgresos"
            Height          =   285
            Index           =   7
            Left            =   10080
            MaxLength       =   10
            TabIndex        =   11
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "Conductor"
            DataSource      =   "DataEgresos"
            Height          =   285
            Index           =   6
            Left            =   10080
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1440
            Width           =   1335
         End
         Begin MSMask.MaskEdBox MskCos 
            DataField       =   "CostoViaje"
            DataSource      =   "DataEgresos"
            Height          =   285
            Left            =   7320
            TabIndex        =   8
            Top             =   1800
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "CargadoPor"
            DataSource      =   "DataEgresos"
            Height          =   285
            Index           =   3
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   6
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "EntregadoPor"
            DataSource      =   "DataEgresos"
            Height          =   285
            Index           =   5
            Left            =   4440
            MaxLength       =   10
            TabIndex        =   7
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "CodigoTransportista"
            DataSource      =   "DataEgresos"
            Height          =   285
            Index           =   0
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   5
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "NumeroDocumento"
            DataSource      =   "DataEgresos"
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
            Index           =   2
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtCli 
            Appearance      =   0  'Flat
            DataField       =   "Cliente"
            DataSource      =   "DataEgresos"
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   4
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox TxtEst 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            DataField       =   "Estado"
            DataSource      =   "DataEgresos"
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
            Left            =   10080
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtLib 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Liberado"
            DataSource      =   "DataEgresos"
            Height          =   285
            Left            =   10080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataEgresos"
            Height          =   285
            Index           =   4
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   9
            Top             =   2160
            Width           =   6975
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataEgresos"
            Height          =   285
            Left            =   1560
            TabIndex        =   0
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtReq 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Requerido"
            DataSource      =   "DataEgresos"
            Height          =   285
            Left            =   10080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtDoc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Documento"
            DataSource      =   "DataEgresos"
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
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label LblDoc 
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
            Left            =   5880
            TabIndex        =   74
            Top             =   720
            Width           =   2655
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Placas Furgon"
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
            Index           =   14
            Left            =   8760
            TabIndex        =   73
            Top             =   2160
            Width           =   1230
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Placas Camion"
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
            Index           =   13
            Left            =   8760
            TabIndex        =   72
            Top             =   1800
            Width           =   1260
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   8
            Left            =   8760
            TabIndex        =   71
            Top             =   1440
            Width           =   885
         End
         Begin VB.Label LblTransportista 
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
            Left            =   3000
            TabIndex        =   70
            Top             =   1440
            Width           =   5535
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Costo Viaje"
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
            Left            =   5880
            TabIndex        =   69
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Cargado Por"
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
            Index           =   11
            Left            =   120
            TabIndex        =   68
            Top             =   1800
            Width           =   1065
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Entregado Por"
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
            Index           =   10
            Left            =   3000
            TabIndex        =   67
            Top             =   1800
            Width           =   1230
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
            Index           =   9
            Left            =   120
            TabIndex        =   66
            Top             =   1440
            Width           =   1125
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "No. Documento"
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
            Index           =   7
            Left            =   120
            TabIndex        =   65
            Top             =   720
            Width           =   1335
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
            Index           =   6
            Left            =   3000
            TabIndex        =   64
            Top             =   720
            Width           =   1410
         End
         Begin VB.Label Label6 
            Caption         =   "Cliente"
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
            TabIndex        =   63
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label LblCliente 
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
            Left            =   3000
            TabIndex        =   62
            Top             =   1080
            Width           =   5535
         End
         Begin VB.Label Label6 
            Caption         =   "Estado"
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
            Index           =   5
            Left            =   8760
            TabIndex        =   59
            Top             =   360
            Width           =   735
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
            Index           =   4
            Left            =   8760
            TabIndex        =   57
            Top             =   1080
            Width           =   1455
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
            Index           =   3
            Left            =   8760
            TabIndex        =   56
            Top             =   720
            Width           =   1455
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
            Index           =   2
            Left            =   120
            TabIndex        =   55
            Top             =   2160
            Width           =   1335
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
            Left            =   3000
            TabIndex        =   53
            Top             =   360
            Width           =   1065
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Width           =   855
         End
      End
   End
   Begin VB.Frame FrameDetalle 
      Caption         =   "Detalle de Salidas"
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
      Height          =   4455
      Left            =   0
      TabIndex        =   41
      Top             =   3600
      Width           =   11805
      Begin VB.CommandButton CmdEditar2 
         Caption         =   "Editar"
         Height          =   495
         Left            =   2040
         Picture         =   "EgresosMateriaPrima.frx":8337
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3840
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton CmdBorrar2 
         Caption         =   "B&orrar"
         Height          =   495
         Left            =   7800
         Picture         =   "EgresosMateriaPrima.frx":8869
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3840
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton CmdCancelar2 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5880
         Picture         =   "EgresosMateriaPrima.frx":8D9B
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3840
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
         Height          =   495
         Left            =   9720
         Picture         =   "EgresosMateriaPrima.frx":92CD
         TabIndex        =   34
         Top             =   3840
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CommandButton CmdGrabar2 
         Caption         =   "G&rabar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3960
         Picture         =   "EgresosMateriaPrima.frx":97FF
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   3840
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton CmdAgregar2 
         Caption         =   "A&gregar"
         Height          =   495
         Left            =   120
         Picture         =   "EgresosMateriaPrima.frx":9D31
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3840
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Frame FrameDetalleRequisiciones 
         Enabled         =   0   'False
         Height          =   1215
         Left            =   120
         TabIndex        =   42
         Top             =   240
         Width           =   11535
         Begin VB.TextBox TxtNumIng 
            Appearance      =   0  'Flat
            DataField       =   "NumeroIngreso"
            DataSource      =   "DataDetalleEgresos"
            Height          =   285
            Left            =   8880
            TabIndex        =   26
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   480
            Width           =   1335
         End
         Begin MSMask.MaskEdBox MskCanMatPri 
            DataField       =   "Cantidad"
            DataSource      =   "DataDetalleEgresos"
            Height          =   285
            Left            =   10320
            TabIndex        =   28
            Top             =   480
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Bodega"
            DataSource      =   "DataDetalleEgresos"
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox TxtDocDet 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataDetalleEgresos"
            Height          =   285
            Left            =   5160
            MaxLength       =   15
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtCodMatPri 
            Appearance      =   0  'Flat
            DataField       =   "Codigo"
            DataSource      =   "DataDetalleEgresos"
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   25
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   480
            Width           =   1695
         End
         Begin VB.TextBox TxtDesPro 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   480
            Width           =   6855
         End
         Begin VB.Label Label1 
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
            TabIndex        =   76
            Top             =   840
            Width           =   735
         End
         Begin VB.Label LblBod 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
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
            Left            =   1920
            TabIndex        =   75
            Top             =   840
            Width           =   6855
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   2
            Left            =   3360
            TabIndex        =   61
            Top             =   240
            Width           =   660
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   1
            Left            =   8880
            TabIndex        =   60
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Left            =   10320
            TabIndex        =   51
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   1920
            TabIndex        =   50
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label1 
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
            TabIndex        =   49
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Data DataDetalleEgresos 
      Caption         =   "Detalle Ingresos"
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
      RecordSource    =   "DetalleEgresosMateriaPrima"
      Top             =   8160
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Data DataEgresos 
      Caption         =   "Salidas"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EncabezadoEgresosMateriaPrima"
      Top             =   8145
      Width           =   11190
   End
End
Attribute VB_Name = "EgresosMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As Long
Dim VDocumentoDetalle As Long
Dim VCantidadMateriaPrima As Double
Dim VCodigoMateriaPrima As String
Dim VBodega As String
Dim VNumeroPedido As String
Dim VCliente As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim BEditar As Boolean
Dim BMateriaPrima As Boolean
Dim BNumeroIngreso As Boolean
Dim BCliente As Boolean
Dim BTransportista As Boolean
Dim BPedido As Boolean
Dim BDocumento As Boolean

Dim RBuscaMateriaPrima As Recordset
Dim RBuscaNumeroIngreso As Recordset
Dim RBuscaCliente As Recordset
Dim RBuscaSigDoc As Recordset
Dim RBuscaDetalle As Recordset
Dim RBuscaBodega As Recordset
Dim RBuscaTransportista As Recordset
Dim RBuscaTipoDocumento As Recordset
Dim RBuscaDocumento As Recordset

Sub Botones1()
    If Bandera = True Then
         FrameRequisiciones.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdBuscar2.Enabled = False
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         DataEgresos.Visible = False
    Else
         FrameRequisiciones.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdBuscar2.Enabled = True
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         DataEgresos.Visible = True
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleRequisiciones.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdEditar2.Enabled = False
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalleRequisiciones.Enabled = False
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




Private Sub CmdAgregar2_Click()
On Error Resume Next
    'AGREGA UN REGISTRO DE DETALLE
    DataDetalleEgresos.Recordset.AddNew
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description
    End If
        
    Bandera2 = True
    Botones2

    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DBGridDetalleRequisiciones.Enabled = False
    
    'SE ASIGNA AL DOCUMENTO DE DETALLE EL DOCUMENTO DEL ENCABEZADO
    TxtDocDet.Text = VDocumento
    TxtCodMatPri.SetFocus
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            If GBorrar = True Then
                'NO HACE NADA PORQUE SI TIENE ACCESO
            ElseIf TxtEst.Text = "LIBERADO" Then
                'REVISA SI YA FUE LIBERADO EL INGRESO
                MsgBox "No Puede BORRAR Este Ingreso Porque Ya Fue Liberado", vbOKOnly + vbCritical, "Informacion"
                Exit Sub
            End If

            VDocumento = TxtDoc.Text

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                
                'BORRA EL ENCABEZADO DE LA FACTURA
                DataEgresos.Recordset.Delete
                DataEgresos.Recordset.MoveLast
                MousePointer = 0
            End If
  
            If DataEgresos.Recordset.EOF Then
                DataEgresos.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next

    'ASIGANMOS A UNA VARIABLE EL DOCUMENTO DETALLE
    VDocumentoDetalle = TxtDocDet.Text
    VCodigoMateriaPrima = TxtCodMatPri.Text
    VCantidadMateriaPrima = MskCanMatPri.Text

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                        
                   'BORRA EL DETALLE DE LA FACTURA
                    DataDetalleEgresos.Recordset.Delete
                    
                        If Err <> 0 Then
                            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                        End If
                                        
                    'SELECCIONA TODOS LOS DETALLES DE LA FACTURA
                    DataDetalleEgresos.RecordSource = ("Select * from DetalleEgresosMateriaPrima where documento = " & VDocumentoDetalle & " order By Codigo")
                    DataDetalleEgresos.Refresh
                    DBGridDetalleRequisiciones.Refresh
                MousePointer = 0
            End If
  
            If DataEgresos.Recordset.EOF Then
                DataEgresos.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
    mensaje = InputBox("Numero De Documento A Buscar")
    DataEgresos.Recordset.FindFirst ("NumeroDocumento = '" & mensaje & "'")
    If Err <> 0 Then
    End If
End Sub

Private Sub CmdBuscar2_Click()
On Error Resume Next
    DataEgresos.Recordset.FindNext ("NumeroDocumento = '" & mensaje & "'")
    If Err <> 0 Then
    End If

End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next

    'GRABA DATOS
    DataEgresos.Recordset.CancelUpdate
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    FrameDetalle.Visible = True
    DBGridDetalleRequisiciones.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    'CANCELA EL INGRESO
    DataDetalleEgresos.Recordset.CancelUpdate
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    DBGridDetalleRequisiciones.Enabled = True
    Bandera2 = False
    Botones2

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
    If GEditar = True Then
        'NO HACE NADA PORQUE SI TIENE ACCESO
    ElseIf TxtEst.Text = "LIBERADO" Then
        'REVISA SI YA FUE LIBERADO EL INGRESO
        MsgBox "No Puede EDITAR Este Ingreso Porque Ya Fue Liberado", vbOKOnly + vbCritical, "Informacion"
        Exit Sub
    End If

    'EDITA EL REGISTRO ACTUAL
    DataEgresos.Recordset.Edit
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    MskFec.SetFocus
    BEditar = True
    
    FrameDetalle.Visible = False
    DBGridDetalleRequisiciones.Visible = False
End Sub


Private Sub CmdEditar2_Click()
On Error Resume Next
    'EDITA EL REGISTRO ACTUAL
    DataDetalleEgresos.Recordset.Edit
    
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DBGridDetalleRequisiciones.Enabled = False
    
    Bandera2 = True
    Botones2
    TxtCodMatPri.SetFocus
End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
    
    'GUARDA VARIABLES
    VCantidadMateriaPrima = MskCanMatPri.Text
    VCodigoMateriaPrima = TxtCodMatPri.Text
    VBodega = TxtBod.Text
            
    'REVISAMOS DATOS
    If Not IsNumeric(MskCanMatPri.Text) Then
       MsgBox "Cantidad De Materia Prima Incorrecta", vbOKOnly + vbCritical, "Error"
       MskCanMatPri.SetFocus
       Exit Sub
    End If
    
    'BUSCA EL NUMERO DE INGRESO
    Set RBuscaNumeroIngreso = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & VCodigoMateriaPrima & "'")
    'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
    If RBuscaNumeroIngreso.RecordCount > 0 Then
    Else
        MsgBox "El Numero De Ingreso Para Esta Materia Prima, No Exite, Revice Las Entradas a Bodega", vbOKOnly + vbExclamation, "Informacion"
        Exit Sub
    End If
    
        
    'GRABA DATOS
    DataDetalleEgresos.Recordset.Update
        
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera2 = False
    Botones2
      
    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE LA FACTURA QUE SE ESTA GRABANDO
    DataDetalleEgresos.RecordSource = ("Select * from DetalleEgresosMateriaPrima where Documento = " & VDocumento & " Order BY Codigo")
    DataDetalleEgresos.Refresh
    DBGridDetalleRequisiciones.Refresh
        
    DBGridDetalleRequisiciones.Enabled = True
    CmdAgregar2.SetFocus
End Sub


Private Sub CmdAgregar_Click()
    On Error Resume Next
    'AGREGA UN REGISTRO
    DataEgresos.Recordset.AddNew
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    BEditar = False
    
    'ASIGNA EL USUARIO
    TxtReq.Text = GUsuario
    'ASIGNA EL ESTADO DEL INGRESO
    TxtEst.Text = "NO LIBERADO"
    'ASIGNA LA FECHA DEL DIA
    MskFec.Text = Date
    MskFec.SetFocus
    
    FrameDetalle.Visible = False
    DBGridDetalleRequisiciones.Visible = False
    
    'BUSCA EL DOCUMENTO MAXIMO Y LE ASIGNA 1
    Set RBuscaSigDoc = Db.OpenRecordset("Select Max(Documento) from EncabezadoEgresosMateriaPrima")
        If RBuscaSigDoc.RecordCount > 0 Then
            If IsNull(RBuscaSigDoc(0)) Then
                TxtDoc.Text = "1"
            Else
                TxtDoc.Text = RBuscaSigDoc(0) + 1
            End If
        End If
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    
'OSEA QUE SI ESTA AGREGANDO UN REGISTRO
    If BEditar = False Then
            'BUSCA SI YA EXISTE EL NUMERO DE DOCUMENTO PARA ESTE TIPO DE DOCUMENTO
            Set RBuscaDocumento = Db.OpenRecordset("Select * From EncabezadoEgresosMateriaPrima Where TipoDeDocumento = '" & TxtEncabezado.Item(2).Text & "' And NumeroDocumento = '" & TxtEncabezado.Item(1).Text & "'")
                    If RBuscaDocumento.RecordCount > 0 Then
                        MsgBox "Numero Documento Para Este Tipo De Documento Ya Existe", vbOKOnly + vbInformation, "Informacion"
                        TxtEncabezado.Item(2).SetFocus
                        Exit Sub
                    End If
    End If
    
    VDocumento = TxtDoc.Text
    VCliente = TxtCli.Text
    
    If Not IsNumeric(MskCos.Text) Then
            MsgBox "Costo Debe Der Numerico"
            Exit Sub
    End If
    
    'VALIDA TRANSPORTISTA
    If TxtEncabezado.Item(0).Text = "" Then
            MsgBox "Transportista No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
            TxtEncabezado.Item(0).SetFocus
            Exit Sub
    End If
    
    'VALIDA TIPO DE DOCUMENTO
    If TxtEncabezado.Item(1).Text = "" Then
            MsgBox "Tipo De Documento No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
            TxtEncabezado.Item(1).SetFocus
            Exit Sub
    End If
    
    'VALIDA CLIENTE
    If TxtCli.Text = "" Then
            MsgBox "Cliente No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
            TxtCli.SetFocus
            Exit Sub
    End If
    
    
    
    'GRABA DATOS
    DataEgresos.Recordset.Update
    
    'SI SE DUPLICA LA LLAVE
    If Err = 3022 Then
        MsgBox "Numero De Documento Ya Existe", vbOKOnly + vbCritical, "Error"
        TxtDoc.SetFocus
        Exit Sub
    ElseIf Err <> 0 And Err <> 3022 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    'SELECCIONA TODO EL DETALLE DE EL INGRESO
    DataDetalleEgresos.RecordSource = ("Select * from DetalleEgresosMateriaPrima where Documento = " & VDocumento & " Order by Codigo")
    DataDetalleEgresos.Refresh
    DBGridDetalleRequisiciones.Refresh
    
    'MUEVE EL RECORDSET A LA FACTURA ACTUAL PARA QUE SE ACTUALIZEN LOS CAMBIOS
    DataEgresos.Recordset.FindFirst ("Documento = " & VDocumento)
    
    'ESCONDE LOS BOTONES DEL ENCABEZADO
    Bandera4 = False
    BotonesVisiblesEncabezado
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Visible = True
    DBGridDetalleRequisiciones.Visible = True
    FrameDetalle.Enabled = True
    FrameEncabezado.Enabled = False
    CmdAgregar2.SetFocus
End Sub

Private Sub CmdImprimir_Click()
MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDoc.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
        
                CrReportes.SelectionFormula = "{EncabezadoEgresosMateriaPrima.Documento} = " & TxtDoc.Text
                CrReportes.ReportFileName = App.Path & "\ReporteSalidasMateriaPrima.rpt"
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
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Visible = True
    FrameEncabezado.Enabled = True
           
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle
    
    'ESCONDE LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesVisiblesEncabezado

End Sub

Private Sub Command1_Click()
    Framebuscar.Visible = False
End Sub


Private Sub Command2_Click()

End Sub

Private Sub DataEgresos_Reposition()
    If IsNumeric(TxtDoc.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL INGRESO
        DataDetalleEgresos.RecordSource = ("Select * from DetalleEgresosMateriaPrima where Documento = " & TxtDoc.Text & " Order by Codigo")
        DataDetalleEgresos.Refresh
        DBGridDetalleRequisiciones.Refresh
    End If


End Sub

Private Sub DBGridBuscar_DblClick()
        If BMateriaPrima = True Then
            TxtCodMatPri.Text = DBGridBuscar.Columns(0)
            TxtCodMatPri.SetFocus
        ElseIf BNumeroIngreso = True Then
            TxtNumIng.Text = DBGridBuscar.Columns(1)
            TxtNumIng.SetFocus
        ElseIf BCliente = True Then
            TxtCli.Text = DBGridBuscar.Columns(0)
            TxtCli.SetFocus
        ElseIf BTransportista = True Then
            TxtEncabezado.Item(0).Text = DBGridBuscar.Columns(0)
            TxtEncabezado.Item(0).SetFocus
        ElseIf BDocumento = True Then
            TxtEncabezado.Item(1).Text = DBGridBuscar.Columns(0)
            TxtEncabezado.Item(1).SetFocus
        End If
            TxtBuscar.Text = ""
            Framebuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BMateriaPrima = True Then
                TxtCodMatPri.Text = DBGridBuscar.Columns(0)
                TxtCodMatPri.SetFocus
            ElseIf BNumeroIngreso = True Then
                TxtNumIng.Text = DBGridBuscar.Columns(1)
                TxtNumIng.SetFocus
            ElseIf BCliente = True Then
                TxtCli.Text = DBGridBuscar.Columns(0)
                TxtCli.SetFocus
            ElseIf BTransportista = True Then
                TxtEncabezado.Item(0).Text = DBGridBuscar.Columns(0)
                TxtEncabezado.Item(0).SetFocus
            ElseIf BDocumento = True Then
                TxtEncabezado.Item(1).Text = DBGridBuscar.Columns(0)
                TxtEncabezado.Item(1).SetFocus
            End If
            TxtBuscar.Text = ""
            Framebuscar.Visible = False
        End If
End Sub



Private Sub Form_Activate()
    If IsNumeric(TxtDoc.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL INGRESO
        DataDetalleEgresos.RecordSource = ("Select * from DetalleEgresosMateriaPrima where Documento = " & TxtDoc.Text & " Order by Codigo")
        DataDetalleEgresos.Refresh
        DBGridDetalleRequisiciones.Refresh
    End If
    

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
    

End Sub

Private Sub Form_Load()
    DataEgresos.ConnectionString = GTipoProveedor
    DataDetalleEgresos.ConnectionString = GTipoProveedor
    DataBuscar.ConnectionString = GTipoProveedor
    
    DataEgresos.Refresh
    DataDetalleEgresos.Refresh
    DataBuscar.Refresh
    
    
    
End Sub




Private Sub MskCanMatPri_GotFocus()
        MskCanMatPri.SelStart = 0
        MskCanMatPri.SelLength = Len(MskCanMatPri.Text)
End Sub

Private Sub MskCanMatPri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub MskCos_GotFocus()
        MskCos.SelStart = 0
        MskCos.SelLength = Len(MskCos.Text)
End Sub

Private Sub MskCos_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
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


Private Sub TxtBod_Change()
    Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtBod.Text & "'")
        If RBuscaBodega.RecordCount > 0 Then
            LblBod.Caption = RBuscaBodega!Descripcion
        Else
            LblBod.Caption = ""
        End If
    
End Sub

Private Sub TxtEncabezado_Change(Index As Integer)
    'BUSCA TRANSPORTISTA
    If Index = 0 Then
        Set RBuscaTransportista = Db.OpenRecordset("Select Descripcion From Transportistas Where CodigoTransportista = '" & TxtEncabezado.Item(0).Text & "'")
            If RBuscaTransportista.RecordCount > 0 Then
                LblTransportista.Caption = RBuscaTransportista!Descripcion
            Else
                LblTransportista.Caption = ""
            End If
    'BUSCA TIPO DE DOCUMENTO
    ElseIf Index = 1 Then
        Set RBuscaTipoDocumento = Db.OpenRecordset("Select Descripcion From Documentos Where CodigoDocumento = '" & TxtEncabezado.Item(1).Text & "'")
            If RBuscaTipoDocumento.RecordCount > 0 Then
                LblDoc.Caption = RBuscaTipoDocumento!Descripcion
            Else
                LblDoc.Caption = ""
            End If
    End If
End Sub

Private Sub TxtNumIng_GotFocus()
        TxtNumIng.SelStart = 0
        TxtNumIng.SelLength = Len(TxtNumIng.Text)
End Sub

Private Sub TxtNumIng_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BMateriaPrima = False
            BCliente = False
            BTransportista = False
            BPedido = False
            BNumeroIngreso = True
            BDocumento = False
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
            'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
            DataBuscar.RecordSource = "Select BodegaDisponibilidad, NumeroIngreso, CantidadTraslado, CantidadSalida, SaldoDisponibilidad From DetalleEntradasMateriaPrima Where Codigo = '" & TxtCodMatPri.Text & "' And SaldoDisponibilidad > 0 Order By BodegaDisponibilidad, NumeroIngreso"
            DataBuscar.Refresh
            DBGridBuscar.Refresh
        End If

End Sub


Private Sub Txtbuscar_Change()
        'MATERIA PRIMA
        If BMateriaPrima = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                If OptPalIni.Value = True Then
                        DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '" & TxtBuscar.Text & "*'")
                Else
                        DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & TxtBuscar.Text & "*'")
                End If
            ElseIf OptDescripcion.Value = True Then
                If OptPalIni.Value = True Then
                        DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*'")
                        
                Else
                        DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                End If
            End If
        'CLIENTE
        ElseIf BCliente = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                If OptPalIni.Value = True Then
                        DataBuscar.RecordSource = ("Select CodigoCliente, Descripcion, Contacto from Clientes Where CodigoCliente Like '" & TxtBuscar.Text & "*'")
                Else
                        DataBuscar.RecordSource = ("Select CodigoCliente, Descripcion, Contacto from Clientes Where CodigoCliente Like '*" & TxtBuscar.Text & "*'")
                End If
            ElseIf OptDescripcion.Value = True Then
                If OptPalIni.Value = True Then
                        DataBuscar.RecordSource = ("Select CodigoCliente, Descripcion, Contacto from Clientes Where Descripcion Like '" & TxtBuscar.Text & "*'")
                Else
                        DataBuscar.RecordSource = ("Select CodigoCliente, Descripcion, Contacto from Clientes Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                End If
            End If
        'TRANSPORTISTA
        ElseIf BCliente = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                If OptPalIni.Value = True Then
                        DataBuscar.RecordSource = ("Select * from Transportistas Where CodigoTransportista Like '" & TxtBuscar.Text & "*'")
                Else
                        DataBuscar.RecordSource = ("Select * from Transportistas Where CodigoTransportista Like '*" & TxtBuscar.Text & "*'")
                End If
            ElseIf OptDescripcion.Value = True Then
                If OptPalIni.Value = True Then
                        DataBuscar.RecordSource = ("Select * from Transportistas Where Descripcion Like '" & TxtBuscar.Text & "*'")
                Else
                        DataBuscar.RecordSource = ("Select * from Transportistas Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                End If
            End If
        'DOCUMENTOS
        ElseIf BDocumento = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                If OptPalIni.Value = True Then
                        DataBuscar.RecordSource = ("Select * from Documentos Where CodigoDocumento Like '" & TxtBuscar.Text & "*'")
                Else
                        DataBuscar.RecordSource = ("Select * from Documentos Where CodigoDocumento Like '*" & TxtBuscar.Text & "*'")
                End If
            ElseIf OptDescripcion.Value = True Then
                If OptPalIni.Value = True Then
                        DataBuscar.RecordSource = ("Select * from Documentos Where Descripcion Like '" & TxtBuscar.Text & "*'")
                Else
                        DataBuscar.RecordSource = ("Select * from Documentos Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                End If
            End If
            
            
        End If
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"

End Sub


Private Sub TxtBuscar_GotFocus()
    TxtBuscar.SelStart = 0
    TxtBuscar.SelLength = Len(TxtBuscar.Text)
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtCli_Change()
            Set RBuscaCliente = Db.OpenRecordset("Select Descripcion From Clientes Where CodigoCliente = '" & TxtCli.Text & "'")
                If RBuscaCliente.RecordCount > 0 Then
                    LblCliente.Caption = RBuscaCliente!Descripcion
                Else
                    LblCliente.Caption = ""
                End If
End Sub

Private Sub TxtCli_DblClick()
            BMateriaPrima = False
            BNumeroIngreso = False
            BCliente = True
            BTransportista = False
            BPedido = False
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
            DataBuscar.RecordSource = ("Select CodigoCliente, Descripcion, Contacto from Clientes Order by CodigoCliente")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"

End Sub

Private Sub TxtCli_GotFocus()
            TxtCli.SelStart = 0
            TxtCli.SelLength = Len(TxtCli.Text)
End Sub

Private Sub TxtCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
            BMateriaPrima = False
            BNumeroIngreso = False
            BCliente = True
            BTransportista = False
            BPedido = False
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
            DataBuscar.RecordSource = ("Select CodigoCliente, Descripcion, Contacto from Clientes Order by CodigoCliente")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"
    End If
End Sub

Private Sub TxtCodMatPri_Change()
                Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCodMatPri.Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaMateriaPrima!Descripcion
                Else
                        TxtDesPro.Text = ""
                End If
End Sub
Private Sub TxtCodMatPri_DblClick()
        BMateriaPrima = True
        BCliente = False
        BTransportista = False
        BPedido = False
        BNumeroIngreso = False
        BDocumento = False
        Framebuscar.Visible = True
        TxtBuscar.SetFocus
        'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
        DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion from correlativosmateriaprima")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
End Sub
Private Sub TxtCodMatPri_GotFocus()
        TxtCodMatPri.SelStart = 0
        TxtCodMatPri.SelLength = Len(TxtCodMatPri.Text)
End Sub

Private Sub TxtCodMatPri_KeyPress(KeyAscii As Integer)
        'SI PRECIONA ENTER
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
        'SI PRECIONA LA TECLA DE SIGNO +
        If KeyAscii = 43 Then
            BMateriaPrima = True
            BNumeroIngreso = False
            BCliente = False
            BPedido = False
            BTransportista = False
            BDocumento = False
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
            DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion from correlativosmateriaprima")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"
        End If
End Sub


Private Sub TxtDesPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtDoc_GotFocus()
    TxtDoc.SelStart = 0
    TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Private Sub TxtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtEncabezado_DblClick(Index As Integer)
    'TRANSPORTISTA
    If Index = 0 Then
            BMateriaPrima = False
            BNumeroIngreso = False
            BCliente = False
            BTransportista = True
            BPedido = False
            BDocumento = False
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
            DataBuscar.RecordSource = ("Select * from Transportistas")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"
    'DOCUMENTOS
    ElseIf Index = 1 Then
            BMateriaPrima = False
            BNumeroIngreso = False
            BCliente = False
            BTransportista = False
            BPedido = False
            BDocumento = True
            Framebuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
            DataBuscar.RecordSource = ("Select * from Documentos")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"

    End If
End Sub

Private Sub TxtEncabezado_GotFocus(Index As Integer)
        TxtEncabezado.Item(Index).SelStart = 0
        TxtEncabezado.Item(Index).SelLength = Len(TxtEncabezado.Item(Index).Text)
End Sub

Private Sub TxtEncabezado_KeyPress(Index As Integer, KeyAscii As Integer)

    If KeyAscii = 13 Then
            SendKeys "{tab}"
    End If
    
    'SI PRECIONA EL SIGNO +
    If KeyAscii = 43 Then
            'TRANSPORTISTA
            If Index = 0 Then
                    BMateriaPrima = False
                    BNumeroIngreso = False
                    BCliente = False
                    BTransportista = True
                    BPedido = False
                    BDocumento = False
                    Framebuscar.Visible = True
                    TxtBuscar.SetFocus
                   'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
                    DataBuscar.RecordSource = ("Select * from Transportistas Order By Descripcion")
                    DataBuscar.Refresh
                    DBGridBuscar.Refresh
                    DBGridBuscar.Columns(1).Width = "4000"
            'DOCUMENTOS
            ElseIf Index = 1 Then
                    BMateriaPrima = False
                    BNumeroIngreso = False
                    BCliente = False
                    BTransportista = False
                    BPedido = False
                    BDocumento = True
                    Framebuscar.Visible = True
                    TxtBuscar.SetFocus
                   'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
                    DataBuscar.RecordSource = ("Select * from Documentos")
                    DataBuscar.Refresh
                    DBGridBuscar.Refresh
                    DBGridBuscar.Columns(1).Width = "4000"
            End If
    End If
End Sub

Private Sub TxtNumIng_DblClick()
        BMateriaPrima = False
        BCliente = False
        BTransportista = False
        BPedido = False
        BNumeroIngreso = True
        BDocumento = False
        Framebuscar.Visible = True
        TxtBuscar.SetFocus
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataBuscar.RecordSource = "Select BodegaDisponibilidad, NumeroIngreso, CantidadTraslado, CantidadSalida, SaldoDisponibilidad From DetalleEntradasMateriaPrima Where Codigo = '" & TxtCodMatPri.Text & "' And SaldoDisponibilidad > 0 Order By BodegaDisponibilidad, NumeroIngreso"
        DataBuscar.Refresh
        DBGridBuscar.Refresh
End Sub



Private Sub TxtNumIng_LostFocus()

    If IsNumeric(TxtNumIng.Text) Then
            'BUSCA EL NUMERO DE INGRESO Y ASIGNA LA BODEGA, CODIGO Y CANTIDAD DE ACUERDO COMO ENTRO A LA BODEGA
            Set RBuscaNumeroIngreso = Db.OpenRecordset("Select SaldoDisponibilidad, Codigo, BodegaDisponibilidad From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCodMatPri.Text & "'")
            'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
            If RBuscaNumeroIngreso.RecordCount > 0 Then
                MskCanMatPri.Text = RBuscaNumeroIngreso!SaldoDisponibilidad
                TxtBod.Text = RBuscaNumeroIngreso!BodegaDisponibilidad
            'SI NO ENCUENTRA EL INGRESO DEJA EN BLANCO
            Else
                MskCanMatPri.Text = 0
                TxtBod.Text = ""
            End If
    End If

End Sub


Public Sub BotonesVisiblesEncabezado()
    If Bandera4 = False Then
         CmdAgregar.Visible = False
         CmdEditar.Visible = False
         CmdGrabar.Visible = False
         CmdBorrar.Visible = False
         CmdCancelar.Visible = False
         CmdBuscar.Visible = False
         CmdBuscar2.Visible = False
         CmdImprimir.Visible = False
         CmdSalida.Visible = False
         DataEgresos.Visible = False
    Else
         CmdAgregar.Visible = True
         CmdEditar.Visible = True
         CmdGrabar.Visible = True
         CmdBorrar.Visible = True
         CmdCancelar.Visible = True
         CmdBuscar.Visible = True
         CmdBuscar2.Visible = True
         CmdImprimir.Visible = True
         CmdSalida.Visible = True
         DataEgresos.Visible = True
    End If

End Sub
