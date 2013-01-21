VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form IngresosMateriaPrima 
   Caption         =   "Captura de INGRESOS De Materia Prima"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "IngresosMateriaPrima.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
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
      Height          =   7335
      Left            =   0
      TabIndex        =   31
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
         TabIndex        =   42
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
            TabIndex        =   26
            Top             =   360
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
            TabIndex        =   25
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   10920
         Picture         =   "IngresosMateriaPrima.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   32
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
         RecordSource    =   "InventarioMateriaPrima"
         Top             =   3360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1680
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGridBuscar 
         Bindings        =   "IngresosMateriaPrima.frx":0884
         Height          =   6255
         Left            =   120
         OleObjectBlob   =   "IngresosMateriaPrima.frx":089D
         TabIndex        =   22
         ToolTipText     =   "Doble Click o Esc Para Seleccionar"
         Top             =   960
         Width           =   11535
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Top             =   600
         Width           =   6255
      End
   End
   Begin MSDBGrid.DBGrid DBGridDetalleIngresos 
      Bindings        =   "IngresosMateriaPrima.frx":1275
      Height          =   3495
      Left            =   120
      OleObjectBlob   =   "IngresosMateriaPrima.frx":1297
      TabIndex        =   47
      Top             =   3840
      Width           =   11535
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
      Caption         =   "Encabezado de Ingreso"
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
      Height          =   2415
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   480
         Left            =   8760
         Picture         =   "IngresosMateriaPrima.frx":1FE0
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&EDITAR"
         Height          =   480
         Left            =   1560
         Picture         =   "IngresosMateriaPrima.frx":2512
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&USCAR"
         Height          =   480
         Left            =   7320
         Picture         =   "IngresosMateriaPrima.frx":2A44
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   10200
         Picture         =   "IngresosMateriaPrima.frx":2F76
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salida"
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&BORRAR"
         Height          =   480
         Left            =   5880
         Picture         =   "IngresosMateriaPrima.frx":33B8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&CANCELAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   4440
         Picture         =   "IngresosMateriaPrima.frx":38EA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&GRABAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   3000
         Picture         =   "IngresosMateriaPrima.frx":3E1C
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&AGREGAR"
         Height          =   480
         Left            =   120
         Picture         =   "IngresosMateriaPrima.frx":434E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1800
         Width           =   1400
      End
      Begin VB.Frame FrameIngresos 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   11535
         Begin VB.TextBox TxtEst 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            DataField       =   "Estado"
            DataSource      =   "DataIngresos"
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
            TabIndex        =   48
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox TxtLib 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Liberado"
            DataSource      =   "DataIngresos"
            Height          =   285
            Left            =   10080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataIngresos"
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1080
            Width           =   6975
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataIngresos"
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
         Begin VB.TextBox TxtBodega 
            Appearance      =   0  'Flat
            DataField       =   "BodegaEntrada"
            DataSource      =   "DataIngresos"
            Height          =   285
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtReq 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Requerido"
            DataSource      =   "DataIngresos"
            Height          =   285
            Left            =   10080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtDoc 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataIngresos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   6480
            TabIndex        =   1
            Top             =   240
            Width           =   2055
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
            TabIndex        =   49
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
            TabIndex        =   46
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
            TabIndex        =   45
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
            TabIndex        =   44
            Top             =   1080
            Width           =   1455
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
            Left            =   3000
            TabIndex        =   43
            Top             =   720
            Width           =   5535
         End
         Begin VB.Label Label6 
            Caption         =   "Bodega Entrada"
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
            TabIndex        =   41
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4800
            TabIndex        =   40
            Top             =   240
            Width           =   1620
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
            TabIndex        =   39
            Top             =   360
            Width           =   1455
         End
      End
   End
   Begin VB.Frame FrameDetalle 
      Caption         =   "Detalle de Ingreso"
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
      Height          =   5535
      Left            =   0
      TabIndex        =   27
      Top             =   2520
      Width           =   11805
      Begin VB.CommandButton CmdEditar2 
         Caption         =   "Editar"
         Height          =   495
         Left            =   2040
         Picture         =   "IngresosMateriaPrima.frx":4880
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   4920
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton CmdBorrar2 
         Caption         =   "B&orrar"
         Height          =   495
         Left            =   7800
         Picture         =   "IngresosMateriaPrima.frx":4DB2
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   4920
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton CmdCancelar2 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   5880
         Picture         =   "IngresosMateriaPrima.frx":52E4
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4920
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
         Picture         =   "IngresosMateriaPrima.frx":5816
         TabIndex        =   20
         Top             =   4920
         Visible         =   0   'False
         Width           =   1920
      End
      Begin VB.CommandButton CmdGrabar2 
         Caption         =   "G&rabar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3960
         Picture         =   "IngresosMateriaPrima.frx":5D48
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4920
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.CommandButton CmdAgregar2 
         Caption         =   "A&gregar"
         Height          =   495
         Left            =   120
         Picture         =   "IngresosMateriaPrima.frx":627A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4920
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.Frame FrameDetalleIngresos 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   11535
         Begin MSMask.MaskEdBox TxtCanMatPri 
            DataField       =   "Cantidad"
            DataSource      =   "DataDetalleIngresos"
            Height          =   255
            Left            =   10080
            TabIndex        =   15
            Top             =   480
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBod 
            Appearance      =   0  'Flat
            DataField       =   "Bodega"
            DataSource      =   "DataDetalleIngresos"
            Height          =   285
            Left            =   6000
            MaxLength       =   3
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox TxtDocDet 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataDetalleIngresos"
            Height          =   285
            Left            =   7200
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtCodMatPri 
            Appearance      =   0  'Flat
            DataField       =   "Codigo"
            DataSource      =   "DataDetalleIngresos"
            Height          =   285
            Left            =   120
            MaxLength       =   10
            TabIndex        =   14
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
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   480
            Width           =   8055
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
            Left            =   10080
            TabIndex        =   38
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
            TabIndex        =   37
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
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin VB.Data DataDetalleIngresos 
      Caption         =   "Detalle Ingresos"
      Connect         =   "Access"
      DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleIngresosMateriaPrima"
      Top             =   8160
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Data DataIngresos 
      Caption         =   "Ingresos"
      Connect         =   "Access"
      DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EncabezadoIngresosMateriaPrima"
      Top             =   8145
      Width           =   11190
   End
End
Attribute VB_Name = "IngresosMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Mensaje As String
Dim VDocumento As Double
Dim VDocumentoDetalle As Double
Dim VCantidadMateriaPrima As Double
Dim VCodigoMateriaPrima As String
Dim VBodega As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim BBodega As Boolean
Dim BMateriaPrima As Boolean

Dim RBuscaMateriaPrima As Recordset
Dim RBuscaSigDoc As Recordset
Dim RBuscaDetalle As Recordset
Dim RBuscaBodega As Recordset

Sub Botones1()
    If Bandera = True Then
         FrameIngresos.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         DataIngresos.Visible = False
    Else
         FrameIngresos.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         DataIngresos.Visible = True
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleIngresos.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdEditar2.Enabled = False
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalleIngresos.Enabled = False
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
    DataDetalleIngresos.Recordset.AddNew
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description
    End If
        
    Bandera2 = True
    Botones2

    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DBGridDetalleIngresos.Enabled = False
    'SE LE ASIGNA A LA BODEGA DETALLE LA BODEGA DEL ENCABEZADO
    TxtBod.Text = VBodega
    'SE ASIGNA AL DOCUMENTO DE DETALLE EL DOCUMENTO DEL ENCABEZADO
    TxtDocDet.Text = VDocumento
    TxtCodMatPri.SetFocus
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next

        'REVISA SI YA FUE LIBERADO EL INGRESO
            If TxtEst.Text = "LIBERADO" Then
                MsgBox "No Puede BORRAR Este Ingreso Porque Ya Fue Liberado", vbOKOnly + vbCritical, "Informacion"
                Exit Sub
            End If

VDocumento = TxtDoc.Text

            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If Mensaje = vbOK Then
                MousePointer = 11
                
                'SE BUSCAN TODO EL DETALLE DE EL INGRESO PARA PODER SUMAR DE NUEVO LA EXISTENCIA
                'Set RBuscaDetalle = Db.OpenRecordset("Select Codigo, Cantidad, Bodega from DetalleIngresosMateriaPrima where Documento = " & VDocumento)
                
                'If RBuscaDetalle.RecordCount > 0 Then
                '    Do Until RBuscaDetalle.EOF
                    
                        'BUSCAMOS LA MATERIA PRIMA PARA REGRESAR LAS CANTIDADES Y SUMAR EXISTENCIAS
                '        Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Entradas From InventarioMateriaPrima where CodigoMateriaPrima = '" & RBuscaDetalle!Codigo & "' and Bodega = '" & RBuscaDetalle!Bodega & "'")
                '            RBuscaMateriaPrima.Edit
                '                RBuscaMateriaPrima!Existencia = RBuscaMateriaPrima!Existencia - RBuscaDetalle!Cantidad
                '                RBuscaMateriaPrima!Entradas = RBuscaMateriaPrima!Entradas - RBuscaDetalle!Cantidad
                '            RBuscaMateriaPrima.Update
                            
                '        RBuscaDetalle.MoveNext
                '    Loop
                'End If
                
                'BORRA EL ENCABEZADO DE LA FACTURA
                DataIngresos.Recordset.Delete
                DataIngresos.Recordset.MoveLast
                MousePointer = 0
            End If
  
            If DataIngresos.Recordset.EOF Then
                DataIngresos.Recordset.MoveLast
                If Err = 3021 Then
                    Mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next

    'ASIGANMOS A UNA VARIABLE EL DOCUMENTO DETALLE
    VDocumentoDetalle = TxtDocDet.Text
    VCodigoMateriaPrima = TxtCodMatPri.Text
    VCantidadMateriaPrima = TxtCanMatPri.Text

            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If Mensaje = vbOK Then
                MousePointer = 11
                
                    'BUSCAMOS LA MATERIA PRIMA SUMAMOS LA EXISTENCIA Y RESTAMOS LAS SALIDAS
                    'Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Entradas From InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and bodega = '" & VBodega & "'")
                    '    RBuscaMateriaPrima.Edit
                    '        RBuscaMateriaPrima!Existencia = RBuscaMateriaPrima!Existencia - VCantidadMateriaPrima
                    '        RBuscaMateriaPrima!Entradas = RBuscaMateriaPrima!Entradas - VCantidadMateriaPrima
                    '    RBuscaMateriaPrima.Update
                        
                    '    If Err <> 0 Then
                    '    End If
                        
                   'BORRA EL DETALLE DE LA FACTURA
                    DataDetalleIngresos.Recordset.Delete
                    
                        If Err <> 0 Then
                            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                        End If
                                        
                    'SELECCIONA TODOS LOS DETALLES DE LA FACTURA
                    DataDetalleIngresos.RecordSource = ("Select * from DetalleIngresosMateriaPrima where documento = " & VDocumentoDetalle & " order By Codigo")
                    DataDetalleIngresos.Refresh
                    DBGridDetalleIngresos.Refresh
                MousePointer = 0
            End If
  
            If DataIngresos.Recordset.EOF Then
                DataIngresos.Recordset.MoveLast
                If Err = 3021 Then
                    Mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
End Sub

Private Sub CmdBuscar_Click()
    Mensaje = InputBox("No. Documento A Buscar")
    DataIngresos.Recordset.FindFirst ("Documento = " & Mensaje)
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next

    'GRABA DATOS
    DataIngresos.Recordset.CancelUpdate
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    'CANCELA EL INGRESO
    DataDetalleIngresos.Recordset.CancelUpdate
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    DBGridDetalleIngresos.Enabled = True
    Bandera2 = False
    Botones2

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next

    'REVISA SI YA FUE LIBERADO EL INGRESO
    If TxtEst.Text = "LIBERADO" Then
        MsgBox "No Puede EDITAR Este Ingreso Porque Ya Fue Liberado", vbOKOnly + vbCritical, "Informacion"
        Exit Sub
    End If

    'EDITA EL REGISTRO ACTUAL
    DataIngresos.Recordset.Edit
    
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    MskFec.SetFocus
End Sub


Private Sub CmdEditar2_Click()
On Error Resume Next
    'EDITA EL REGISTRO ACTUAL
    DataDetalleIngresos.Recordset.Edit
    
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DBGridDetalleIngresos.Enabled = False
    
    Bandera2 = True
    Botones2
End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
    
    'GUARDA VARIABLES
    VCantidadMateriaPrima = TxtCanMatPri.Text
    VCodigoMateriaPrima = TxtCodMatPri.Text
            
    'REVISAMOS DATOS
    If Not IsNumeric(TxtCanMatPri.Text) Then
       MsgBox "Cantidad De Materia Prima Incorrecta", vbOKOnly + vbCritical, "Error"
       TxtCanMatPri.SetFocus
       Exit Sub
    End If
    
    'REVISA SI EXISTE LA MATERIA PRIMA
    Set RBuscaMateriaPrima = Db.OpenRecordset("Select * from InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and Bodega = '" & TxtBod.Text & "'")
    If RBuscaMateriaPrima.RecordCount > 0 Then
    Else
        MsgBox "Codigo De Materia Prima No Existe En La Bodega Que Esta Trabajando", vbOKOnly + vbCritical, "Error"
        TxtCodMatPri.SetFocus
        Exit Sub
    End If
        
    'GRABA DATOS
    DataDetalleIngresos.Recordset.Update
        
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera2 = False
    Botones2
    
     'BUSCA LA EXISTENCIA Y LAS ENTRADAS DE LA MATERIA PRIMA Y SUMA LA EXISTENCIA Y SUMA ENTRADAS
    'Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Entradas From InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and Bodega = '" & VBodega & "'")
    '    If RBuscaMateriaPrima.RecordCount > 0 Then
    '            RBuscaMateriaPrima.Edit
    '                RBuscaMateriaPrima!Existencia = (RBuscaMateriaPrima!Existencia + VCantidadMateriaPrima)
    '                RBuscaMateriaPrima!Entradas = (RBuscaMateriaPrima!Entradas + VCantidadMateriaPrima)
    '            RBuscaMateriaPrima.Update
    '    End If
  
    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE LA FACTURA QUE SE ESTA GRABANDO
    DataDetalleIngresos.RecordSource = ("Select * from DetalleIngresosMateriaPrima where Documento = " & VDocumento & " Order BY Codigo")
    DataDetalleIngresos.Refresh
    DBGridDetalleIngresos.Refresh
        
    DBGridDetalleIngresos.Enabled = True
    CmdAgregar2_Click
End Sub


Private Sub CmdAgregar_Click()
    On Error Resume Next
    'AGREGA UN REGISTRO
    DataIngresos.Recordset.AddNew
    'SI HAY ERRORES SE SALE
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    
    'ASIGNA EL USUARIO
    TxtReq.Text = GUsuario
    'ASIGNA EL ESTADO DEL INGRESO
    TxtEst.Text = "NO LIBERADO"
    'ASIGNA LA FECHA DEL DIA
    MskFec.Text = Date
    MskFec.SetFocus
    
    'BUSCA EL DOCUMENTO MAXIMO Y LE ASIGNA 1
    Set RBuscaSigDoc = Db.OpenRecordset("Select Max(Documento) from EncabezadoIngresosMateriaPrima")
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
    
    VDocumento = TxtDoc.Text
    VBodega = TxtBodega.Text
    
    'VERIFICA SI ES NUMERICO EL DOCUMENTO
    If Not IsNumeric(TxtDoc.Text) Then
        MsgBox "Numero De Documento Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        TxtDoc.SetFocus
        Exit Sub
    End If
        
    'GRABA DATOS
    DataIngresos.Recordset.Update
    
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
    DataDetalleIngresos.RecordSource = ("Select * from DetalleIngresosMateriaPrima where Documento = " & VDocumento & " Order by Codigo")
    DataDetalleIngresos.Refresh
    DBGridDetalleIngresos.Refresh
    
    'MUEVE EL RECORDSET A LA FACTURA ACTUAL PARA QUE SE ACTUALIZEN LOS CAMBIOS
    DataIngresos.Recordset.FindFirst ("Documento = " & VDocumento)
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = True
    'FrameDetalle.Visible = True
    FrameEncabezado.Enabled = False
    CmdAgregar2_Click
End Sub

Private Sub CmdImprimir_Click()
MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDoc.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
        
                CrReportes.SelectionFormula = "{EncabezadoIngresosMateriaPrima.Documento} = " & TxtDoc.Text
                CrReportes.ReportFileName = App.Path & "\FormatoIngresosMateriaPrima.rpt"
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
    FrameDetalle.Enabled = False
    'FrameDetalle.Visible = False
    FrameEncabezado.Enabled = True
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle

End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub


Private Sub DataIngresos_Reposition()
    If IsNumeric(TxtDoc.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL INGRESO
        DataDetalleIngresos.RecordSource = ("Select * from DetalleIngresosMateriaPrima where Documento = " & TxtDoc.Text & " Order by Codigo")
        DataDetalleIngresos.Refresh
        DBGridDetalleIngresos.Refresh
    End If

End Sub

Private Sub DBGridBuscar_DblClick()
        If BMateriaPrima = True Then
            TxtCodMatPri.Text = DBGridBuscar.Columns(0)
        ElseIf BBodega = True Then
            TxtBodega.Text = DBGridBuscar.Columns(0)
            TxtBodega.SetFocus
        End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BMateriaPrima = True Then
                TxtCodMatPri.Text = DBGridBuscar.Columns(0)
            ElseIf BBodega = True Then
                TxtBodega.Text = DBGridBuscar.Columns(0)
                TxtBodega.SetFocus
            End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
        End If
End Sub



Private Sub Form_Activate()
    If IsNumeric(TxtDoc.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL INGRESO
        DataDetalleIngresos.RecordSource = ("Select * from DetalleIngresosMateriaPrima where Documento = " & TxtDoc.Text & " Order by Codigo")
        DataDetalleIngresos.Refresh
        DBGridDetalleIngresos.Refresh
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
    

End Sub

Private Sub Form_Load()
    'ASIGNAMOS A QUE BASE DE DATOS SE VAN A CONECTAR LOS DATAS
    DataIngresos.DatabaseName = BasedeDatos
    DataDetalleIngresos.DatabaseName = BasedeDatos
    DataBuscar.DatabaseName = BasedeDatos
    
    
    
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


Private Sub TxtBodega_Change()
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtBodega.Text & "'")
        If RBuscaBodega.RecordCount > 0 Then
            LblBodega.Caption = RBuscaBodega!Descripcion
        Else
            LblBodega.Caption = ""
        End If
    
End Sub

Private Sub TxtBodega_DblClick()
        BMateriaPrima = False
        BBodega = True
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        'SELECCIONA LAS BODEGAS DE MATERIA PRIMA
        DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima")
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
        BMateriaPrima = False
        BBodega = True
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        'SELECCIONA LAS BODEGAS DE MATERIA PRIMA
        DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
    End If

End Sub

Private Sub TxtBuscar_Change()
If BMateriaPrima = True Then
    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
    If OptCodigo.Value = True Then
        If OptPalIni.Value = True Then
                DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' And CodigoMateriaPrima Like '" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
        Else
                DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' And CodigoMateriaPrima Like '*" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
        End If
    ElseIf OptDescripcion.Value = True Then
        If OptPalIni.Value = True Then
                DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' And Descripcion Like '" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
                
        Else
                DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' And Descripcion Like '*" & TxtBuscar.Text & "*' Order by CodigoMateriaPrima")
        End If
    End If
ElseIf BBodega = True Then
    'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
    If OptCodigo.Value = True Then
        If OptPalIni.Value = True Then
                DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where CodigoBodega Like '" & TxtBuscar.Text & "*'")
        Else
                DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where CodigoBodega Like '*" & TxtBuscar.Text & "*'")
        End If
    ElseIf OptDescripcion.Value = True Then
        If OptPalIni.Value = True Then
                DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*'")
                
        Else
                DataBuscar.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*'")
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

Private Sub txtCanMatPri_GotFocus()
        TxtCanMatPri.SelStart = 0
        TxtCanMatPri.SelLength = Len(TxtCanMatPri.Text)
End Sub
Private Sub txtCanMatPri_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub
Private Sub TxtCodMatPri_Change()
                Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From InventarioMateriaPrima Where CodigoMateriaPrima = '" & TxtCodMatPri.Text & "' And Bodega = '" & TxtBod.Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaMateriaPrima!Descripcion
                Else
                        TxtDesPro.Text = ""
                End If
End Sub
Private Sub TxtCodMatPri_DblClick()
        BMateriaPrima = True
        BBodega = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
        DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' Order by CodigoMateriaPrima")
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
            BBodega = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL INVENTARIO DE ACUERDO A LA BODEGA QUE VA A UTILIZAR
            DataBuscar.RecordSource = ("Select CodigoMateriaPrima, Descripcion, Existencia, TipoDeMateriaPrima from InventarioMateriaPrima Where Bodega = '" & VBodega & "' Order by CodigoMateriaPrima")
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

Private Sub TxtObs_GotFocus()
    TxtObs.SelStart = 0
    TxtObs.SelLength = Len(TxtObs.Text)
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub
