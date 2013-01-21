VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CapturaProduccionEncajada 
   BackColor       =   &H000000FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura de Produccion Encajada"
   ClientHeight    =   7125
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "CapturaProduccionEncajada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   11775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   7095
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Data DataConsultas 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1200
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   10680
         Picture         =   "CapturaProduccionEncajada.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "CapturaProduccionEncajada.frx":0614
         Height          =   6735
         Left            =   120
         OleObjectBlob   =   "CapturaProduccionEncajada.frx":0630
         TabIndex        =   35
         Top             =   240
         Width           =   10455
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   6120
      Top             =   8040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      CopiesToPrinter =   3
      WindowState     =   2
   End
   Begin VB.Data DataProduccion 
      Caption         =   "Produccion Encajada"
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
      RecordSource    =   "ProduccionEncajada"
      Top             =   6720
      Width           =   11520
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Im&primir Boleta"
      Height          =   585
      Index           =   5
      Left            =   8520
      MouseIcon       =   "CapturaProduccionEncajada.frx":100B
      Picture         =   "CapturaProduccionEncajada.frx":144D
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6000
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   585
      Index           =   6
      Left            =   10200
      MouseIcon       =   "CapturaProduccionEncajada.frx":1597
      Picture         =   "CapturaProduccionEncajada.frx":19D9
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6000
      Width           =   1485
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   585
      Index           =   4
      Left            =   6840
      MouseIcon       =   "CapturaProduccionEncajada.frx":3A4B
      Picture         =   "CapturaProduccionEncajada.frx":3E8D
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6000
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   585
      Index           =   3
      Left            =   5160
      MouseIcon       =   "CapturaProduccionEncajada.frx":43BF
      Picture         =   "CapturaProduccionEncajada.frx":4801
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   585
      Index           =   2
      Left            =   3480
      MouseIcon       =   "CapturaProduccionEncajada.frx":4D33
      Picture         =   "CapturaProduccionEncajada.frx":5175
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6000
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   585
      Index           =   1
      Left            =   1800
      MouseIcon       =   "CapturaProduccionEncajada.frx":56A7
      Picture         =   "CapturaProduccionEncajada.frx":5AE9
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6000
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   585
      Index           =   0
      Left            =   240
      MouseIcon       =   "CapturaProduccionEncajada.frx":601B
      Picture         =   "CapturaProduccionEncajada.frx":645D
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6000
      Width           =   1485
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10398
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CapturaProduccionEncajada.frx":698F
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameProduccion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General "
      TabPicture(1)   =   "CapturaProduccionEncajada.frx":6CA9
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridProduccion"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda O Seleccion De Datos"
      TabPicture(2)   =   "CapturaProduccionEncajada.frx":70FB
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LblEtiqueta(0)"
      Tab(2).Control(1)=   "LblEtiqueta(2)"
      Tab(2).Control(2)=   "LblEtiqueta(3)"
      Tab(2).Control(3)=   "FrameBuscar"
      Tab(2).Control(4)=   "CmdActualizar"
      Tab(2).Control(5)=   "CmdBuscar"
      Tab(2).Control(6)=   "TxtBuscar"
      Tab(2).Control(7)=   "DtpFecIni"
      Tab(2).Control(8)=   "DtpFecFin"
      Tab(2).ControlCount=   9
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   375
         Left            =   -65280
         TabIndex        =   30
         Top             =   2760
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24772611
         CurrentDate     =   37213
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   375
         Left            =   -68040
         TabIndex        =   29
         Top             =   2760
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24772611
         CurrentDate     =   37213
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -65280
         TabIndex        =   31
         ToolTipText     =   " "
         Top             =   3480
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Datos"
         Height          =   855
         Left            =   -66480
         Picture         =   "CapturaProduccionEncajada.frx":754D
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3960
         Width           =   3015
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Seleccionar Todos Datos"
         Height          =   825
         Left            =   -66480
         Picture         =   "CapturaProduccionEncajada.frx":798F
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   4920
         Width           =   3045
      End
      Begin VB.Frame FrameBuscar 
         BackColor       =   &H8000000B&
         Caption         =   "Opciones de Busqueda"
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   23
         Top             =   840
         Width           =   7455
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fecha Y Ficha Tecnica"
            ForeColor       =   &H8000000D&
            Height          =   1000
            Index           =   2
            Left            =   3000
            Picture         =   "CapturaProduccionEncajada.frx":7C99
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   360
            Width           =   1400
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "MP 9301"
            Height          =   1000
            Index           =   5
            Left            =   5880
            Picture         =   "CapturaProduccionEncajada.frx":8563
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   360
            Width           =   1400
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas"
            Height          =   1000
            Index           =   0
            Left            =   120
            Picture         =   "CapturaProduccionEncajada.frx":886D
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Value           =   -1  'True
            Width           =   1400
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fecha Y Linea"
            ForeColor       =   &H8000000D&
            Height          =   1000
            Index           =   1
            Left            =   1560
            Picture         =   "CapturaProduccionEncajada.frx":8B77
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   360
            Width           =   1400
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Batch"
            ForeColor       =   &H80000008&
            Height          =   1000
            Index           =   4
            Left            =   4440
            Picture         =   "CapturaProduccionEncajada.frx":8E81
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   360
            Width           =   1400
         End
      End
      Begin VB.Frame FrameProduccion 
         Caption         =   "Datos Generales Captura "
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
         Height          =   5055
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   11535
         Begin VB.TextBox TxtTur 
            Appearance      =   0  'Flat
            DataField       =   "Turno"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   10
            Top             =   3360
            Width           =   735
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataProduccion"
            Height          =   555
            Index           =   8
            Left            =   1320
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   13
            Top             =   4440
            Width           =   10035
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "NoMP9301"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   7
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   11
            ToolTipText     =   "No. Hoja De Identificacion"
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Usuario"
            DataSource      =   "DataProduccion"
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
            Index           =   9
            Left            =   9960
            MaxLength       =   10
            TabIndex        =   14
            Top             =   360
            Width           =   1455
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Envases"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   6
            Left            =   1320
            TabIndex        =   8
            Top             =   2520
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Batch"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   5
            Left            =   1320
            TabIndex        =   7
            ToolTipText     =   "agrupacion de 16 tarimas"
            Top             =   2160
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            DataField       =   "Tarima"
            DataSource      =   "DataProduccion"
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
            Left            =   1320
            TabIndex        =   6
            Top             =   1800
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Hor_Prd"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   3
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   5
            Top             =   1440
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Fec_Prd"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   2
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1080
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Esp_Tec"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   1
            Left            =   1320
            MaxLength       =   12
            TabIndex        =   3
            Top             =   720
            Width           =   1400
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Linea"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   0
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   2
            Top             =   360
            Width           =   1400
         End
         Begin VB.ComboBox CboColor 
            Appearance      =   0  'Flat
            DataField       =   "ColorMP9301"
            DataSource      =   "DataProduccion"
            Height          =   315
            ItemData        =   "CapturaProduccionEncajada.frx":92C3
            Left            =   1320
            List            =   "CapturaProduccionEncajada.frx":92D0
            TabIndex        =   12
            Text            =   "BLANCA"
            Top             =   4080
            Width           =   1335
         End
         Begin VB.ComboBox CboCal 
            BackColor       =   &H0080C0FF&
            DataField       =   "Calidad"
            DataSource      =   "DataProduccion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "CapturaProduccionEncajada.frx":92EB
            Left            =   1320
            List            =   "CapturaProduccionEncajada.frx":92FB
            TabIndex        =   9
            Text            =   "A"
            ToolTipText     =   "Calidad De Tarima"
            Top             =   2880
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Observaciones"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   55
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Label LblFichaTecnica 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   52
            Top             =   720
            Width           =   6015
         End
         Begin VB.Label LblLinea 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   51
            Top             =   360
            Width           =   6015
         End
         Begin VB.Label lblLabels 
            Caption         =   "Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lblLabels 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   48
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Hora"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   47
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "Ficha Tecnica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   46
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tarima"
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
            Index           =   22
            Left            =   120
            TabIndex        =   45
            Top             =   1800
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Batch"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   44
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Envases"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   43
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Calidad"
            ForeColor       =   &H80000007&
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   42
            Top             =   2880
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Turno"
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   41
            Top             =   3360
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "No. MP9 301"
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
            Index           =   7
            Left            =   120
            TabIndex        =   40
            Top             =   3720
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Color MP9 301"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   39
            Top             =   4080
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Caption         =   "Usuario"
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
            Index           =   10
            Left            =   9240
            TabIndex        =   38
            Top             =   360
            Width           =   855
         End
         Begin VB.Label LblBatch 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   615
            Left            =   2760
            TabIndex        =   37
            Top             =   2160
            Width           =   735
         End
      End
      Begin MSDBGrid.DBGrid DBGridProduccion 
         Bindings        =   "CapturaProduccionEncajada.frx":930B
         Height          =   5055
         Left            =   -74880
         OleObjectBlob   =   "CapturaProduccionEncajada.frx":9328
         TabIndex        =   22
         Tag             =   "Click En Encabezado De Columna Para Indexar"
         Top             =   720
         Width           =   11535
      End
      Begin VB.Label LblEtiqueta 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
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
         Height          =   195
         Index           =   3
         Left            =   -66000
         TabIndex        =   54
         Top             =   2880
         Width           =   510
      End
      Begin VB.Label LblEtiqueta 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Desde"
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
         Height          =   195
         Index           =   2
         Left            =   -69000
         TabIndex        =   53
         Top             =   2880
         Width           =   795
      End
      Begin VB.Label LblEtiqueta 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Index           =   0
         Left            =   -67920
         TabIndex        =   50
         Top             =   3480
         Width           =   2535
      End
   End
End
Attribute VB_Name = "CapturaProduccionEncajada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim Vlinea As String
Dim VFicha As String
Dim VFecha As Date

Dim RLineas As Recordset
Dim RBuscaProduccion As Recordset
Dim RBuscaEnvases As Recordset
Dim RReporteIdentificacionInterno As Recordset
Dim RBuscaUltimaFicha As Recordset

'VARIABLES PARA DESPLEGAR DATOS DE FICHA TECNICA
Dim RBuscaFichaTecnica As Recordset
Dim RBuscaAtributo As Recordset

Dim VLineas As Boolean

Dim VDia As String
Dim VMes As String
Dim VAño As String

Dim RCuentaTarimas As Recordset

Dim BEditar As Boolean

Dim RVerificaTarima As Recordset
Dim Vtarima As String

Dim RBuscaLinea As Recordset




Sub botones()
    If Bandera = True Then
         FrameProduccion.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         CmdBotones.Item(6).Enabled = False
         LblBatch.Visible = True
         TxtTexto.Item(1).SetFocus
         FrameBuscar.Visible = False
         DataProduccion.Visible = False
         DBGridProduccion.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         CmdBotones.Item(6).Enabled = True
         FrameProduccion.Enabled = False
         LblBatch.Visible = False
         FrameBuscar.Visible = True
         DataProduccion.Visible = True
         DBGridProduccion.Visible = True

    End If
End Sub

Private Sub CboCal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub


Private Sub CboColor_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub



Private Sub CmdActualizar_Click()
    DataProduccion.RecordSource = "Select * From ProduccionEncajada Order By Fec_Prd, Hor_Prd"
    DataProduccion.Refresh
    DBGridProduccion.Refresh
    SSTab1.Tab = 1
End Sub

Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    'AGREGAR
    If Index = 0 Then
                    DataProduccion.Recordset.AddNew
                    If Err <> 0 Then
                            MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                    End If
                    Bandera = True
                    botones
                    TxtTexto.Item(0).SetFocus
                    
                    'SI LA HORA ES MENOR QUE LAS 7 DE LA MAÑANA ENTONCES DA LA FECHA ANTERIOR
                    'If Format(Time, "hh:mm") < "07:00" Then
                    '    TxtTexto.Item(2).Text = Format(DateValue(Date) - 1, "dd/mm/yyyy")
                    'Else
                        TxtTexto.Item(2).Text = Format(Date, "dd/mm/yyyy")
                    'End If
                    
                    'If Format(Time, "hh:mm") >= "07:00" And Format(Time, "hh:mm") <= "19:00" Then
                    '    CboTur.Text = "D"
                    'Else
                    '    CboTur.Text = "N"
                    'End If
                    
                    TxtTexto.Item(3).Text = Format(Time, "hh:mm")
                    
                    CboCal.Text = "A"
                    
                    BEditar = False
    'EDITAR
    ElseIf Index = 1 Then
                    DataProduccion.Recordset.Edit
                    If Err <> 0 Then
                         MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                         Exit Sub
                    End If
                    Bandera = True
                    botones
                    TxtTexto.Item(1).SetFocus
                    BEditar = True
    'GRABAR
    ElseIf Index = 2 Then
                    
                   'REVISA EL TIPO DE CALIDAD
                   If CboCal.Text <> "A" And CboCal.Text <> "C" And CboCal.Text <> "I" And CboCal.Text <> "R" Then
                        MsgBox "CALIDAD INCORRECTA", vbOKOnly + vbInformation, "Informacion"
                        MousePointer = 0
                        Exit Sub
                   End If
                
                   'REVISA EL BATCH
                   If Not IsNumeric(TxtTexto.Item(5).Text) Then
                        MsgBox "Numero de Batch Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        MousePointer = 0
                        Exit Sub
                   End If
                   
                   Vlinea = TxtTexto.Item(0).Text
                   VFicha = TxtTexto.Item(1).Text
                   VFecha = TxtTexto.Item(2).Text
                   Vtarima = TxtTexto.Item(4).Text
                                           
                   'SI NO ESTA EDITANDO SOLO GRABANDO
                   If BEditar = False Then
                            'VERIFICA SI YA EXISTE LA TARIMA
                            Set RVerificaTarima = Db.OpenRecordset("Select * from ProduccionEncajada Where Linea = '" & Vlinea & "' and Esp_tec = '" & VFicha & "' and Fec_prd = #" & Format(VFecha, "dd/mm/yyyy") & "# and Tarima = " & Vtarima)
                            If RVerificaTarima.RecordCount > 0 Then
                                 mensaje = MsgBox("Ya Existe Tarima " & Vtarima & " De Ficha " & VFicha & " Con Fecha " & VFecha & " Ya Existe, No Se Puede Grabar ", vbOKOnly + vbInformation, "Verificacion")
                                 Exit Sub
                            End If
                   End If
                   
                   '__________________________________________________________________________________________________________________
                     
                   'GRABA DATOS
                   DataProduccion.Recordset.Update
                   
                   If Err <> 0 Then
                      MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                      TxtTexto.Item(0).SetFocus
                      Exit Sub
                   Else
                        If BEditar = False Then
                                    'BUSCAMOS LA LINEA Y LE ACTUALIZAMOS EL CONTADOR DE TARIMA
                                    Set RLineas = Db.OpenRecordset("Select Tarima from Lineas Where Linea = '" & Vlinea & "'")
                                    
                                    If RLineas.RecordCount > 0 Then
                                        RLineas.Edit
                                            RLineas!Tarima = Val(RLineas!Tarima) + 1
                                        RLineas.Update
                                    End If
                        End If
                
                      Bandera = False
                      botones
                      DataProduccion.Recordset.MoveLast
                      CmdBotones.Item(0).SetFocus
                  End If
    
    'CANCELAR
    ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    DataProduccion.Recordset.CancelUpdate
        
                    If Err <> 0 Then
                            MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                    End If
    'BORRAR
    ElseIf Index = 4 Then
                  mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
                  If mensaje = vbOK Then
                      DataProduccion.Recordset.Delete
                      DataProduccion.Recordset.MoveNext
                  End If
        
                  If DataProduccion.Recordset.EOF Then
                      DataProduccion.Recordset.MoveLast
                      If Err = 3021 Then
                          mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                      End If
                  End If
                  
                  If Err <> 0 Then
                      MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                      Exit Sub
                  End If
    'IMPRIMIR
    ElseIf Index = 5 Then
        MousePointer = 11
                VDia = Day(TxtTexto.Item(2).Text)
                VMes = Month(TxtTexto.Item(2).Text)
                VAño = Year(TxtTexto.Item(2).Text)
                Set RReporteIdentificacionInterno = Db.OpenRecordset("Select * From ReporteIdentificacionInterno Where Fec_Prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# and Linea = '" & TxtTexto.Item(0).Text & "' and Tarima = " & TxtTexto.Item(4).Text & " and Esp_Tec = '" & TxtTexto.Item(1).Text & "'")
                    If RReporteIdentificacionInterno.RecordCount > 0 Then
                            RReporteIdentificacionInterno.Edit
                                    RReporteIdentificacionInterno!Linea = TxtTexto.Item(0).Text
                                    RReporteIdentificacionInterno!Esp_tec = TxtTexto.Item(1).Text
                                    RReporteIdentificacionInterno!Fec_Prd = TxtTexto.Item(2).Text
                                    RReporteIdentificacionInterno!Tarima = TxtTexto.Item(4).Text
                                    RReporteIdentificacionInterno!Envases = TxtTexto.Item(6).Text
                                    RReporteIdentificacionInterno!Hor_Prd = TxtTexto.Item(3).Text
                                    RReporteIdentificacionInterno!Batch = TxtTexto.Item(5).Text
                                    RReporteIdentificacionInterno!Cod_emp = TxtTexto.Item(9).Text
                                    RReporteIdentificacionInterno!Hojalata = ""
                                    RReporteIdentificacionInterno!Fondo = ""
                                    RReporteIdentificacionInterno!Orden = ""
                            RReporteIdentificacionInterno.Update
                    Else
                            RReporteIdentificacionInterno.AddNew
                                    RReporteIdentificacionInterno!Linea = TxtTexto.Item(0).Text
                                    RReporteIdentificacionInterno!Esp_tec = TxtTexto.Item(1).Text
                                    RReporteIdentificacionInterno!Fec_Prd = TxtTexto.Item(2).Text
                                    RReporteIdentificacionInterno!Tarima = TxtTexto.Item(4).Text
                                    RReporteIdentificacionInterno!Envases = TxtTexto.Item(6).Text
                                    RReporteIdentificacionInterno!Hor_Prd = TxtTexto.Item(3).Text
                                    RReporteIdentificacionInterno!Batch = TxtTexto.Item(5).Text
                                    RReporteIdentificacionInterno!Cod_emp = TxtTexto.Item(9).Text
                                    RReporteIdentificacionInterno!Hojalata = ""
                                    RReporteIdentificacionInterno!Fondo = ""
                                    RReporteIdentificacionInterno!Orden = ""
                            RReporteIdentificacionInterno.Update
                    End If
                    
                
                CrReportes.SelectionFormula = "{ReporteIdentificacionInterno.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño & "," & VMes & "," & VDia & ") and {ReporteIdentificacionInterno.Linea} = '" & TxtTexto.Item(0).Text & "' and {ReporteIdentificacionInterno.Tarima} = " & TxtTexto.Item(4).Text & " and {ReporteIdentificacionInterno.Esp_Tec} = '" & TxtTexto.Item(1).Text & "'"
                CrReportes.ReportFileName = App.Path & "\Identificacion.rpt"
                
        MousePointer = 0
                CrReportes.Action = 1
            
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
                'BORRA LA IDENTIFICACION INGRESADA A LA BASE DE DATOS
                Db.Execute "Delete * From ReporteIdentificacionInterno"
    'SALIDA
    ElseIf Index = 6 Then
        Unload Me
    End If
End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
MousePointer = 11

            'FECHAS
            If OptOpcion.Item(0).Value = True Then
                    DataProduccion.RecordSource = ("Select * from ProduccionEncajada where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#")
            'FECHAS Y LINEA
            ElseIf OptOpcion.Item(1).Value = True Then
                    DataProduccion.RecordSource = ("Select * from ProduccionEncajada where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBuscar.Text & "'")
            'FECHAS Y FICHA TECNICA
            ElseIf OptOpcion.Item(2).Value = True Then
                    DataProduccion.RecordSource = ("Select * from ProduccionEncajada where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Esp_Tec = '" & TxtBuscar.Text & "'")
            'BATCH
            ElseIf OptOpcion.Item(4).Value = True Then
                    If Not IsNumeric(TxtBuscar.Text) Then
                        MsgBox "Numero De Batch Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    Else
                        DataProduccion.RecordSource = ("Select * from ProduccionEncajada where batch = " & TxtBuscar.Text)
                    End If
            'NO MP9301
            ElseIf OptOpcion.Item(5).Value = True Then
                If Not IsNumeric(TxtBuscar.Text) Then
                        MsgBox "Numero De MP9301 Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                Else
                        DataProduccion.RecordSource = ("Select * from ProduccionEncajada where NoMP9301 = " & TxtBuscar.Text)
                End If
            End If
                DataProduccion.Refresh
                DBGridProduccion.Refresh

MousePointer = 0

            If Err <> 0 Then
                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                Exit Sub
            End If
            
            SSTab1.Tab = 1

End Sub


Private Sub Form_Activate()
    DataProduccion.RecordSource = "Select * From ProduccionEncajada Order By Fec_Prd, Hor_Prd"
    DataProduccion.Refresh
    If DataProduccion.Recordset.EOF = True Then
    Else
        DataProduccion.Recordset.MoveLast
    End If
    
    
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        'FECHAS
        If OptOpcion.Item(0).Value = True Then
            TxtBuscar.Visible = False
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            Lbletiqueta.Item(0).Caption = ""
            Lbletiqueta.Item(2).Visible = True
            Lbletiqueta.Item(3).Visible = True
        'FECHAS Y LINEA
        ElseIf OptOpcion.Item(1).Value = True Then
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            Lbletiqueta.Item(0).Caption = "Numero Linea"
            Lbletiqueta.Item(2).Visible = True
            Lbletiqueta.Item(3).Visible = True
        'FECHAS Y FICHA TECNICA
        ElseIf OptOpcion.Item(2).Value = True Then
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            Lbletiqueta.Item(0).Caption = "Ficha Tecnica"
            Lbletiqueta.Item(2).Visible = True
            Lbletiqueta.Item(3).Visible = True
        'BATCH
        ElseIf OptOpcion.Item(4).Value = True Then
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            DtpFecIni.Visible = False
            DtpFecFin.Visible = False
            Lbletiqueta.Item(0).Caption = "Numero De Batch"
            Lbletiqueta.Item(2).Visible = False
            Lbletiqueta.Item(3).Visible = False
        'MP9 301
        ElseIf OptOpcion.Item(5).Value = True Then
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            DtpFecIni.Visible = False
            DtpFecFin.Visible = False
            Lbletiqueta.Item(0).Caption = "MP9 301"
            Lbletiqueta.Item(2).Visible = False
            Lbletiqueta.Item(3).Visible = False
        End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next

If SSTab1.Tab = 2 Then
    DtpFecIni.Value = Date
    DtpFecFin.Value = Date
End If

End Sub


Private Sub Text1_GotFocus()
    
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


Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_DblClick()
    'PARA SELECCIONAR LA LINEA
    If VLineas = True Then
        TxtTexto.Item(0).Text = DBGridConsultas.Columns(0)
        TxtTexto.Item(0).SetFocus
        FrameConsultas.Visible = False
    End If
    
    
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 43 Then
          'PARA SELECCIONAR LA LINEA
          If VLineas = True Then
              TxtTexto.Item(0).Text = DBGridConsultas.Columns(0)
              TxtTexto.Item(0).SetFocus
              FrameConsultas.Visible = False
          End If
    End If
End Sub

Private Sub DbgridProduccion_HeadClick(ByVal ColIndex As Integer)
    DataProduccion.RecordSource = ("Select * from ProduccionEncajada order by " & DBGridProduccion.Columns(ColIndex).DataField)
    DataProduccion.Refresh
    DBGridProduccion.Refresh
    
End Sub

Private Sub Form_Load()
    DataProduccion.Connect = GConnect
    DataProduccion.Connect = GConnect
    DataProduccion.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
    
    If GEditar = True Then
        DBGridProduccion.AllowUpdate = True
    Else
        DBGridProduccion.AllowUpdate = False
    End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
    'LINEA
    If Index = 0 Then
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(0).Text & "'")
        If RBuscaLinea.RecordCount > 0 Then
            LblLinea.Caption = RBuscaLinea!Descrip
        Else
            LblLinea.Caption = ""
        End If
    End If
    
    'FICHA TECNICA
    If Index = 1 Then
        Set RBuscaFichaTecnica = Db.OpenRecordset("Select * From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(1).Text & "'")
        If RBuscaFichaTecnica.RecordCount > 0 Then
                
                LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
        Else
                LblFichaTecnica.Caption = ""
        End If
    End If
    
    'BATCH
    If Index = 5 Then
        If IsNumeric(TxtTexto.Item(7).Text) Then
            Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From ProduccionEncajada Where Batch = " & TxtTexto.Item(5).Text)
                If RCuentaTarimas.RecordCount > 0 Then
                    LblBatch.Caption = RCuentaTarimas(0)
                Else
                    LblBatch.Caption = 1
                End If
        End If
    End If
    
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    'LINEA
    If Index = 0 Then
        VLineas = True
        
        DataConsultas.RecordSource = ("Select * from Lineas")
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        FrameConsultas.Visible = True
        DBGridConsultas.SetFocus
        TxtTexto.Item(0).Text = ""
        
    End If 'FIN DE INDICE
    

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
    TxtTexto.Item(Index).SelStart = 0
    TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index))
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    'SI PRECIONAN A ENTER EN CUALQUIER TEXT
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
            'LINEA
                If Index = 0 Then
                    VLineas = True
                  
                    
                    DataConsultas.RecordSource = ("Select * from Lineas")
                    DataConsultas.Refresh
                    DBGridConsultas.Refresh
                    FrameConsultas.Visible = True
                    DBGridConsultas.SetFocus
                    TxtTexto.Item(0).Text = ""
                End If 'FIN DE INDICE
    
    
    End If
    
    
End Sub

Private Sub TxtTexto_LostFocus(Index As Integer)
    'LINEA
    If Index = 0 Then
        'SI NO ESTA EDITANDO BUSCA LOS ULTIMOS DATOS
        If BEditar = False Then
                If TxtTexto.Item(0).Text = "+" Then
                
                ElseIf TxtTexto.Item(0).Text = "" Then
                
                Else
                                
                                'VERIFICA SI LA FICHA TECNICA ESTA ACTIVA
                                Set RLineas = Db.OpenRecordset("Select Esp_Tec, Tarima, Orden from Lineas Where Linea = '" & TxtTexto.Item(0).Text & "' and Activa = -1")
                                'SI LA LINEA ESTA ACTIVA
                                If RLineas.RecordCount > 0 Then
                                                        'FICHA TECNICA
                                                        TxtTexto.Item(1).Text = RLineas!Esp_tec
                                                        'TARIMA
                                                        TxtTexto.Item(4).Text = Val(RLineas!Tarima) + 1
                                                                                                                
                                                        'BUSCA LA FICHA TECNICA Y JALA LOS CODIGOS DE ALAMBRE BARNIZES SELLO Y NYLON
                                                        Set RBuscaFichaTecnica = Db.OpenRecordset("Select * From FichaTecnica Where Esp_Tec = '" & RLineas!Esp_tec & "'")
                                                        
                                                        'SI ENCUENTRA LA FICHA TECNICA
                                                        If RBuscaFichaTecnica.RecordCount > 0 Then
                                                            'ENVASES
                                                            TxtTexto.Item(6).Text = RBuscaFichaTecnica!Envases
                                                        End If
                                                                                                        
                                                        'BUSCA EL ULTIMO REGISTRO INGRESADO Y EXTRAE LOS DATOS
                                                        Set RBuscaProduccion = Db.OpenRecordset("Select * From ProduccionEncajada Where Linea = '" & TxtTexto.Item(0).Text & "' and Esp_Tec = '" & RLineas!Esp_tec & "' and Fec_Prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# Order By Tarima")
                                                        
                                                        If RBuscaProduccion.RecordCount > 0 Then
                                                                'SE MUEVE AL ULTIMO REGISTRO
                                                                RBuscaProduccion.MoveLast
                                                
                                                                'TURNO
                                                                If Not IsNull(RBuscaProduccion!Turno) Then
                                                                    TxtTur.Text = RBuscaProduccion!Turno
                                                                End If
                                                
                                                                'BATCH
                                                                If Not IsNull(RBuscaProduccion!Batch) Then
                                                                    TxtTexto.Item(5).Text = RBuscaProduccion!Batch
                                                                End If
                                                                
                                                                'CUENTA CUANTAS TARIMAS LLEVA EL BATCH
                                                                Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From ProduccionEncajada Where Batch = " & TxtTexto.Item(5).Text)
                                                                If RCuentaTarimas.RecordCount > 0 Then
                                                                            LblBatch.Caption = RCuentaTarimas(0)
                                                                Else
                                                                            LblBatch.Caption = 1
                                                                End If
                                                                                                                        
                                                                'USUARIO
                                                                If Not IsNull(RBuscaProduccion!Usuario) Then
                                                                    TxtTexto.Item(9).Text = RBuscaProduccion!Usuario
                                                                End If
                                                                                                                                
                                                        End If
                                                            
                
                            Else
                                    MsgBox "Esta Linea No Esta Activa", vbOKOnly + vbExclamation, "Informacion"
                            End If
                End If
        End If
    'FECHA
    ElseIf Index = 2 Then
        TxtTexto.Item(2).Text = Format(TxtTexto.Item(2).Text, "dd/mm/yyyy")
        
                    'BUSCA EL ULTIMO REGISTRO INGRESADO Y EXTRAE LOS DATOS
                    Set RBuscaProduccion = Db.OpenRecordset("Select * From ProduccionEncajada Where Linea = '" & TxtTexto.Item(0).Text & "' and Esp_Tec = '" & RLineas!Esp_tec & "' and Fec_Prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# Order By Tarima")
                                                        
                        If RBuscaProduccion.RecordCount > 0 Then
                               'SE MUEVE AL ULTIMO REGISTRO
                                RBuscaProduccion.MoveLast
                                                
                               'TURNO
                               If Not IsNull(RBuscaProduccion!Turno) Then
                                  TxtTur.Text = RBuscaProduccion!Turno
                               End If
                                                
                               'BATCH
                               If Not IsNull(RBuscaProduccion!Batch) Then
                                  TxtTexto.Item(5).Text = RBuscaProduccion!Batch
                               End If
                                                                
                               'CUENTA CUANTAS TARIMAS LLEVA EL BATCH
                               Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From ProduccionEncajada Where Batch = " & TxtTexto.Item(5).Text)
                               If RCuentaTarimas.RecordCount > 0 Then
                                  LblBatch.Caption = RCuentaTarimas(0)
                               Else
                                  LblBatch.Caption = 1
                               End If
                                                                                                                        
                               'USUARIO
                               If Not IsNull(RBuscaProduccion!Usuario) Then
                                      TxtTexto.Item(9).Text = RBuscaProduccion!Usuario
                               End If
                      End If
                
    'BATCH
    ElseIf Index = 5 Then
        If IsNumeric(TxtTexto.Item(5).Text) Then
            Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From ProduccionEncajada Where Batch = " & TxtTexto.Item(5).Text)
                If RCuentaTarimas.RecordCount > 0 Then
                    LblBatch.Caption = RCuentaTarimas(0)
                Else
                    LblBatch.Caption = 1
                End If
        Else
            TxtTexto.Item(5).Text = "0"
        End If
    End If
End Sub

Private Sub TxtTur_GotFocus()
        TxtTur.SelStart = 0
        TxtTur.SelLength = Len(TxtTur.Text)
End Sub

Private Sub TxtTur_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub
