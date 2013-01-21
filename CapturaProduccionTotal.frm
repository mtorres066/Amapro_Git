VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CapturaProduccionTotal 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura de Produccion Liberada"
   ClientHeight    =   7995
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "CapturaProduccionTotal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   7935
      Left            =   0
      TabIndex        =   48
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
         Picture         =   "CapturaProduccionTotal.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "CapturaProduccionTotal.frx":0BD4
         Height          =   7575
         Left            =   120
         OleObjectBlob   =   "CapturaProduccionTotal.frx":0BF0
         TabIndex        =   49
         Top             =   240
         Width           =   10455
      End
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Im&primir"
      Height          =   555
      Left            =   8040
      MouseIcon       =   "CapturaProduccionTotal.frx":15CB
      Picture         =   "CapturaProduccionTotal.frx":1A0D
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   6960
      Width           =   1500
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   555
      Left            =   9600
      MouseIcon       =   "CapturaProduccionTotal.frx":1F3F
      Picture         =   "CapturaProduccionTotal.frx":2381
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   " "
      Top             =   6960
      Width           =   1500
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   555
      Left            =   6480
      MouseIcon       =   "CapturaProduccionTotal.frx":43F3
      Picture         =   "CapturaProduccionTotal.frx":4835
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   " "
      Top             =   6960
      Width           =   1500
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   4920
      MouseIcon       =   "CapturaProduccionTotal.frx":4D67
      Picture         =   "CapturaProduccionTotal.frx":51A9
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   " "
      Top             =   6960
      Width           =   1500
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3360
      MouseIcon       =   "CapturaProduccionTotal.frx":56DB
      Picture         =   "CapturaProduccionTotal.frx":5B1D
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   " "
      Top             =   6960
      Width           =   1500
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   555
      Left            =   1800
      MouseIcon       =   "CapturaProduccionTotal.frx":604F
      Picture         =   "CapturaProduccionTotal.frx":6491
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   " "
      Top             =   6960
      Width           =   1500
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   555
      Left            =   240
      MouseIcon       =   "CapturaProduccionTotal.frx":69C3
      Picture         =   "CapturaProduccionTotal.frx":6E05
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   " "
      Top             =   6960
      Width           =   1500
   End
   Begin VB.Data DataProduccion 
      Caption         =   "Produccion Liberada"
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
      RecordSource    =   "ProduccionTotal"
      Top             =   7560
      Width           =   11535
   End
   Begin TabDlg.SSTab TabProduccionTotal 
      Height          =   6855
      Left            =   0
      TabIndex        =   50
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12091
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CapturaProduccionTotal.frx":7337
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameProduccion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaProduccionTotal.frx":7651
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridProduccion"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Seleccion O Busqueda De Datos"
      TabPicture(2)   =   "CapturaProduccionTotal.frx":7AA3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPFecFin"
      Tab(2).Control(1)=   "DTPFecIni"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "CmdBusqueda"
      Tab(2).Control(4)=   "CmdBuscar"
      Tab(2).Control(5)=   "FrameBuscar"
      Tab(2).Control(6)=   "LblFecFin"
      Tab(2).Control(7)=   "LblFecIni"
      Tab(2).Control(8)=   "LblEtiqueta"
      Tab(2).ControlCount=   9
      Begin MSComCtl2.DTPicker DTPFecFin 
         Height          =   255
         Left            =   -65160
         TabIndex        =   44
         Top             =   3480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24707075
         CurrentDate     =   37335
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   255
         Left            =   -65160
         TabIndex        =   43
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24707075
         CurrentDate     =   37335
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -65640
         TabIndex        =   45
         ToolTipText     =   " "
         Top             =   4200
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CommandButton CmdBusqueda 
         Caption         =   "Seleccionar Datos"
         Height          =   855
         Left            =   -67200
         Picture         =   "CapturaProduccionTotal.frx":7EF5
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   4800
         Width           =   3495
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   945
         Left            =   -67200
         Picture         =   "CapturaProduccionTotal.frx":8337
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5760
         Width           =   3525
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
         Height          =   1815
         Left            =   -74640
         TabIndex        =   39
         Top             =   960
         Width           =   4935
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas Y Linea"
            ForeColor       =   &H8000000D&
            Height          =   1035
            Index           =   2
            Left            =   1680
            Picture         =   "CapturaProduccionTotal.frx":8641
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   480
            Width           =   1500
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas"
            ForeColor       =   &H8000000D&
            Height          =   1035
            Index           =   1
            Left            =   120
            Picture         =   "CapturaProduccionTotal.frx":894B
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   480
            Value           =   -1  'True
            Width           =   1500
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Batch"
            ForeColor       =   &H8000000D&
            Height          =   1035
            Index           =   3
            Left            =   3240
            Picture         =   "CapturaProduccionTotal.frx":8C55
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   480
            Width           =   1500
         End
      End
      Begin VB.Frame FrameProduccion 
         Caption         =   "Datos Generales Captura de Produccion"
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
         Height          =   6045
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   11535
         Begin VB.TextBox TxtFicTec 
            Appearance      =   0  'Flat
            DataField       =   "ESP_TEC"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   840
            MaxLength       =   12
            TabIndex        =   4
            Top             =   1800
            Width           =   1335
         End
         Begin VB.TextBox MskEnvCom 
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            DataField       =   "EnvCom"
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
            Left            =   9600
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox MskEnvIncRec 
            Appearance      =   0  'Flat
            BackColor       =   &H00FF8080&
            DataField       =   "EnvasesIncRec"
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
            Height          =   330
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox TxtCanDef4 
            Appearance      =   0  'Flat
            DataField       =   "Cantidad4"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            TabIndex        =   23
            Top             =   4440
            Width           =   1335
         End
         Begin VB.TextBox TxtCanDef3 
            Appearance      =   0  'Flat
            DataField       =   "Cantidad3"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            TabIndex        =   21
            Top             =   4080
            Width           =   1335
         End
         Begin VB.TextBox TxtCanDef2 
            Appearance      =   0  'Flat
            DataField       =   "Cantidad2"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            TabIndex        =   19
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox TxtCanDef1 
            Appearance      =   0  'Flat
            DataField       =   "Cantidad1"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            TabIndex        =   17
            Top             =   3360
            Width           =   1335
         End
         Begin VB.TextBox TxtDefecto4 
            Appearance      =   0  'Flat
            DataField       =   "Defecto4"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   22
            Top             =   4440
            Width           =   1335
         End
         Begin VB.TextBox TxtDefecto3 
            Appearance      =   0  'Flat
            DataField       =   "Defecto3"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   20
            Top             =   4080
            Width           =   1335
         End
         Begin VB.TextBox TxtDefecto2 
            Appearance      =   0  'Flat
            DataField       =   "Defecto2"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   18
            Top             =   3720
            Width           =   1335
         End
         Begin VB.TextBox TxtDefecto1 
            Appearance      =   0  'Flat
            DataField       =   "Defecto1"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   2400
            MaxLength       =   4
            TabIndex        =   16
            Top             =   3360
            Width           =   1335
         End
         Begin VB.TextBox TxtCalRI 
            Appearance      =   0  'Flat
            DataField       =   "CalidadRI"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox TxtDes 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            DataField       =   "Desperdicio"
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
            Height          =   330
            Left            =   4920
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   4800
            Width           =   1335
         End
         Begin VB.TextBox TxtEnvLib 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            DataField       =   "EnvasesLiberados"
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
            Height          =   330
            Left            =   4920
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   5160
            Width           =   1335
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   53
            TabStop         =   0   'False
            Top             =   5160
            Width           =   2055
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   4200
            Width           =   2055
         End
         Begin VB.TextBox TxtFecPro 
            Appearance      =   0  'Flat
            DataField       =   "FEC_PRD"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   840
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox TxtCodEmp 
            Appearance      =   0  'Flat
            DataField       =   "COD_EMP"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   840
            MaxLength       =   10
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox TxtEnv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H0000C000&
            DataField       =   "ENVASES"
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
            Height          =   330
            Left            =   840
            TabIndex        =   7
            Top             =   2835
            Width           =   1335
         End
         Begin VB.TextBox TxtBat 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "BATCH"
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
            ForeColor       =   &H00008000&
            Height          =   285
            Left            =   840
            TabIndex        =   6
            Top             =   2520
            Width           =   1335
         End
         Begin VB.TextBox TxtTar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            DataField       =   "TARIMA"
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
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   1440
            TabIndex        =   5
            Top             =   2160
            Width           =   735
         End
         Begin VB.TextBox TxtHorPro 
            Appearance      =   0  'Flat
            DataField       =   "HOR_PRD"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   840
            MaxLength       =   5
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            DataField       =   "LINEA"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   840
            MaxLength       =   2
            TabIndex        =   1
            ToolTipText     =   "Busca La Tarima"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtlinIncRec 
            Appearance      =   0  'Flat
            DataField       =   "LineaIncRec"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   10
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtFicTecIncRec 
            Appearance      =   0  'Flat
            DataField       =   "FichaTecnicaIncRec"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            MaxLength       =   12
            TabIndex        =   12
            Top             =   1440
            Width           =   1335
         End
         Begin VB.TextBox TxtLinCom 
            Appearance      =   0  'Flat
            DataField       =   "LineaCom"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   9600
            MaxLength       =   2
            TabIndex        =   26
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtFicTecCom 
            Appearance      =   0  'Flat
            DataField       =   "FichaTecnicaCom"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   9600
            MaxLength       =   12
            TabIndex        =   28
            Top             =   1440
            Width           =   1335
         End
         Begin VB.ComboBox CboCal 
            DataField       =   "CALIDAD"
            DataSource      =   "DataProduccion"
            Height          =   315
            ItemData        =   "CapturaProduccionTotal.frx":9097
            Left            =   840
            List            =   "CapturaProduccionTotal.frx":90A7
            TabIndex        =   8
            Text            =   "A"
            Top             =   3240
            Width           =   1335
         End
         Begin MSMask.MaskEdBox MskTarCom 
            DataField       =   "TarimaCom"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   9600
            TabIndex        =   29
            ToolTipText     =   "Resta Envases Menos Envases a Liberar"
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecCom 
            DataField       =   "FechaTarCom"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   9600
            TabIndex        =   27
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/MM/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTarIncRec 
            DataField       =   "TarimaIncRec"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            TabIndex        =   13
            ToolTipText     =   "Busca La Tarima"
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecIncRec 
            DataField       =   "FechaTarIncRec"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   4920
            TabIndex        =   11
            Top             =   1080
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/MM/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Shape Shape1 
            Height          =   5055
            Left            =   2280
            Top             =   600
            Width           =   4215
         End
         Begin VB.Shape ShapeCalidad 
            BackStyle       =   1  'Opaque
            Height          =   495
            Left            =   4080
            Shape           =   3  'Circle
            Top             =   1920
            Width           =   375
         End
         Begin VB.Label LblDefectos 
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
            Index           =   3
            Left            =   6720
            TabIndex        =   86
            Top             =   4440
            Width           =   4695
         End
         Begin VB.Label LblDefectos 
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
            Index           =   1
            Left            =   6720
            TabIndex        =   85
            Top             =   3720
            Width           =   4695
         End
         Begin VB.Label LblDefectos 
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
            Index           =   2
            Left            =   6720
            TabIndex        =   84
            Top             =   4080
            Width           =   4695
         End
         Begin VB.Label LblDefectos 
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
            Index           =   0
            Left            =   6720
            TabIndex        =   83
            Top             =   3360
            Width           =   4695
         End
         Begin VB.Label Label11 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Cantidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   4920
            TabIndex        =   82
            Top             =   3000
            Width           =   1095
         End
         Begin VB.Label Label11 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Defecto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   81
            Top             =   3000
            Width           =   975
         End
         Begin VB.Label Label13 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Calidad"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   80
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Envases"
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
            Left            =   7320
            TabIndex        =   54
            Top             =   2160
            Width           =   975
         End
         Begin VB.Label Label11 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Envases Rechazados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   2400
            TabIndex        =   55
            Top             =   2520
            Width           =   2295
         End
         Begin VB.Label Label11 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Desperdicio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   2400
            TabIndex        =   79
            Top             =   4800
            Width           =   1815
         End
         Begin VB.Label Label11 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Envases A Liberar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   78
            Top             =   5160
            Width           =   2055
         End
         Begin VB.Label LblBatch 
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   840
            TabIndex        =   77
            Top             =   2160
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label sdf 
            Caption         =   "Usuario"
            Height          =   255
            Left            =   120
            TabIndex        =   76
            Top             =   3600
            Width           =   1095
         End
         Begin VB.Label Label26 
            BackStyle       =   0  'Transparent
            Caption         =   "Fondo ó Tapa"
            Height          =   255
            Left            =   120
            TabIndex        =   75
            Top             =   4920
            Width           =   1215
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Hojalata"
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   3960
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Calidad"
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   73
            Top             =   3240
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Envases"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   72
            Top             =   2880
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Batch"
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   71
            Top             =   2520
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tarima"
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   70
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Ficha Tecnica"
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   69
            Top             =   1680
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Hora"
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   67
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Linea"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   66
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Linea"
            Height          =   255
            Left            =   2400
            TabIndex        =   65
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label2 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   255
            Left            =   2400
            TabIndex        =   64
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label3 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha Tecnica"
            Height          =   255
            Left            =   2400
            TabIndex        =   63
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label4 
            BackColor       =   &H00008000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tarima"
            Height          =   255
            Left            =   2400
            TabIndex        =   62
            Top             =   1800
            Width           =   735
         End
         Begin VB.Label Label5 
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Linea"
            Height          =   255
            Left            =   7320
            TabIndex        =   61
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label6 
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha"
            Height          =   255
            Left            =   7320
            TabIndex        =   60
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label7 
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha Tecnica"
            Height          =   255
            Left            =   7320
            TabIndex        =   59
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label8 
            BackColor       =   &H00808000&
            BackStyle       =   0  'Transparent
            Caption         =   "Tarima"
            Height          =   255
            Left            =   7320
            TabIndex        =   58
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label Label9 
            Caption         =   "TARIMAS RECHAZADAS O INCOMPLETAS"
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
            Left            =   2280
            TabIndex        =   57
            Top             =   240
            Width           =   4095
         End
         Begin VB.Label Label10 
            Caption         =   "TARIMAS DE COMPLEMENTO"
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
            Left            =   6720
            TabIndex        =   56
            Top             =   240
            Width           =   3855
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H00808000&
            Height          =   2055
            Left            =   6720
            Top             =   600
            Width           =   4695
         End
      End
      Begin MSDBGrid.DBGrid DBGridProduccion 
         Bindings        =   "CapturaProduccionTotal.frx":90B7
         Height          =   6135
         Left            =   -74880
         OleObjectBlob   =   "CapturaProduccionTotal.frx":90D4
         TabIndex        =   38
         Top             =   720
         Width           =   11535
      End
      Begin VB.Label LblFecFin 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   -66360
         TabIndex        =   89
         Top             =   3480
         Width           =   1005
      End
      Begin VB.Label LblFecIni 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         Left            =   -66360
         TabIndex        =   88
         Top             =   3000
         Width           =   1110
      End
      Begin VB.Label LblEtiqueta 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   -67440
         TabIndex        =   87
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   120
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
End
Attribute VB_Name = "CapturaProduccionTotal"
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
Dim RBuscaTarima As Recordset
Dim RBuscaUltimaFicha As Recordset
Dim RBuscaFondo As Recordset
Dim RBuscaPlatina As Recordset
Dim RReporteIdentificacionInterno As Recordset

'VARIABLES PARA DESPLEGAR DATOS DE FICHA TECNICA
Dim RBuscaFichaTecnica As Recordset

Dim RBuscaAtributo As Recordset

Dim VLineas As Boolean
Dim VDia As String
Dim VMes As String
Dim VAño As String

Dim RCuentaTarimas As Recordset

Dim VEditar As Boolean

Dim RVerificaTarima As Recordset
Dim Vtarima As String

Dim RBuscaDefectos As Recordset
Dim RBuscaFoto As Recordset

Dim BLinea As Boolean
Dim BDefecto1 As Boolean
Dim BDefecto2 As Boolean
Dim BDefecto3 As Boolean
Dim BDefecto4 As Boolean


Sub botones()
    If Bandera = True Then
         FrameProduccion.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdBuscar.Enabled = False
         CmdImprimir.Enabled = False
         CmdBusqueda.Enabled = False
         LblBatch.Visible = True
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtFicTec.SetFocus
         FrameBuscar.Visible = False
         DataProduccion.Visible = False
         DBGridProduccion.Visible = False
    Else
         FrameProduccion.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdBuscar.Enabled = True
         CmdImprimir.Enabled = True
         CmdBusqueda.Enabled = True
         LblBatch.Visible = False
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         FrameBuscar.Visible = True
         DataProduccion.Visible = True
         DBGridProduccion.Visible = True

    End If
End Sub


Private Sub CboCal_LostFocus()
   If CboCal.Text <> "A" And CboCal.Text <> "C" And CboCal.Text <> "I" And CboCal.Text <> "R" Then
        MsgBox "CALIDAD INCORRECTA", vbOKOnly + vbInformation, "Informacion"
   End If
   
End Sub

Private Sub CmdBuscar_Click()
    DataProduccion.RecordSource = "Select * from ProduccionTotal"
    DataProduccion.Refresh
    DBGridProduccion.Refresh
End Sub

Private Sub CmdBusqueda_Click()
On Error Resume Next

            
            'FECHAS
            If OptBusqueda.Item(1).Value = True Then
                    DataProduccion.RecordSource = ("Select * from ProduccionTotal where Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#")
            'FECHAS Y LINEA
            ElseIf OptBusqueda.Item(2).Value = True Then
                    DataProduccion.RecordSource = ("Select * from ProduccionTotal where Fec_Prd >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBuscar.Text & "'")
            'BATCH
            ElseIf OptBusqueda.Item(3).Value = True Then
                If TxtBuscar.Text <> "" Then
                    If Not IsNumeric(TxtBuscar.Text) Then
                        MsgBox "Numero De Batch Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    Else
                        DataProduccion.RecordSource = ("Select * from ProduccionTotal where batch = " & TxtBuscar.Text)
                    End If
                End If
            End If
                        DataProduccion.Refresh
                        DBGridProduccion.Refresh
                        TabProduccionTotal.Tab = 1
            
            If Err <> 0 Then
                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                Exit Sub
            End If


End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next

                'BORRA LA IDENTIFICACION INGRESADA A LA BASE DE DATOS
                Db.Execute "Delete * From ReporteIdentificacionInterno"

MousePointer = 11
                VDia = Day(TxtFecPro.Text)
                VMes = Month(TxtFecPro.Text)
                VAño = Year(TxtFecPro.Text)
                
                Set RReporteIdentificacionInterno = Db.OpenRecordset("Select * From ReporteIdentificacionInterno Where Fec_Prd = #" & Format(TxtFecPro.Text, "mm/dd/yyyy") & "# and Linea = '" & TxtLin.Text & "' and Tarima = " & TxtTar.Text & " and Esp_Tec = '" & TxtFicTec.Text & "'")
                    If RReporteIdentificacionInterno.RecordCount > 0 Then
                            RReporteIdentificacionInterno.Edit
                                    RReporteIdentificacionInterno!Linea = TxtLin.Text
                                    RReporteIdentificacionInterno!Esp_Tec = TxtFicTec.Text
                                    RReporteIdentificacionInterno!Fec_prd = TxtFecPro.Text
                                    RReporteIdentificacionInterno!Tarima = TxtTar.Text
                                    RReporteIdentificacionInterno!Envases = TxtEnv.Text
                                    RReporteIdentificacionInterno!Hor_prd = TxtHorPro.Text
                                    RReporteIdentificacionInterno!Batch = TxtBat.Text
                                    RReporteIdentificacionInterno!Cod_Emp = TxtCodEmp.Text
                                    RReporteIdentificacionInterno!Hojalata = ""
                                    RReporteIdentificacionInterno!Fondo = ""
                                    RReporteIdentificacionInterno!Orden = ""
                            RReporteIdentificacionInterno.Update
                    Else
                            RReporteIdentificacionInterno.AddNew
                                    RReporteIdentificacionInterno!Linea = TxtLin.Text
                                    RReporteIdentificacionInterno!Esp_Tec = TxtFicTec.Text
                                    RReporteIdentificacionInterno!Fec_prd = TxtFecPro.Text
                                    RReporteIdentificacionInterno!Tarima = TxtTar.Text
                                    RReporteIdentificacionInterno!Envases = TxtEnv.Text
                                    RReporteIdentificacionInterno!Hor_prd = TxtHorPro.Text
                                    RReporteIdentificacionInterno!Batch = TxtBat.Text
                                    RReporteIdentificacionInterno!Cod_Emp = TxtCodEmp.Text
                                    RReporteIdentificacionInterno!Hojalata = ""
                                    RReporteIdentificacionInterno!Fondo = ""
                                    RReporteIdentificacionInterno!Orden = ""
                            RReporteIdentificacionInterno.Update
                    End If
                    
                
                CrReportes.SelectionFormula = "{ReporteIdentificacionInterno.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño & "," & VMes & "," & VDia & ") and {ReporteIdentificacionInterno.Linea} = '" & TxtLin.Text & "' and {ReporteIdentificacionInterno.Tarima} = " & TxtTar.Text & " and {ReporteIdentificacionInterno.Esp_Tec} = '" & TxtFicTec.Text & "'"
                CrReportes.ReportFileName = App.Path & "\Identificacion2.rpt"
                
        MousePointer = 0
                CrReportes.Action = 1
            
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
                

    
End Sub

Private Sub Form_Activate()
'    DataProduccion.RecordSource = ("Select * from Produccion Order BY Fec_prd, Hor_prd")
'    DataProduccion.Refresh
If DataProduccion.Recordset.RecordCount > 0 Then
    DataProduccion.Recordset.MoveLast
End If
End Sub


Private Sub MskEnvCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub MskEnvIncRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub MskFecCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub MskFecIncRec_GotFocus()
        MskFecIncRec.SelStart = 0
        MskFecIncRec.SelLength = Len(MskFecIncRec.Text)
End Sub

Private Sub MskFecIncRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub MskTarCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub MskTarCom_LostFocus()
    MskEnvCom.Text = Val(TxtEnv.Text) - Val(TxtEnvLib.Text)
    CmdGrabar.SetFocus
End Sub

Private Sub MskTarIncRec_GotFocus()
        MskTarIncRec.SelStart = 0
        MskTarIncRec.SelLength = Len(MskTarIncRec.Text)
End Sub

Private Sub MskTarIncRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub MskTarIncRec_LostFocus()
    If IsDate(MskFecIncRec.Text) Then
        Set RBuscaTarima = Db.OpenRecordset("Select Envases, Calidad From Produccion Where Fec_Prd = #" & Format(MskFecIncRec.Text, "mm/dd/yyyy") & "# and Esp_Tec = '" & TxtFicTecIncRec.Text & "' and Linea = '" & TxtlinIncRec.Text & "' and Tarima = " & MskTarIncRec.Text)
        If RBuscaTarima.RecordCount > 0 Then
            MskEnvIncRec.Text = RBuscaTarima!Envases
            TxtCalRI.Text = RBuscaTarima!Calidad
        Else
            MsgBox "No Existe La Tarima " & MskTarIncRec.Text & " En " & MskFecIncRec.Text & " De La Ficha " & TxtFicTecIncRec.Text & " Para La Linea " & TxtlinIncRec.Text, vbOKOnly + vbCritical, "Error"
        End If
    End If
End Sub


Private Sub OptBusqueda_Click(Index As Integer)
        
        If OptBusqueda.Item(1).Value = True Then
            LblEtiqueta.Caption = ""
            DTPFecIni.Visible = True
            DTPFecFin.Visible = True
            LblFecIni.Visible = True
            LblFecFin.Visible = True
            DTPFecIni.Value = Date
            DTPFecFin.Value = Date
            TxtBuscar.Visible = False
        ElseIf OptBusqueda.Item(2).Value = True Then
            LblEtiqueta.Caption = "Linea"
            DTPFecIni.Visible = True
            DTPFecFin.Visible = True
            LblFecIni.Visible = True
            LblFecFin.Visible = True
            DTPFecIni.Value = Date
            DTPFecFin.Value = Date
            TxtBuscar.Visible = True
        ElseIf OptBusqueda.Item(3).Value = True Then
            LblEtiqueta.Caption = "Batch"
            DTPFecIni.Visible = False
            DTPFecFin.Visible = False
            LblFecIni.Visible = False
            LblFecFin.Visible = False
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
        End If
            
End Sub

Private Sub TabProduccionTotal_Click(PreviousTab As Integer)
        DTPFecIni.Value = Date
        DTPFecFin.Value = Date
End Sub



Private Sub TxtBat_LostFocus()
Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From Produccion Where Batch = " & TxtBat.Text)
    If RCuentaTarimas.RecordCount > 0 Then
        LblBatch.Caption = RCuentaTarimas(0)
    Else
        LblBatch.Caption = 1
    End If
End Sub

Private Sub Txtbat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub CboCal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
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

Private Sub TxtCalRI_Change()
    If TxtCalRI.Text = "R" Then
                ShapeCalidad.BackColor = vbRed
    ElseIf TxtCalRI.Text = "I" Then
                ShapeCalidad.BackColor = vbBlue
    Else
                ShapeCalidad.BackColor = vbWhite
    End If
End Sub

Private Sub TxtCalRI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtCalRI_Validate(Cancel As Boolean)
    If TxtCalRI.Text = "R" Then
                ShapeCalidad.BackColor = vbRed
    ElseIf TxtCalRI.Text = "I" Then
                ShapeCalidad.BackColor = vbBlue
    Else
                ShapeCalidad.BackColor = vbWhite
    End If

End Sub

Private Sub TxtCanDef1_GotFocus()
        TxtCanDef1.SelStart = 0
        TxtCanDef1.SelLength = Len(TxtCanDef1.Text)
End Sub

Private Sub TxtCanDef1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtCanDef1_LostFocus()
        TxtDes.Text = Val(TxtCanDef1) + Val(TxtCanDef2) + Val(TxtCanDef3) + Val(TxtCanDef4)
        
        If IsNumeric(TxtDes.Text) Then
            If IsNumeric(MskEnvIncRec.Text) Then
                TxtEnvLib.Text = MskEnvIncRec - TxtDes.Text
            End If
        Else
                TxtEnvLib.Text = 0
        End If

End Sub

Private Sub TxtCanDef2_GotFocus()
        TxtCanDef2.SelStart = 0
        TxtCanDef2.SelLength = Len(TxtCanDef2.Text)
End Sub

Private Sub TxtCanDef2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtCanDef2_LostFocus()
        TxtDes.Text = Val(TxtCanDef1) + Val(TxtCanDef2) + Val(TxtCanDef3) + Val(TxtCanDef4)
        
        If IsNumeric(TxtDes.Text) Then
            If IsNumeric(MskEnvIncRec.Text) Then
                TxtEnvLib.Text = MskEnvIncRec - TxtDes.Text
            End If
        Else
                TxtEnvLib.Text = 0
        End If
End Sub

Private Sub TxtCanDef3_GotFocus()
        TxtCanDef3.SelStart = 0
        TxtCanDef3.SelLength = Len(TxtCanDef3.Text)
End Sub

Private Sub TxtCanDef3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtCanDef3_LostFocus()
        TxtDes.Text = Val(TxtCanDef1) + Val(TxtCanDef2) + Val(TxtCanDef3) + Val(TxtCanDef4)
        
        If IsNumeric(TxtDes.Text) Then
            If IsNumeric(MskEnvIncRec.Text) Then
                TxtEnvLib.Text = MskEnvIncRec - TxtDes.Text
            End If
        Else
                TxtEnvLib.Text = 0
        End If
        
End Sub

Private Sub TxtCanDef4_GotFocus()
        TxtCanDef4.SelStart = 0
        TxtCanDef4.SelLength = Len(TxtCanDef4.Text)
End Sub

Private Sub TxtCanDef4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtCanDef4_LostFocus()
        TxtDes.Text = Val(TxtCanDef1) + Val(TxtCanDef2) + Val(TxtCanDef3) + Val(TxtCanDef4)
        
        If IsNumeric(TxtDes.Text) Then
            If IsNumeric(MskEnvIncRec.Text) Then
                TxtEnvLib.Text = MskEnvIncRec - TxtDes.Text
            End If
        Else
                TxtEnvLib.Text = 0
        End If
        
End Sub


Private Sub TxtCodEmp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub
Private Sub TxtDefecto1_Change()
            Set RBuscaDefectos = Db.OpenRecordset("Select Descrip From Defectos Where Defecto = '" & TxtDefecto1.Text & "'")
            If RBuscaDefectos.RecordCount > 0 Then
                    LblDefectos.Item(0).Caption = RBuscaDefectos!Descrip
            Else
                    LblDefectos.Item(0).Caption = ""
            End If
End Sub

Private Sub TxtDefecto1_DblClick()

    BLinea = False
    BDefecto1 = True
    BDefecto2 = False
    BDefecto3 = False
    BDefecto4 = False
    DataConsultas.RecordSource = ("Select * from Defectos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End Sub

Private Sub TxtDefecto1_GotFocus()
        TxtDefecto1.SelStart = 0
        TxtDefecto1.SelLength = Len(TxtDefecto1.Text)
End Sub

Private Sub TxtDefecto1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

If KeyAscii = 43 Then
    BLinea = False
    BDefecto1 = True
    BDefecto2 = False
    BDefecto3 = False
    BDefecto4 = False
    
    DataConsultas.RecordSource = ("Select * from Defectos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End If

End Sub

Private Sub TxtDefecto2_Change()
Set RBuscaDefectos = Db.OpenRecordset("Select Descrip From Defectos Where Defecto = '" & TxtDefecto2.Text & "'")
            If RBuscaDefectos.RecordCount > 0 Then
                    LblDefectos.Item(1).Caption = RBuscaDefectos!Descrip
            Else
                    LblDefectos.Item(1).Caption = ""
            End If
End Sub

Private Sub TxtDefecto2_DblClick()
    BLinea = False
    BDefecto1 = False
    BDefecto2 = True
    BDefecto3 = False
    BDefecto4 = False
    
    DataConsultas.RecordSource = ("Select * from Defectos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End Sub

Private Sub TxtDefecto2_GotFocus()
        TxtDefecto2.SelStart = 0
        TxtDefecto2.SelLength = Len(TxtDefecto2.Text)
End Sub

Private Sub TxtDefecto2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

If KeyAscii = 43 Then
    BLinea = False
    BDefecto1 = False
    BDefecto2 = True
    BDefecto3 = False
    BDefecto4 = False
    
    DataConsultas.RecordSource = ("Select * from Defectos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End If

End Sub

Private Sub TxtDefecto3_Change()
Set RBuscaDefectos = Db.OpenRecordset("Select Descrip From Defectos Where Defecto = '" & TxtDefecto3.Text & "'")
            If RBuscaDefectos.RecordCount > 0 Then
                    LblDefectos.Item(2).Caption = RBuscaDefectos!Descrip
            Else
                    LblDefectos.Item(2).Caption = ""
            End If
End Sub

Private Sub TxtDefecto3_DblClick()
    BLinea = False
    BDefecto1 = False
    BDefecto2 = False
    BDefecto3 = True
    BDefecto4 = False
    
    DataConsultas.RecordSource = ("Select * from Defectos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus


End Sub

Private Sub TxtDefecto3_GotFocus()
        TxtDefecto3.SelStart = 0
        TxtDefecto3.SelLength = Len(TxtDefecto3.Text)
End Sub

Private Sub TxtDefecto3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

If KeyAscii = 43 Then
    BLinea = False
    BDefecto1 = False
    BDefecto2 = False
    BDefecto3 = True
    BDefecto4 = False
    
    DataConsultas.RecordSource = ("Select * from Defectos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End If

End Sub

Private Sub TxtDefecto4_Change()
Set RBuscaDefectos = Db.OpenRecordset("Select Descrip From Defectos Where Defecto = '" & TxtDefecto4.Text & "'")
            If RBuscaDefectos.RecordCount > 0 Then
                    LblDefectos.Item(3).Caption = RBuscaDefectos!Descrip
            Else
                    LblDefectos.Item(3).Caption = ""
            End If
End Sub

Private Sub TxtDefecto4_DblClick()
    BLinea = False
    BDefecto1 = False
    BDefecto2 = False
    BDefecto3 = False
    BDefecto4 = True
    
    DataConsultas.RecordSource = ("Select * from Defectos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus


End Sub

Private Sub TxtDefecto4_GotFocus()
        TxtDefecto4.SelStart = 0
        TxtDefecto4.SelLength = Len(TxtDefecto4.Text)
End Sub

Private Sub TxtDefecto4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

If KeyAscii = 43 Then
    BLinea = False
    BDefecto1 = False
    BDefecto2 = False
    BDefecto3 = False
    BDefecto4 = True
    
    DataConsultas.RecordSource = ("Select * from Defectos")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
End If

End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtDes_LostFocus()
    If IsNumeric(TxtDes.Text) Then
        If IsNumeric(MskEnvIncRec.Text) Then
            TxtEnvLib.Text = MskEnvIncRec - TxtDes.Text
        End If
    Else
            TxtEnvLib.Text = 0
    End If

End Sub

Private Sub TxtEnv_GotFocus()
        TxtEnv.SelStart = 0
        TxtEnv.SelLength = Len(TxtEnv.Text)
End Sub

Private Sub TxtEnv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtEnvLib_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtFecPro_LostFocus()
        TxtFecPro.Text = Format(TxtFecPro.Text, "dd/mm/yyyy")
End Sub

Private Sub TxtFicTec_Change()

    Set RBuscaFichaTecnica = Db.OpenRecordset("Select * From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
    If RBuscaFichaTecnica.RecordCount > 0 Then
            
            'PLATINAS
            'Set RBuscaPlatina = Db.OpenRecordset("Select Descrip From Platinas Where Platina = '" & RBuscaFichaTecnica!Platina & "'")
            'If RBuscaPlatina.RecordCount > 0 Then
            '    Text1.Text = RBuscaPlatina(0)
            'Else
            '    Text1.Text = ""
            'End If
            
            'FONDOS
            'Set RBuscaFondo = Db.OpenRecordset("Select Descrip From Fondos Where Fondo = '" & RBuscaFichaTecnica!Fondo & "'")
            'If RBuscaFondo.RecordCount > 0 Then
            '    Text4.Text = RBuscaFondo(0)
            'Else
            '    Text4.Text = ""
            'End If
    End If
        
End Sub

Private Sub TxtFicTec_GotFocus()
        TxtFicTec.SelStart = 0
        TxtFicTec.SelLength = Len(TxtFicTec.Text)
End Sub

Private Sub TxtFicTec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Private Sub TxtFicTec_LostFocus()
    TxtFicTecIncRec.Text = TxtFicTec.Text
End Sub


Private Sub TxtFicTecCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtFicTecIncRec_GotFocus()
        TxtFicTecIncRec.SelStart = 0
        TxtFicTecIncRec.SelLength = Len(TxtFicTecIncRec.Text)
End Sub

Private Sub TxtFicTecIncRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtHorPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub


Private Sub TxtLin_DblClick()
    BLinea = True
    BDefecto1 = False
    BDefecto2 = False
    BDefecto3 = False
    BDefecto4 = False

    DataConsultas.RecordSource = ("Select * from Lineas")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
    TxtLin.Text = ""
End Sub

Private Sub TxtLin_GotFocus()
        TxtLin.SelStart = 0
        TxtLin.SelLength = Len(TxtLin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

If KeyAscii = 43 Then
    BLinea = True
    BDefecto1 = False
    BDefecto2 = False
    BDefecto3 = False
    BDefecto4 = False
    
    DataConsultas.RecordSource = ("Select * from Lineas")
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    DBGridConsultas.Columns(1).Width = "4000"
    FrameConsultas.Visible = True
    DBGridConsultas.SetFocus
    TxtLin.Text = ""
End If

End Sub

Private Sub txtfecpro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
        
        DataProduccion.Recordset.AddNew
        
        If Err <> 0 Then
                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                Exit Sub
        End If
                
        Bandera = True
        botones
        
        CboCal.Text = "A"
        TxtLin.SetFocus
        
        'SI LA HORA ES MENOR QUE LAS 7 DE LA MAÑANA ENTONCES DA LA FECHA ANTERIOR
        If Format(Time, "hh:mm") < "07:00" Then
            TxtFecPro.Text = Format(DateValue(Date) - 1, "dd/mm/yyyy")
        Else
            TxtFecPro.Text = Format(Date, "dd/mm/yyyy")
        End If
                
        TxtHorPro.Text = Format(Time, "hh:mm")
        TxtCodEmp.Text = GUsuario
        
        VEditar = False
        
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

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

            
            
End Sub


Private Sub CmdCancelar_Click()
On Error Resume Next

        DataProduccion.Recordset.CancelUpdate
        
        If Err <> 0 Then
                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                Exit Sub
        End If
        
        Bandera = False
        botones

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
                
        DataProduccion.Recordset.Edit
        
        If Err <> 0 Then
             MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
             Exit Sub
        End If
        
        Bandera = True
        botones
       
        TxtFicTec.SetFocus
        
        VEditar = True
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   If CboCal.Text <> "A" And CboCal.Text <> "C" And CboCal.Text <> "I" And CboCal.Text <> "R" Then
        MsgBox "CALIDAD INCORRECTA", vbOKOnly + vbInformation, "Informacion"
        MousePointer = 0
        Exit Sub
   End If
       
   If Not IsNumeric(TxtBat.Text) Then
        MsgBox "Numero de Batch Incorrecto", vbOKOnly + vbInformation, "Informacion"
        MousePointer = 0
        Exit Sub
   End If
   
    'ENVASES A LIBERAR SON ENVASES DE LA TARIMA R O I MENOS EL DESPERDICIO
    If IsNumeric(TxtDes.Text) Then
        If IsNumeric(MskEnvIncRec.Text) Then
            TxtEnvLib.Text = MskEnvIncRec - TxtDes.Text
        End If
    Else
            TxtEnvLib.Text = 0
    End If
    
    'LA CANTIDAD DE ENVASES DE LA TARIMA DE COMPLEMENTO ES IGUAL A LA CANTIDAD DE ENVASES DE LA NUEVA TARIMA
    'MENOS LOS ENVASES LIBERADOS DE LA TARIMA R O I
    MskEnvCom.Text = Val(TxtEnv.Text) - Val(TxtEnvLib.Text)

   
   Vlinea = TxtLin.Text
   VFicha = TxtFicTec.Text
   VFecha = TxtFecPro.Text
   Vtarima = TxtTar.Text
   
   'VERIFICA LA LINEA
   If Vlinea = "" Then
        MsgBox "Linea No Puede Estar Vacia", vbOKOnly + vbInformation, "Informacion"
        TxtLin.SetFocus
        Exit Sub
   End If
   'VERIFICA LA FICHA TECNICA
   If VFicha = "" Then
        MsgBox "Ficha Tecnica No Puede Estar Vacia", vbOKOnly + vbInformation, "Informacion"
        TxtFicTec.SetFocus
        Exit Sub
   End If
   
   'VERIFICA LA FECHA
   If Not IsDate(VFecha) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
        TxtFecPro.SetFocus
        Exit Sub
   End If
   
   'VERIFICA LA TARIMA
   If Not IsNumeric(Vtarima) Then
        MsgBox "Tarima Incorrecta", vbOKOnly + vbInformation, "Informacion"
        TxtTar.SetFocus
        Exit Sub
   End If
   
   'VERIFICA SI YA EXISTE LA TARIMA
   Set RVerificaTarima = Db.OpenRecordset("Select * from produccion Where Linea = '" & Vlinea & "' and Esp_tec = '" & VFicha & "' and Fec_prd = #" & Format(VFecha, "dd/mm/yyyy") & "# and Tarima = " & Vtarima)
   If RVerificaTarima.RecordCount > 0 Then
        mensaje = MsgBox("Ya Existe Tarima " & Vtarima & " De Ficha " & VFicha & " Con Fecha " & VFecha & " En Produccion Interna", vbOKOnly + vbInformation, "Informacion")
        Exit Sub
   End If
     
   'GRABA DATOS
   DataProduccion.Recordset.Update
   
   If Err <> 0 Then
      MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
      TxtLin.SetFocus
      Exit Sub
   Else
        If VEditar = False Then
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
      
      
      'DataProduccion.RecordSource = ("Select * from Produccion Order BY Fec_prd, Hor_Prd")
      'DataProduccion.Refresh
      DataProduccion.Recordset.MoveLast
      
      CmdAgregar.SetFocus
  End If
            
       
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_DblClick()
    'PARA SELECCIONAR LA PLATINA
    If BLinea = True Then
        TxtLin.Text = DBGridConsultas.Columns(0)
        TxtLin.SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR LA DEFECTO 1
    If BDefecto1 = True Then
        TxtDefecto1.Text = DBGridConsultas.Columns(0)
        TxtDefecto1.SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR LA DEFECTO 2
    If BDefecto2 = True Then
        TxtDefecto2.Text = DBGridConsultas.Columns(0)
        TxtDefecto2.SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR LA DEFECTO 3
    If BDefecto3 = True Then
        TxtDefecto3.Text = DBGridConsultas.Columns(0)
        TxtDefecto3.SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR LA DEFECTO 4
    If BDefecto4 = True Then
        TxtDefecto4.Text = DBGridConsultas.Columns(0)
        TxtDefecto4.SetFocus
        FrameConsultas.Visible = False
    End If
    
    
    
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)

If KeyAscii = 43 Then
'PARA SELECCIONAR LA PLATINA
    If BLinea = True Then
        TxtLin.Text = DBGridConsultas.Columns(0)
        TxtLin.SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR LA DEFECTO 1
    If BDefecto1 = True Then
        TxtDefecto1.Text = DBGridConsultas.Columns(0)
        TxtDefecto1.SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR LA DEFECTO 2
    If BDefecto2 = True Then
        TxtDefecto2.Text = DBGridConsultas.Columns(0)
        TxtDefecto2.SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR LA DEFECTO 3
    If BDefecto3 = True Then
        TxtDefecto3.Text = DBGridConsultas.Columns(0)
        TxtDefecto3.SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR LA DEFECTO 4
    If BDefecto4 = True Then
        TxtDefecto4.Text = DBGridConsultas.Columns(0)
        TxtDefecto4.SetFocus
        FrameConsultas.Visible = False
    End If
    
    
    
End If
End Sub

Private Sub DbgridProduccion_HeadClick(ByVal ColIndex As Integer)
    DataProduccion.RecordSource = ("Select * from ProduccionTotal order by " & DBGridProduccion.Columns(ColIndex).DataField)
    DataProduccion.Refresh
    DBGridProduccion.Refresh
    
End Sub

Private Sub Form_Load()
    DataProduccion.Connect = GConnect
    DataConsultas.Connect = GConnect
    DataProduccion.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
    
    
End Sub


Private Sub TxtLin_LostFocus()
    
If TxtLin.Text = "+" Then
ElseIf TxtLin.Text = "" Then
Else
    Set RLineas = Db.OpenRecordset("Select Esp_Tec, Tarima from Lineas Where Linea = '" & TxtLin.Text & "' and Activa = -1")
    If RLineas.RecordCount > 0 Then
        TxtFicTec.Text = RLineas!Esp_Tec
        TxtFicTecIncRec.Text = RLineas!Esp_Tec
        'TxtFicTecCom.Text = RLineas!Esp_Tec
        TxtTar.Text = Val(RLineas!Tarima) + 1
        
        'BUSCA LA FICHA TECNICA Y JALA LOS CODIGOS DE ALAMBRE BARNIZES SELLO Y NYLON
        Set RBuscaFichaTecnica = Db.OpenRecordset("Select Envases From FichaTecnica Where Esp_Tec = '" & RLineas!Esp_Tec & "'")
        If RBuscaFichaTecnica.RecordCount > 0 Then
            TxtEnv.Text = RBuscaFichaTecnica!Envases
        End If
        
        'SELECCIONA TODOS LOS REGISTROS DE PRODUCCION Y SE VA AL ULTIMO
        'BUSCA EL ULTIMO REGISTRO INGRESADO Y EXTRAE LOS DATOS
        Set RBuscaProduccion = Db.OpenRecordset("Select * From Produccion Where Linea = '" & TxtLin.Text & "' and Esp_Tec = '" & RLineas!Esp_Tec & "' and Fec_Prd = #" & Format(TxtFecPro.Text, "mm/dd/yyyy") & "# Order By Tarima")
        If RBuscaProduccion.RecordCount > 0 Then
                'SE MUEVE AL ULTIMO REGISTRO
                RBuscaProduccion.MoveLast

                'BATCH
                If Not IsNull(RBuscaProduccion!Batch) Then
                    TxtBat.Text = RBuscaProduccion!Batch
                    
                    Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From Produccion Where Batch = " & TxtBat.Text)
                        If RCuentaTarimas.RecordCount > 0 Then
                            LblBatch.Caption = RCuentaTarimas(0)
                        Else
                            LblBatch.Caption = 1
                        End If
                End If
                
                'USUARIO
                If Not IsNull(RBuscaProduccion!Cod_Emp) Then
                    TxtCodEmp.Text = RBuscaProduccion!Cod_Emp
                End If
                
        End If
        
        
    Else
        MsgBox "Esta Linea No Esta Activa", vbOKOnly + vbExclamation, "Informacion"
    End If
End If
End Sub



Private Sub TxtLinCom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtlinIncRec_GotFocus()
        TxtlinIncRec.SelStart = 0
        TxtlinIncRec.SelLength = Len(TxtlinIncRec.Text)
End Sub

Private Sub TxtlinIncRec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub Txttar_GotFocus()
    TxtTar.SelStart = 0
    TxtTar.SelLength = Len(TxtTar.Text)
End Sub

Private Sub Txttar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub



Private Sub Txtbat_GotFocus()
    TxtBat.SelStart = 0
    TxtBat.SelLength = Len(TxtBat.Text)
End Sub



