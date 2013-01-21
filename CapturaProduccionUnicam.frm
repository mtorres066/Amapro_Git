VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CapturaProduccionUnicam 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura de Produccion EXTERNA"
   ClientHeight    =   8595
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11775
   Icon            =   "CapturaProduccionUnicam.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8175
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
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
         Picture         =   "CapturaProduccionUnicam.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "CapturaProduccionUnicam.frx":0BD4
         Height          =   7695
         Left            =   120
         OleObjectBlob   =   "CapturaProduccionUnicam.frx":0BF0
         TabIndex        =   47
         Top             =   240
         Width           =   10455
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   8400
      Top             =   7680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Im&primir"
      Height          =   585
      Left            =   8400
      MouseIcon       =   "CapturaProduccionUnicam.frx":15CB
      Picture         =   "CapturaProduccionUnicam.frx":1A0D
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   585
      Left            =   10080
      MouseIcon       =   "CapturaProduccionUnicam.frx":1F3F
      Picture         =   "CapturaProduccionUnicam.frx":2381
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7080
      Width           =   1485
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   585
      Left            =   6720
      MouseIcon       =   "CapturaProduccionUnicam.frx":43F3
      Picture         =   "CapturaProduccionUnicam.frx":4835
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   585
      Left            =   5040
      MouseIcon       =   "CapturaProduccionUnicam.frx":4D67
      Picture         =   "CapturaProduccionUnicam.frx":51A9
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   585
      Left            =   3360
      MouseIcon       =   "CapturaProduccionUnicam.frx":56DB
      Picture         =   "CapturaProduccionUnicam.frx":5B1D
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   585
      Left            =   1680
      MouseIcon       =   "CapturaProduccionUnicam.frx":604F
      Picture         =   "CapturaProduccionUnicam.frx":6491
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7080
      Width           =   1600
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   585
      Left            =   120
      MouseIcon       =   "CapturaProduccionUnicam.frx":69C3
      Picture         =   "CapturaProduccionUnicam.frx":6E05
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   7080
      Width           =   1485
   End
   Begin VB.Data DataProduccion 
      Caption         =   "Produccion Externa"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Erick\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "ProduccionUnicam"
      Top             =   7800
      Width           =   11505
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   44
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CapturaProduccionUnicam.frx":7337
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameProduccion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General "
      TabPicture(1)   =   "CapturaProduccionUnicam.frx":7651
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridProduccion"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda O Seleccion De Datos"
      TabPicture(2)   =   "CapturaProduccionUnicam.frx":7AA3
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DtpFecFin"
      Tab(2).Control(1)=   "DtpFecIni"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "CmdBuscar"
      Tab(2).Control(4)=   "CmdActualizar"
      Tab(2).Control(5)=   "FrameBuscar"
      Tab(2).Control(6)=   "LblEtiqueta"
      Tab(2).ControlCount=   7
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   375
         Left            =   -66240
         TabIndex        =   40
         Top             =   3480
         Visible         =   0   'False
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
         Format          =   24576003
         CurrentDate     =   37213
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   375
         Left            =   -68400
         TabIndex        =   39
         Top             =   3480
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Format          =   24576003
         CurrentDate     =   37213
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -66240
         TabIndex        =   41
         ToolTipText     =   " "
         Top             =   4400
         Visible         =   0   'False
         Width           =   1845
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   855
         Left            =   -67320
         Picture         =   "CapturaProduccionUnicam.frx":7EF5
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   5000
         Width           =   3015
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Actuali&ar"
         Height          =   825
         Left            =   -67320
         Picture         =   "CapturaProduccionUnicam.frx":8337
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   5960
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
         Height          =   1455
         Left            =   -74760
         TabIndex        =   34
         Top             =   1400
         Width           =   6375
         Begin VB.OptionButton OptFecLin 
            Caption         =   "Fechas Y Linea"
            ForeColor       =   &H80000008&
            Height          =   1000
            Left            =   3240
            Picture         =   "CapturaProduccionUnicam.frx":8641
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Width           =   1400
         End
         Begin VB.OptionButton OptBusFic 
            Caption         =   "Fechas Y Ficha Tecnica"
            ForeColor       =   &H8000000D&
            Height          =   1000
            Left            =   1680
            Picture         =   "CapturaProduccionUnicam.frx":8F0B
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   240
            Width           =   1400
         End
         Begin VB.OptionButton OptBusFec 
            Caption         =   "Fecha"
            ForeColor       =   &H8000000D&
            Height          =   1000
            Left            =   120
            Picture         =   "CapturaProduccionUnicam.frx":9215
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   240
            Value           =   -1  'True
            Width           =   1400
         End
         Begin VB.OptionButton optBusBatch 
            Caption         =   "Batch"
            ForeColor       =   &H80000008&
            Height          =   1000
            Left            =   4800
            Picture         =   "CapturaProduccionUnicam.frx":9ADF
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   1400
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
         Height          =   6135
         Left            =   120
         TabIndex        =   45
         Top             =   680
         Width           =   11535
         Begin VB.TextBox TxtTur 
            Appearance      =   0  'Flat
            DataField       =   "TURNO"
            DataSource      =   "DataProduccion"
            Height          =   285
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   8
            Top             =   5040
            Width           =   615
         End
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataProduccion"
            Height          =   495
            Left            =   2880
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   25
            ToolTipText     =   "100 caracteres maximo"
            Top             =   5400
            Width           =   8415
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NumeroMC9404"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   21
            Left            =   9840
            MaxLength       =   10
            TabIndex        =   20
            Top             =   1920
            Width           =   1575
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NumeroMC9410"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   20
            Left            =   9840
            MaxLength       =   10
            TabIndex        =   19
            Top             =   1560
            Width           =   1575
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "FechaUnicam"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   19
            Left            =   9960
            MaxLength       =   10
            TabIndex        =   24
            Top             =   4800
            Width           =   1300
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "CantidadEnvases"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   18
            Left            =   7920
            TabIndex        =   23
            Top             =   4800
            Width           =   1300
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Paleta"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   17
            Left            =   5640
            TabIndex        =   22
            Top             =   4800
            Width           =   1300
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Orden"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   16
            Left            =   3480
            TabIndex        =   21
            Top             =   4800
            Width           =   1300
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "COD_EMP"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   8
            Left            =   9840
            MaxLength       =   10
            TabIndex        =   16
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Defecto2"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   11
            Left            =   5640
            MaxLength       =   4
            TabIndex        =   12
            Top             =   3120
            Width           =   615
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Cantidad1"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   10
            Left            =   6360
            MaxLength       =   30
            TabIndex        =   11
            Top             =   2760
            Width           =   495
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Defecto1"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   9
            Left            =   5640
            MaxLength       =   4
            TabIndex        =   10
            Top             =   2760
            Width           =   615
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "TARIMA"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   4
            Top             =   3600
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "HOR_PRD"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   3
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   3
            Top             =   3000
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "FEC_PRD"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   2
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   2
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "ESP_TEC"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   1
            Left            =   1080
            MaxLength       =   12
            TabIndex        =   1
            Top             =   840
            Width           =   1575
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NoMP9301"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   15
            Left            =   9840
            MaxLength       =   30
            TabIndex        =   17
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Cantidad3"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   14
            Left            =   6360
            MaxLength       =   30
            TabIndex        =   15
            Top             =   3480
            Width           =   495
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Defecto3"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   13
            Left            =   5640
            MaxLength       =   4
            TabIndex        =   14
            Top             =   3480
            Width           =   615
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Cantidad2"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   12
            Left            =   6360
            MaxLength       =   30
            TabIndex        =   13
            Top             =   3120
            Width           =   495
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "MUESTRA"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   7
            Left            =   1320
            TabIndex        =   7
            Top             =   4680
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "ENVASES"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   6
            Left            =   1320
            TabIndex        =   6
            Top             =   4320
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "BATCH"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   5
            Left            =   1320
            TabIndex        =   5
            Top             =   3960
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "LINEA"
            DataSource      =   "DataProduccion"
            Height          =   285
            Index           =   0
            Left            =   1080
            MaxLength       =   2
            TabIndex        =   0
            Top             =   480
            Width           =   1575
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
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   1200
            Width           =   6975
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
            Height          =   285
            Left            =   1440
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   1560
            Width           =   6975
         End
         Begin VB.ComboBox CboColor 
            Appearance      =   0  'Flat
            DataField       =   "ColorMP9301"
            DataSource      =   "DataProduccion"
            Height          =   315
            ItemData        =   "CapturaProduccionUnicam.frx":9F21
            Left            =   9840
            List            =   "CapturaProduccionUnicam.frx":9F31
            TabIndex        =   18
            Text            =   "BLANCA"
            Top             =   1080
            Width           =   1575
         End
         Begin VB.ComboBox CboCal 
            DataField       =   "CALIDAD"
            DataSource      =   "DataProduccion"
            Enabled         =   0   'False
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
            ItemData        =   "CapturaProduccionUnicam.frx":9F58
            Left            =   1320
            List            =   "CapturaProduccionUnicam.frx":9F68
            TabIndex        =   9
            Text            =   "A"
            Top             =   5520
            Width           =   615
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FF0000&
            BorderWidth     =   3
            X1              =   5520
            X2              =   11400
            Y1              =   4440
            Y2              =   4440
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H000080FF&
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   7
            Left            =   2880
            TabIndex        =   91
            Top             =   5160
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            Caption         =   "Numero MC 404"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   6
            Left            =   8520
            TabIndex        =   90
            Top             =   1920
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Numero MC 410"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   5
            Left            =   8520
            TabIndex        =   89
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   4
            Left            =   9360
            TabIndex        =   88
            Top             =   4800
            Width           =   540
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   7080
            TabIndex        =   87
            Top             =   4800
            Width           =   765
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            Caption         =   "Paleta"
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
            Height          =   195
            Index           =   2
            Left            =   4920
            TabIndex        =   86
            Top             =   4800
            Width           =   555
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H000080FF&
            Caption         =   "Orden"
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
            Index           =   1
            Left            =   2880
            TabIndex        =   85
            Top             =   4800
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            BackColor       =   &H00C0C0C0&
            Caption         =   " Datos De Producto De Exterior"
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
            Index           =   0
            Left            =   2760
            TabIndex        =   84
            Top             =   4320
            Width           =   2895
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   1335
            Left            =   2760
            Top             =   4680
            Width           =   8655
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
            TabIndex        =   83
            Top             =   840
            Width           =   5775
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
            TabIndex        =   82
            Top             =   480
            Width           =   5775
         End
         Begin VB.Shape ShapeCalidad 
            BackStyle       =   1  'Opaque
            Height          =   375
            Index           =   0
            Left            =   2040
            Shape           =   3  'Circle
            Top             =   5520
            Width           =   495
         End
         Begin VB.Label Label6 
            Caption         =   "Defectos"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Index           =   3
            Left            =   4680
            TabIndex        =   80
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Line Line2 
            BorderWidth     =   3
            X1              =   4680
            X2              =   11400
            Y1              =   2280
            Y2              =   2280
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Defecto"
            Height          =   195
            Index           =   2
            Left            =   5040
            TabIndex        =   79
            Top             =   2400
            Width           =   1110
         End
         Begin VB.Label Label6 
            Caption         =   "Cantidad"
            Height          =   255
            Index           =   1
            Left            =   6240
            TabIndex        =   78
            Top             =   2400
            Width           =   735
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
            TabIndex        =   77
            Top             =   480
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
            TabIndex        =   76
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Hora"
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
            Index           =   20
            Left            =   120
            TabIndex        =   75
            Top             =   3000
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
            TabIndex        =   74
            Top             =   720
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
            TabIndex        =   73
            Top             =   3600
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Batch"
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
            Index           =   29
            Left            =   120
            TabIndex        =   72
            Top             =   3960
            Width           =   495
         End
         Begin VB.Label lblLabels 
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
            Index           =   30
            Left            =   120
            TabIndex        =   71
            Top             =   4320
            Width           =   1815
         End
         Begin VB.Label lblLabels 
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
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   70
            Top             =   5520
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Muestra"
            Height          =   255
            Index           =   33
            Left            =   120
            TabIndex        =   69
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "Turno"
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
            TabIndex        =   68
            Top             =   5040
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Tipo"
            Height          =   255
            Left            =   11040
            TabIndex        =   67
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label Label23 
            BackStyle       =   0  'Transparent
            Caption         =   "Hojalata"
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
            Left            =   120
            TabIndex        =   66
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Fondo ó Tapa"
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
            Height          =   195
            Left            =   120
            TabIndex        =   65
            Top             =   1560
            Width           =   1200
         End
         Begin VB.Label Label6 
            Caption         =   "Defecto"
            Height          =   255
            Index           =   0
            Left            =   6960
            TabIndex        =   64
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label Label27 
            Caption         =   "Defecto 1"
            Height          =   255
            Left            =   4680
            TabIndex        =   63
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label Label28 
            Caption         =   "Defecto 2"
            Height          =   255
            Left            =   4680
            TabIndex        =   62
            Top             =   3120
            Width           =   855
         End
         Begin VB.Label Label29 
            Caption         =   "Defecto 3"
            Height          =   255
            Left            =   4680
            TabIndex        =   61
            Top             =   3480
            Width           =   855
         End
         Begin VB.Label Lbl1 
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
            Left            =   6960
            TabIndex        =   60
            Top             =   2760
            Width           =   3975
         End
         Begin VB.Label Lbl3 
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
            Left            =   6960
            TabIndex        =   59
            Top             =   3480
            Width           =   3975
         End
         Begin VB.Label Lbl11 
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
            Left            =   11040
            TabIndex        =   58
            Top             =   2760
            Width           =   375
         End
         Begin VB.Label Lbl22 
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
            Left            =   11040
            TabIndex        =   57
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label Lbl33 
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
            Left            =   11040
            TabIndex        =   56
            Top             =   3480
            Width           =   375
         End
         Begin VB.Label LblDos 
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
            Left            =   6960
            TabIndex        =   55
            Top             =   3120
            Width           =   3975
         End
         Begin VB.Label Label3 
            Caption         =   "No. MP9 301"
            Height          =   255
            Left            =   8760
            TabIndex        =   54
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Color MP9 301"
            Height          =   255
            Left            =   8760
            TabIndex        =   53
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label sdf 
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
            Left            =   8760
            TabIndex        =   52
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label LblBatch 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   720
            TabIndex        =   51
            Top             =   3960
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSDBGrid.DBGrid DBGridProduccion 
         Bindings        =   "CapturaProduccionUnicam.frx":9F78
         Height          =   6135
         Left            =   -74880
         OleObjectBlob   =   "CapturaProduccionUnicam.frx":9F95
         TabIndex        =   33
         Tag             =   "Click En Encabezado De Columna Para Indexar"
         Top             =   680
         Width           =   11535
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
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -69000
         TabIndex        =   81
         Top             =   4400
         Width           =   2535
      End
   End
End
Attribute VB_Name = "CapturaProduccionUnicam"
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

Dim RDefecto1 As Recordset
Dim RDefecto2 As Recordset
Dim RDefecto3 As Recordset
Dim RBuscaProduccion As Recordset
Dim RBuscaEnvases As Recordset
Dim RLineas As Recordset
Dim RBuscaUltimaFicha As Recordset
Dim RReporteIdentificacionInterno As Recordset

'VARIABLES PARA DESPLEGAR DATOS DE FICHA TECNICA
Dim RBuscaFichaTecnica As Recordset
Dim RBuscaPlatina As Recordset
Dim RBuscaForma As Recordset
Dim RBuscaFondo As Recordset
Dim RBuscaTapa As Recordset
Dim RBuscaDefecto1 As Recordset
Dim RBuscaDefecto2 As Recordset
Dim RBuscaDefecto3 As Recordset

Dim RBuscaAtributo As Recordset

Dim VLineas As Boolean
Dim VDefecto As Boolean
Dim VD1 As Boolean
Dim VD2 As Boolean
Dim VD3 As Boolean

Dim VDefecto1 As String
Dim VDefecto2 As String
Dim VDefecto3 As String

Dim VDia As String
Dim VMes As String
Dim VAño As String

Dim VSumaDefectos As Integer
Dim VSumaDefectos2 As Integer

Dim RCuentaTarimas As Recordset

Dim VEditar As Boolean

Dim RVerificaTarima As Recordset
Dim Vtarima As String

Dim RBuscaFoto As Recordset

Dim RBuscaMateriasPrimas As Recordset

Dim RBuscaLinea As Recordset




Sub botones()
    If Bandera = True Then
         FrameProduccion.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdBuscar.Enabled = False
         CmdImprimir.Enabled = False
         LblBatch.Visible = True
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         Txttexto.Item(1).SetFocus
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
         LblBatch.Visible = False
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         FrameBuscar.Visible = True
         DataProduccion.Visible = True
         DBGridProduccion.Visible = True

    End If
End Sub

Private Sub CboCal_Change()
    If CboCal.Text = "A" Then
                ShapeCalidad.Item(0).BackColor = vbGreen
    ElseIf CboCal.Text = "R" Then
                ShapeCalidad.Item(0).BackColor = vbRed
    ElseIf CboCal.Text = "I" Then
                ShapeCalidad.Item(0).BackColor = vbBlue
    ElseIf CboCal.Text = "C" Then
                ShapeCalidad.Item(0).BackColor = vbCyan
    End If
    
End Sub


Private Sub CboCal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub CboCal_LostFocus()
    Txttexto.Item(15).SetFocus
End Sub


Private Sub CboColor_Change()
        If CboColor.Text = "BLANCA" Then
                CboColor.BackColor = vbWhite
        ElseIf CboColor.Text = "ROSADA" Then
                CboColor.BackColor = &HC0C0FF
        ElseIf CboColor.Text = "VERDE" Then
                CboColor.BackColor = vbGreen
        ElseIf CboColor.Text = "ANARANJADA" Then
                CboColor.BackColor = &H80FF&
        End If
End Sub

Private Sub CboColor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub CboColor_Scroll()
        If CboColor.Text = "BLANCA" Then
                CboColor.BackColor = vbWhite
        ElseIf CboColor.Text = "ROSADA" Then
                CboColor.BackColor = &HC0C0FF
        ElseIf CboColor.Text = "VERDE" Then
                CboColor.BackColor = vbGreen
        ElseIf CboColor.Text = "ANARANJADA" Then
                CboColor.BackColor = &H80FF&
        End If
End Sub


Private Sub CmdActualizar_Click()
    DataProduccion.RecordSource = "Select * from ProduccionUnicam"
    DataProduccion.Refresh
    DBGridProduccion.Refresh
    
    SSTab1.Tab = 1

End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next

MousePointer = 11

            If OptBusFic.Value = True Then
                DataProduccion.RecordSource = ("Select * from ProduccionUnicam where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Esp_Tec = '" & TxtBuscar.Text & "'")
            ElseIf OptBusFec.Value = True Then
                    DataProduccion.RecordSource = ("Select * from ProduccionUnicam where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#")
            ElseIf optBusBatch.Value = True Then
                If TxtBuscar.Text <> "" Then
                    If Not IsNumeric(TxtBuscar.Text) Then
                    Else
                        DataProduccion.RecordSource = ("Select * from ProduccionUnicam where batch = " & TxtBuscar.Text)
                    End If
                End If
            ElseIf OptFecLin.Value = True Then
                DataProduccion.RecordSource = ("Select * from ProduccionUnicam where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBuscar.Text & "'")
            End If
            
                    DataProduccion.Refresh
                    DBGridProduccion.Refresh
                
            
            SSTab1.Tab = 1
            
MousePointer = 0
            If Err <> 0 Then
                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                Exit Sub
            End If
        
SSTab1.Tab = 1

End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next

                'BORRA LA IDENTIFICACION INGRESADA A LA BASE DE DATOS
                Db.Execute "Delete * From ReporteIdentificacionInterno"

MousePointer = 11
                VDia = Day(Txttexto.Item(2).Text)
                VMes = Month(Txttexto.Item(2).Text)
                VAño = Year(Txttexto.Item(2).Text)
                
                Set RReporteIdentificacionInterno = Db.OpenRecordset("Select * From ReporteIdentificacionInterno Where Fec_Prd = #" & Format(Txttexto.Item(2).Text, "mm/dd/yyyy") & "# and Linea = '" & Txttexto.Item(0).Text & "' and Tarima = " & Txttexto.Item(4).Text & " and Esp_Tec = '" & Txttexto.Item(1).Text & "'")
                    If RReporteIdentificacionInterno.RecordCount > 0 Then
                            RReporteIdentificacionInterno.Edit
                                    RReporteIdentificacionInterno!Linea = Txttexto.Item(0).Text
                                    RReporteIdentificacionInterno!Esp_tec = Txttexto.Item(1).Text
                                    RReporteIdentificacionInterno!Fec_Prd = Txttexto.Item(2).Text
                                    RReporteIdentificacionInterno!Tarima = Txttexto.Item(4).Text
                                    RReporteIdentificacionInterno!Envases = Txttexto.Item(6).Text
                                    RReporteIdentificacionInterno!Hor_Prd = Txttexto.Item(3).Text
                                    RReporteIdentificacionInterno!Batch = Txttexto.Item(5).Text
                                    RReporteIdentificacionInterno!Cod_emp = Txttexto.Item(8).Text
                                    RReporteIdentificacionInterno!Hojalata = ""
                                    RReporteIdentificacionInterno!Fondo = ""
                                    RReporteIdentificacionInterno!Orden = ""
                            RReporteIdentificacionInterno.Update
                    Else
                            RReporteIdentificacionInterno.AddNew
                                    RReporteIdentificacionInterno!Linea = Txttexto.Item(0).Text
                                    RReporteIdentificacionInterno!Esp_tec = Txttexto.Item(1).Text
                                    RReporteIdentificacionInterno!Fec_Prd = Txttexto.Item(2).Text
                                    RReporteIdentificacionInterno!Tarima = Txttexto.Item(4).Text
                                    RReporteIdentificacionInterno!Envases = Txttexto.Item(6).Text
                                    RReporteIdentificacionInterno!Hor_Prd = Txttexto.Item(3).Text
                                    RReporteIdentificacionInterno!Batch = Txttexto.Item(5).Text
                                    RReporteIdentificacionInterno!Cod_emp = Txttexto.Item(8).Text
                                    RReporteIdentificacionInterno!Hojalata = ""
                                    RReporteIdentificacionInterno!Fondo = ""
                                    RReporteIdentificacionInterno!Orden = ""
                            RReporteIdentificacionInterno.Update
                    End If
                    
                
                CrReportes.SelectionFormula = "{ReporteIdentificacionInterno.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño & "," & VMes & "," & VDia & ") and {ReporteIdentificacionInterno.Linea} = '" & Txttexto.Item(0).Text & "' and {ReporteIdentificacionInterno.Tarima} = " & Txttexto.Item(4).Text & " and {ReporteIdentificacionInterno.Esp_Tec} = '" & Txttexto.Item(1).Text & "'"
                CrReportes.ReportFileName = App.Path & "\Identificacion.rpt"
                
        MousePointer = 0
                CrReportes.Action = 1
            
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
                

End Sub

Private Sub Form_Activate()
On Error Resume Next
    
    DataProduccion.RecordSource = ("Select * from ProduccionUnicam Order BY Fec_prd")
    DataProduccion.Refresh
    DataProduccion.Recordset.MoveLast
    If Err <> 0 Then
        
    End If
End Sub


Private Sub optBusBatch_Click()
    TxtBuscar.Visible = True
    TxtBuscar.SetFocus
    Lbletiqueta.Caption = "Numero De Batch"
    DtpFecIni.Visible = False
    DtpFecFin.Visible = False

End Sub

Private Sub OptBusFec_Click()
    TxtBuscar.Visible = False
    Lbletiqueta.Caption = ""
    DtpFecIni.Visible = True
    DtpFecFin.Visible = True
    DtpFecIni.Value = Date
    DtpFecFin.Value = Date
    DtpFecIni.SetFocus
End Sub

Private Sub OptBusFic_Click()
    DtpFecIni.Visible = True
    DtpFecFin.Visible = True
    DtpFecIni.Value = Date
    DtpFecFin.Value = Date
    TxtBuscar.Visible = True
    TxtBuscar.SetFocus
    Lbletiqueta.Caption = "Ficha Tecnica"
End Sub
Private Sub CmdAgregar_Click()
On Error Resume Next
        Bandera = True
        botones
        DataProduccion.Recordset.AddNew
        
        If Err <> 0 Then
                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                Exit Sub
        End If
        
        Txttexto.Item(0).SetFocus
        
        'SI LA HORA ES MENOR QUE LAS 7 DE LA MAÑANA ENTONCES DA LA FECHA ANTERIOR
        'If Format(Time, "hh:mm") < "07:00" Then
        '    TxtTexto.Item(2).Text = Format(DateValue(Date) - 1, "dd/mm/yyyy")
        'Else
            Txttexto.Item(2).Text = Format(Date, "dd/mm/yyyy")
        'End If
                
        Txttexto.Item(3).Text = Format(Time, "hh:mm")
        'TxtCodEmp.Text = GUsuario
        
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

        Bandera = False
        botones
        DataProduccion.Recordset.CancelUpdate
        
        If Err <> 0 Then
                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                Exit Sub
        End If

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
        Bandera = True
        botones
        DataProduccion.Recordset.Edit
        
        If Err <> 0 Then
             MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
             Exit Sub
        End If
        Txttexto.Item(1).SetFocus
        VEditar = True
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
    MousePointer = 11
  
  'SUMA LA CANTIDAD DE DEFECTOS
   VSumaDefectos2 = Val(Txttexto.Item(9).Text) + Val(Txttexto.Item(11).Text) + Val(Txttexto.Item(13).Text)
   If VSumaDefectos2 = 0 Then
       CboCal.Text = "A"
   End If
  
    'REVISA LOS ENVASES ASIGNADOS A LA FICHA TECNICA PARA COMPARAR QUE TIPO DE TARIMA ES Y SI ES DE EXTERNA O INTERNA
    Set RBuscaEnvases = Db.OpenRecordset("Select Envases, Origen From FichaTecnica Where Esp_Tec = '" & Txttexto.Item(1).Text & "'")
    If RBuscaEnvases.RecordCount > 0 Then
        'REVISA SI ESTA FICHA TECNICA PERTENECE A EXTERNA
        If RBuscaEnvases!Origen = "INTERNO" Then
            MsgBox "La Ficha Tecnica Que Esta Asignada A Esta Linea No Pertenece A Una Empresa EXTERNO", vbOKOnly + vbExclamation, "ADVERTENECIA"
            Exit Sub
        End If
    
        'REVISA SI LOS ENVASES NO SON LOS ASIGNADOS A LA FICHA PREGUNTA SI ES INCOMPLETA O COMPLEMENTO
        If Val(Txttexto.Item(6).Text) < Val(RBuscaEnvases(0)) Then
            mensaje = MsgBox("Esta Tarima Es Incompleta", vbYesNo, "Informacion")
                If mensaje = vbYes Then
                    CboCal.Text = "I"
                Else
                    mensaje = MsgBox("Esta Tarima Es Complemento", vbYesNo, "Informacion")
                        If mensaje = vbYes Then
                            CboCal.Text = "C"
                        End If
                End If
        End If
    End If
   
   'VERIFICA EL COMBO DE CALIDAD
   If CboCal.Text <> "A" And CboCal.Text <> "C" And CboCal.Text <> "I" And CboCal.Text <> "R" Then
        MsgBox "CALIDAD INCORRECTA", vbOKOnly + vbInformation, "Informacion"
        MousePointer = 0
        Exit Sub
   End If
   
   'REVISA EL NUMERO DE BATCH
   If Not IsNumeric(Txttexto.Item(5).Text) Then
        MsgBox "Numero de Batch Incorrecto", vbOKOnly + vbInformation, "Informacion"
        MousePointer = 0
        Exit Sub
   End If
   
   Vlinea = Txttexto.Item(0).Text
   VFicha = Txttexto.Item(1).Text
   VFecha = Txttexto.Item(2).Text
   Vtarima = Txttexto.Item(4).Text
       
   'VERIFICA SI YA EXISTE LA TARIMA Y PREGUNTA SI LA GRABA
   Set RVerificaTarima = Db.OpenRecordset("Select * from ProduccionUnicam Where Linea = '" & Vlinea & "' and Esp_tec = '" & VFicha & "' and Fec_prd = #" & Format(VFecha, "dd/mm/yyyy") & "# and Tarima = " & Vtarima)
   If RVerificaTarima.RecordCount > 0 Then
            mensaje = MsgBox("Ya Existe Tarima " & Vtarima & " De Ficha " & VFicha & " Con Fecha " & VFecha & " DESEA GRABARLA ", vbYesNo + vbInformation, "Informacion")
                If mensaje = vbYes Then
                Else
                    MousePointer = 0
                    Exit Sub
                End If
   End If
     
   'GRABA LOS DATOS
   DataProduccion.Recordset.Update
   
   If Err <> 0 Then
      MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
      Txttexto.Item(0).SetFocus
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
      
      
      'DataProduccion.RecordSource = ("Select * from ProduccionUnicam Order BY Fec_prd, Hor_Prd")
      'DataProduccion.Refresh
      DataProduccion.Recordset.MoveLast
      
      CmdAgregar.SetFocus
  End If
            
       MousePointer = 0
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_DblClick()
    'PARA SELECCIONAR LA LINEA
    If VLineas = True Then
        Txttexto.Item(0).Text = DBGridConsultas.Columns(0)
        Txttexto.Item(0).SetFocus
        FrameConsultas.Visible = False
    End If
    'PARA SELECCIONAR EL DEFECTO 1
    If VD1 = True Then
        Txttexto.Item(9).Text = DBGridConsultas.Columns(0)
        Txttexto.Item(0).SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR EL DEFECTO 2
    If VD2 = True Then
        Txttexto.Item(11).Text = DBGridConsultas.Columns(0)
        Txttexto.Item(11).SetFocus
        FrameConsultas.Visible = False
    End If
    
    'PARA SELECCIONAR EL DEFECTO 3
    If VD3 = True Then
        Txttexto.Item(13).Text = DBGridConsultas.Columns(0)
        Txttexto.Item(13).SetFocus
        FrameConsultas.Visible = False
    End If
        
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
    'PARA SELECCIONAR LA LINEA
        If VLineas = True Then
            Txttexto.Item(0).Text = DBGridConsultas.Columns(0)
            Txttexto.Item(0).SetFocus
            FrameConsultas.Visible = False
        End If
        'PARA SELECCIONAR EL DEFECTO 1
        If VD1 = True Then
            Txttexto.Item(9).Text = DBGridConsultas.Columns(0)
            Txttexto.Item(0).SetFocus
            FrameConsultas.Visible = False
        End If
        
        'PARA SELECCIONAR EL DEFECTO 2
        If VD2 = True Then
            Txttexto.Item(11).Text = DBGridConsultas.Columns(0)
            Txttexto.Item(11).SetFocus
            FrameConsultas.Visible = False
        End If
        
        'PARA SELECCIONAR EL DEFECTO 3
        If VD3 = True Then
            Txttexto.Item(13).Text = DBGridConsultas.Columns(0)
            Txttexto.Item(13).SetFocus
            FrameConsultas.Visible = False
        End If
    End If
    
End Sub

Private Sub DbgridProduccion_HeadClick(ByVal ColIndex As Integer)
    DataProduccion.RecordSource = ("Select * from ProduccionUnicam order by " & DBGridProduccion.Columns(ColIndex).DataField)
    DataProduccion.Refresh
    DBGridProduccion.Refresh
    
End Sub

Private Sub Form_Load()
    DataProduccion.Connect = GConnect
    DataConsultas.Connect = GConnect
    DataProduccion.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
    
    If GEditar = True Then
        DBGridProduccion.AllowUpdate = True
    Else
        DBGridProduccion.AllowUpdate = False
    End If
    
End Sub






Private Sub OptFecLin_Click()
    DtpFecIni.Visible = True
    DtpFecFin.Visible = True
    DtpFecIni.Value = Date
    DtpFecFin.Value = Date
    TxtBuscar.Visible = True
    TxtBuscar.SetFocus
    Lbletiqueta.Caption = "Linea"

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

Private Sub TxtObs_GotFocus()
        TxtObs.SelStart = 0
        TxtObs.SelLength = Len(TxtObs.Text)
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        'LINEA
        If Index = 0 Then
            Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txttexto.Item(0).Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                LblLinea.Caption = RBuscaLinea!Descrip
            Else
                LblLinea.Caption = ""
            End If
        'FICHA TECNICA
        ElseIf Index = 1 Then
            Set RBuscaFichaTecnica = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & Txttexto.Item(1).Text & "'")
                If RBuscaFichaTecnica.RecordCount > 0 Then
                        'PLATINAS
                        'Set RBuscaPlatina = Db.OpenRecordset("Select Descrip From Platinas Where Platina = '" & RBuscaFichaTecnica(0) & "'")
                        'If RBuscaPlatina.RecordCount > 0 Then
                        '    Text1.Text = RBuscaPlatina(0)
                       ' Else
                        '    Text1.Text = ""
                       ' End If
                        
                        'FONDOS
                        'Set RBuscaFondo = Db.OpenRecordset("Select Descrip From Fondos Where Fondo = '" & RBuscaFichaTecnica(2) & "'")
                        'If RBuscaFondo.RecordCount > 0 Then
                        '    Text4.Text = RBuscaFondo(0)
                        'Else
                            Text4.Text = ""
                        'End If
                        
                        LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
                Else
                        LblFichaTecnica.Caption = ""
                End If
        'CODIGO DEFECTO 1
        ElseIf Index = 9 Then
            Set RDefecto1 = Db.OpenRecordset("Select Descrip, Tipo From Defectos Where Defecto = '" & Txttexto.Item(9).Text & "'")
                If RDefecto1.RecordCount > 0 Then
                    Lbl1.Caption = RDefecto1(0)
                    Lbl11.Caption = RDefecto1(1)
                Else
                    Lbl1.Caption = ""
                    Lbl11.Caption = ""
                End If
        'CODIGO DEFECTO 2
        ElseIf Index = 11 Then
            Set RDefecto2 = Db.OpenRecordset("Select Descrip, Tipo From Defectos Where Defecto = '" & Txttexto.Item(11).Text & "'")
                If RDefecto2.RecordCount > 0 Then
                    LblDos.Caption = RDefecto2(0)
                    Lbl22.Caption = RDefecto2(1)
                Else
                    LblDos.Caption = ""
                    Lbl22.Caption = ""
                End If

        'CODIGO DEFECTO 3
        ElseIf Index = 13 Then
            Set RDefecto3 = Db.OpenRecordset("Select Descrip, Tipo From Defectos Where Defecto = '" & Txttexto.Item(13).Text & "'")
                If RDefecto3.RecordCount > 0 Then
                    Lbl3.Caption = RDefecto3(0)
                    Lbl33.Caption = RDefecto3(1)
                Else
                    Lbl3.Caption = ""
                    Lbl33.Caption = ""
                End If
        End If
            
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        'LINEAS
        If Index = 0 Then
            VLineas = True
            VD1 = False
            VD2 = False
            VD3 = False
            DataConsultas.RecordSource = ("Select * from Lineas")
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            DBGridConsultas.SetFocus
            Txttexto.Item(0).Text = ""
            
        'CODIGO DE DEFECTO 1
        ElseIf Index = 9 Then
            VLineas = False
            VD1 = True
            VD2 = False
            VD3 = False
            DataConsultas.RecordSource = ("Select * from Defectos")
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            DBGridConsultas.SetFocus
        
        'CODIGO DE DEFECTO 2
        ElseIf Index = 11 Then
            VLineas = False
            VD1 = False
            VD2 = True
            VD3 = False
            
            DataConsultas.RecordSource = ("Select * from Defectos")
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            DBGridConsultas.SetFocus
        
        'CODIGO DE DEFECTO 3
        ElseIf Index = 13 Then
            VLineas = False
            VD1 = False
            VD2 = False
            VD3 = True
            
            DataConsultas.RecordSource = ("Select * from Defectos")
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            DBGridConsultas.SetFocus
        End If
        
            DBGridConsultas.Columns(0).Width = "1200"
            DBGridConsultas.Columns(1).Width = "3500"

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        Txttexto.Item(Index).SelStart = 0
        Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index))
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
        
        'LINEA
        If Index = 0 Then
            If KeyAscii = 43 Then
                VLineas = True
                VD1 = False
                VD2 = False
                VD3 = False
                DataConsultas.RecordSource = ("Select * from Lineas")
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                FrameConsultas.Visible = True
                DBGridConsultas.SetFocus
                Txttexto.Item(0).Text = ""
            End If
        
        'CODIGO DE DEFECTO 1
        ElseIf Index = 9 Then
                If KeyAscii = 43 Then
                    VLineas = False
                    VD1 = True
                    VD2 = False
                    VD3 = False
                    DataConsultas.RecordSource = ("Select * from Defectos")
                    DataConsultas.Refresh
                    DBGridConsultas.Refresh
                    FrameConsultas.Visible = True
                    DBGridConsultas.SetFocus
                End If
        'CODIGO DE DEFECTO 2
        ElseIf Index = 11 Then
                If KeyAscii = 43 Then
                    VLineas = False
                    VD1 = False
                    VD2 = True
                    VD3 = False
                    DataConsultas.RecordSource = ("Select * from Defectos")
                    DataConsultas.Refresh
                    DBGridConsultas.Refresh
                    FrameConsultas.Visible = True
                    DBGridConsultas.SetFocus
                End If

        'CODIGO DE DEFECTO 3
        ElseIf Index = 13 Then
                If KeyAscii = 43 Then
                    VLineas = False
                    VD1 = False
                    VD2 = False
                    VD3 = True
                    DataConsultas.RecordSource = ("Select * from Defectos")
                    DataConsultas.Refresh
                    DBGridConsultas.Refresh
                    FrameConsultas.Visible = True
                    DBGridConsultas.SetFocus
                End If
        End If
        
        DBGridConsultas.Columns(0).Width = "1200"
        DBGridConsultas.Columns(1).Width = "3500"

End Sub

Private Sub TxtTexto_LostFocus(Index As Integer)
        'LINEA
        If Index = 0 Then
                If Txttexto.Item(0).Text = "+" Then
                
                ElseIf Txttexto.Item(0).Text = "" Then
                
                Else
                                'VERIFICA SI LA LINEA ESTA ACTIVA
                                Set RLineas = Db.OpenRecordset("Select Esp_Tec, Tarima from Lineas Where Linea = '" & Txttexto.Item(0).Text & "' and Activa = -1")
                                
                                If RLineas.RecordCount > 0 Then
                                                        'FICHA TECNICA
                                                        Txttexto.Item(1).Text = RLineas!Esp_tec
                                                        'TARIMA
                                                        Txttexto.Item(4).Text = Val(RLineas!Tarima) + 1
                                                        
                                                        'BUSCA LA FICHA TECNICA Y JALA LOS CODIGOS DE ALAMBRE BARNIZES SELLO Y NYLON
                                                        Set RBuscaFichaTecnica = Db.OpenRecordset("Select Envases, Origen From FichaTecnica Where Esp_Tec = '" & RLineas!Esp_tec & "'")
                                                        
                                                        If RBuscaFichaTecnica.RecordCount > 0 Then
                                                            'REVISA SI ESTA FICHA TECNICA PERTENECE A EXTERNA
                                                            If RBuscaFichaTecnica!Origen = "INTERNA" Then
                                                                MsgBox "La Ficha Tecnica Que Esta Asignada A Esta Linea No Pertenece A Empresa Externa", vbOKOnly + vbExclamation, "ADVERTENECIA"
                                                                Exit Sub
                                                            End If
                                                            'ENVASES
                                                            Txttexto.Item(6).Text = RBuscaFichaTecnica!Envases
                                                        Else
                                                            Txttexto.Item(6).Text = "0"
                                                            Txttexto.Item(7).Text = "0"
                                                        End If
                                                        
                                                        'SELECCIONA TODOS LOS REGISTROS DE ProduccionUnicam Y SE VA AL ULTIMO
                                                        Set RBuscaProduccion = Db.OpenRecordset("Select * From ProduccionUnicam Where Linea = '" & Txttexto.Item(0).Text & "' and Esp_Tec = '" & RLineas!Esp_tec & "' and Fec_Prd = #" & Format(Txttexto.Item(2).Text, "mm/dd/yyyy") & "# Order By Tarima")
                                                        
                                                        If RBuscaProduccion.RecordCount > 0 Then
                                                                'SE MUEVE AL ULTIMO REGISTRO
                                                                RBuscaProduccion.MoveLast
                                                
                                                                'BATCH
                                                                If Not IsNull(RBuscaProduccion!Batch) Then
                                                                    Txttexto.Item(5).Text = RBuscaProduccion!Batch
                                                                End If
                                                                    
                                                                'TURNO
                                                                If Not IsNull(RBuscaProduccion!Turno) Then
                                                                    TxtTur.Text = RBuscaProduccion!Turno
                                                                End If
                                                                    
                                                                'CUENTA CUANTAS TARIMAS LLEVA EL BATCH
                                                                Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From ProduccionUnicam Where Batch = " & Txttexto.Item(5).Text)
                                                                If RCuentaTarimas.RecordCount > 0 Then
                                                                            LblBatch.Caption = RCuentaTarimas(0)
                                                                Else
                                                                            LblBatch.Caption = 1
                                                                End If
                                                                
                                                                'MUESTRA
                                                                If Not IsNull(RBuscaProduccion!Muestra) Then
                                                                    Txttexto.Item(7).Text = RBuscaProduccion!Muestra
                                                                End If
                                                                'USUARIO
                                                                If Not IsNull(RBuscaProduccion!Cod_emp) Then
                                                                    Txttexto.Item(8).Text = RBuscaProduccion!Cod_emp
                                                                End If
                                                                                                                                
                                                        End If
                                                                
                            Else
                                    MsgBox "Esta Linea No Esta Activa", vbOKOnly + vbExclamation, "Informacion"
                                    
                            End If
                End If
                
        'FECHA
        ElseIf Index = 2 Then
            Txttexto.Item(2).Text = Format(Txttexto.Item(2).Text, "dd/mm/yyyy")
            
                                                        'SELECCIONA TODOS LOS REGISTROS DE ProduccionUnicam Y SE VA AL ULTIMO
                                                        Set RBuscaProduccion = Db.OpenRecordset("Select * From ProduccionUnicam Where Linea = '" & Txttexto.Item(0).Text & "' and Esp_Tec = '" & RLineas!Esp_tec & "' and Fec_Prd = #" & Format(Txttexto.Item(2).Text, "mm/dd/yyyy") & "# Order By Tarima")
                                                        
                                                        If RBuscaProduccion.RecordCount > 0 Then
                                                                'SE MUEVE AL ULTIMO REGISTRO
                                                                RBuscaProduccion.MoveLast
                                                
                                                                'BATCH
                                                                If Not IsNull(RBuscaProduccion!Batch) Then
                                                                    Txttexto.Item(5).Text = RBuscaProduccion!Batch
                                                                End If
                                                                    
                                                                'TURNO
                                                                If Not IsNull(RBuscaProduccion!Turno) Then
                                                                    TxtTur.Text = RBuscaProduccion!Turno
                                                                End If
                                                                    
                                                                'CUENTA CUANTAS TARIMAS LLEVA EL BATCH
                                                                Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From ProduccionUnicam Where Batch = " & Txttexto.Item(5).Text)
                                                                If RCuentaTarimas.RecordCount > 0 Then
                                                                            LblBatch.Caption = RCuentaTarimas(0)
                                                                Else
                                                                            LblBatch.Caption = 1
                                                                End If
                                                                
                                                                'MUESTRA
                                                                If Not IsNull(RBuscaProduccion!Muestra) Then
                                                                    Txttexto.Item(7).Text = RBuscaProduccion!Muestra
                                                                End If
                                                                'USUARIO
                                                                If Not IsNull(RBuscaProduccion!Cod_emp) Then
                                                                    Txttexto.Item(8).Text = RBuscaProduccion!Cod_emp
                                                                End If
                                                                                                                                
                                                        End If
        
        'CODIGO DE DEFECTO 1
        ElseIf Index = 9 Then
                'BUSCAMOS EL TIPO DE DEFECTO QUE ES SI ES MENOR MAYOR O CRITICO
                Set RBuscaDefecto1 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(9).Text & "'")
                If RBuscaDefecto1.RecordCount > 0 Then
                    VDefecto1 = RBuscaDefecto1(0)
                Else
                    VDefecto1 = ""
                End If
                        
                Set RBuscaDefecto2 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(11).Text & "'")
                If RBuscaDefecto2.RecordCount > 0 Then
                    VDefecto2 = RBuscaDefecto2(0)
                Else
                    VDefecto2 = ""
                End If
                
                Set RBuscaDefecto3 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(2).Text & "'")
                If RBuscaDefecto3.RecordCount > 0 Then
                    VDefecto3 = RBuscaDefecto3(0)
                Else
                    VDefecto3 = ""
                End If
                
                    'COMPARA SI SON DEL MISMO TIPO
                If (VDefecto1 = VDefecto2) And (VDefecto1 = VDefecto3) And (VDefecto2 = VDefecto3) Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto1 = VDefecto2 And VDefecto1 <> "" And VDefecto2 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text)
                ElseIf VDefecto1 = VDefecto3 And VDefecto1 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto2 = VDefecto3 And VDefecto2 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                Else
                    VSumaDefectos = Txttexto.Item(10).Text
                End If
            
                        
                If RBuscaDefecto1.RecordCount > 0 Then
                    'BUSCAMOS EN LA FICHA TECNICA EL CODIGO DE ATRIBUTO QUE TIENE
                    Set RBuscaFichaTecnica = Db.OpenRecordset("Select A.Menores, A.Mayores, A.Criticos From FichaTecnica as F, Atributos As A Where F.Esp_Tec = '" & Txttexto.Item(1).Text & "' And F.Atributos = A.Codigo")
                        If RBuscaFichaTecnica.RecordCount > 0 Then
                            
                            If RBuscaDefecto1(0) = 0 Then
                                If VSumaDefectos > RBuscaFichaTecnica(0) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto1(0) = 1 Then
                                If VSumaDefectos > RBuscaFichaTecnica(1) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto1(0) = 2 Then
                                If VSumaDefectos > RBuscaFichaTecnica(2) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            
                            End If
                            
                            
                    End If
                End If
                            'SUMA LA CANTIDAD DE DEFECTOS
                            VSumaDefectos2 = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                            If VSumaDefectos2 = 0 Then
                                CboCal.Text = "A"
                            End If
                            
        'CANTIDAD DE DEFECTO 1
        ElseIf Index = 10 Then
                'BUSCAMOS EL TIPO DE DEFECTO QUE ES SI ES MENOR MAYOR O CRITICO
                Set RBuscaDefecto1 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(9).Text & "'")
                If RBuscaDefecto1.RecordCount > 0 Then
                    VDefecto1 = RBuscaDefecto1(0)
                Else
                    VDefecto1 = ""
                End If
                        
                Set RBuscaDefecto2 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(11).Text & "'")
                If RBuscaDefecto2.RecordCount > 0 Then
                    VDefecto2 = RBuscaDefecto2(0)
                Else
                    VDefecto2 = ""
                End If
                
                Set RBuscaDefecto3 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(13).Text & "'")
                If RBuscaDefecto3.RecordCount > 0 Then
                    VDefecto3 = RBuscaDefecto3(0)
                Else
                    VDefecto3 = ""
                End If
                
                    'COMPARA SI SON DEL MISMO TIPO
                If (VDefecto1 = VDefecto2) And (VDefecto1 = VDefecto3) And (VDefecto2 = VDefecto3) Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto1 = VDefecto2 And VDefecto1 <> "" And VDefecto2 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text)
                ElseIf VDefecto1 = VDefecto3 And VDefecto1 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto2 = VDefecto3 And VDefecto2 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                Else
                    VSumaDefectos = Txttexto.Item(10).Text
                End If
            
                        
                If RBuscaDefecto1.RecordCount > 0 Then
                    'BUSCAMOS EN LA FICHA TECNICA EL CODIGO DE ATRIBUTO QUE TIENE
                    Set RBuscaFichaTecnica = Db.OpenRecordset("Select A.Menores, A.Mayores, A.Criticos From FichaTecnica as F, Atributos As A Where F.Esp_Tec = '" & Txttexto.Item(1).Text & "' And F.Atributos = A.Codigo")
                        If RBuscaFichaTecnica.RecordCount > 0 Then
                            
                            If RBuscaDefecto1(0) = 0 Then
                                If VSumaDefectos > RBuscaFichaTecnica(0) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto1(0) = 1 Then
                                If VSumaDefectos > RBuscaFichaTecnica(1) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto1(0) = 2 Then
                                If VSumaDefectos > RBuscaFichaTecnica(2) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            
                            End If
                    End If
                End If
                
                            'SUMA LA CANTIDAD DE DEFECTOS
                            VSumaDefectos2 = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                            If VSumaDefectos2 = 0 Then
                                CboCal.Text = "A"
                            End If

        'CODIGO DE DEFECTO 2
        ElseIf Index = 11 Then
        
                'BUSCAMOS EL TIPO DE DEFECTO QUE ES SI ES MENOR MAYOR O CRITICO
                Set RBuscaDefecto1 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(9).Text & "'")
                If RBuscaDefecto1.RecordCount > 0 Then
                    VDefecto1 = RBuscaDefecto1(0)
                Else
                    VDefecto1 = ""
                End If
                        
                Set RBuscaDefecto2 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(11).Text & "'")
                If RBuscaDefecto2.RecordCount > 0 Then
                    VDefecto2 = RBuscaDefecto2(0)
                Else
                    VDefecto2 = ""
                End If
                
                Set RBuscaDefecto3 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(13).Text & "'")
                If RBuscaDefecto3.RecordCount > 0 Then
                    VDefecto3 = RBuscaDefecto3(0)
                Else
                    VDefecto3 = ""
                End If
                
                    'COMPARA SI SON DEL MISMO TIPO
                If (VDefecto1 = VDefecto2) And (VDefecto1 = VDefecto3) And (VDefecto2 = VDefecto3) Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto1 = VDefecto2 And VDefecto1 <> "" And VDefecto2 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text)
                ElseIf VDefecto1 = VDefecto3 And VDefecto1 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto2 = VDefecto3 And VDefecto2 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                Else
                    VSumaDefectos = Txttexto.Item(12).Text
                End If
            
                        
                If RBuscaDefecto2.RecordCount > 0 Then
                    'BUSCAMOS EN LA FICHA TECNICA EL CODIGO DE ATRIBUTO QUE TIENE
                    Set RBuscaFichaTecnica = Db.OpenRecordset("Select A.Menores, A.Mayores, A.Criticos From FichaTecnica as F, Atributos As A Where F.Esp_Tec = '" & Txttexto.Item(1).Text & "' And F.Atributos = A.Codigo")
                        If RBuscaFichaTecnica.RecordCount > 0 Then
                            
                            If RBuscaDefecto2(0) = 0 Then
                                If VSumaDefectos > RBuscaFichaTecnica(0) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto2(0) = 1 Then
                                If VSumaDefectos > RBuscaFichaTecnica(1) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto2(0) = 2 Then
                                If VSumaDefectos > RBuscaFichaTecnica(2) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            
                            End If
                            
                        
                    End If
                End If
                            'SUMA LA CANTIDAD DE DEFECTOS
                            VSumaDefectos2 = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                            If VSumaDefectos2 = 0 Then
                                CboCal.Text = "A"
                            End If
        
        'CANTIDAD DE DEFECTO 2
        ElseIf Index = 12 Then
                'BUSCAMOS EL TIPO DE DEFECTO QUE ES SI ES MENOR MAYOR O CRITICO
                Set RBuscaDefecto1 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(9).Text & "'")
                If RBuscaDefecto1.RecordCount > 0 Then
                    VDefecto1 = RBuscaDefecto1(0)
                Else
                    VDefecto1 = ""
                End If
                        
                Set RBuscaDefecto2 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(11).Text & "'")
                If RBuscaDefecto2.RecordCount > 0 Then
                    VDefecto2 = RBuscaDefecto2(0)
                Else
                    VDefecto2 = ""
                End If
                
                Set RBuscaDefecto3 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(13).Text & "'")
                If RBuscaDefecto3.RecordCount > 0 Then
                    VDefecto3 = RBuscaDefecto3(0)
                Else
                    VDefecto3 = ""
                End If
                        
                'COMPARA SI SON DEL MISMO TIPO
                If (VDefecto1 = VDefecto2) And (VDefecto1 = VDefecto3) And (VDefecto2 = VDefecto3) Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto1 = VDefecto2 And VDefecto1 <> "" And VDefecto2 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text)
                ElseIf VDefecto1 = VDefecto3 And VDefecto1 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto2 = VDefecto3 And VDefecto2 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                Else
                    VSumaDefectos = Txttexto.Item(12).Text
                End If
                        
                If RBuscaDefecto2.RecordCount > 0 Then
                    'BUSCAMOS EN LA FICHA TECNICA EL CODIGO DE ATRIBUTO QUE TIENE
                    Set RBuscaFichaTecnica = Db.OpenRecordset("Select A.Menores, A.Mayores, A.Criticos From FichaTecnica as F, Atributos As A Where F.Esp_Tec = '" & Txttexto.Item(1).Text & "' And F.Atributos = A.Codigo")
                        If RBuscaFichaTecnica.RecordCount > 0 Then
                            
                            If RBuscaDefecto2(0) = 0 Then
                                If VSumaDefectos > RBuscaFichaTecnica(0) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto2(0) = 1 Then
                                If VSumaDefectos > RBuscaFichaTecnica(1) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto2(0) = 2 Then
                                If VSumaDefectos > RBuscaFichaTecnica(2) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            
                            End If
                            
                    End If
                End If
                            'SUMA LA CANTIDAD DE DEFECTOS
                            VSumaDefectos2 = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                            If VSumaDefectos2 = 0 Then
                                CboCal.Text = "A"
                            End If
            

        'CODIGO DE DEFECTO 3
        ElseIf Index = 13 Then
                'BUSCAMOS EL TIPO DE DEFECTO QUE ES SI ES MENOR MAYOR O CRITICO
                Set RBuscaDefecto1 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(9).Text & "'")
                If RBuscaDefecto1.RecordCount > 0 Then
                    VDefecto1 = RBuscaDefecto1(0)
                Else
                    VDefecto1 = ""
                End If
                        
                Set RBuscaDefecto2 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(11).Text & "'")
                If RBuscaDefecto2.RecordCount > 0 Then
                    VDefecto2 = RBuscaDefecto2(0)
                Else
                    VDefecto2 = ""
                End If
                
                Set RBuscaDefecto3 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(13).Text & "'")
                If RBuscaDefecto3.RecordCount > 0 Then
                    VDefecto3 = RBuscaDefecto3(0)
                Else
                    VDefecto3 = ""
                End If
                
                
                    'COMPARA SI SON DEL MISMO TIPO
                If (VDefecto1 = VDefecto2) And (VDefecto1 = VDefecto3) And (VDefecto2 = VDefecto3) Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto1 = VDefecto2 And VDefecto1 <> "" And VDefecto2 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text)
                ElseIf VDefecto1 = VDefecto3 And VDefecto1 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto2 = VDefecto3 And VDefecto2 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                Else
                    VSumaDefectos = Txttexto.Item(14).Text
                End If
                        
                If RBuscaDefecto3.RecordCount > 0 Then
                    'BUSCAMOS EN LA FICHA TECNICA EL CODIGO DE ATRIBUTO QUE TIENE
                    Set RBuscaFichaTecnica = Db.OpenRecordset("Select A.Menores, A.Mayores, A.Criticos From FichaTecnica as F, Atributos As A Where F.Esp_Tec = '" & Txttexto.Item(1).Text & "' And F.Atributos = A.Codigo")
                        If RBuscaFichaTecnica.RecordCount > 0 Then
                            
                            If RBuscaDefecto3(0) = 0 Then
                                If VSumaDefectos > RBuscaFichaTecnica(0) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto3(0) = 1 Then
                                If VSumaDefectos > RBuscaFichaTecnica(1) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto3(0) = 2 Then
                                If VSumaDefectos > RBuscaFichaTecnica(2) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            End If
                    End If
                End If
                            'SUMA LA CANTIDAD DE DEFECTOS
                            VSumaDefectos2 = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                            If VSumaDefectos2 = 0 Then
                                CboCal.Text = "A"
                            End If
        
        'CANTIDAD DE DEFECTO 3
        ElseIf Index = 14 Then
                'BUSCAMOS EL TIPO DE DEFECTO QUE ES SI ES MENOR MAYOR O CRITICO
                Set RBuscaDefecto1 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(9).Text & "'")
                If RBuscaDefecto1.RecordCount > 0 Then
                    VDefecto1 = RBuscaDefecto1(0)
                Else
                    VDefecto1 = ""
                End If
                        
                Set RBuscaDefecto2 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(11).Text & "'")
                If RBuscaDefecto2.RecordCount > 0 Then
                    VDefecto2 = RBuscaDefecto2(0)
                Else
                    VDefecto2 = ""
                End If
                
                Set RBuscaDefecto3 = Db.OpenRecordset("select Tipo from Defectos Where Defecto = '" & Txttexto.Item(13).Text & "'")
                If RBuscaDefecto3.RecordCount > 0 Then
                    VDefecto3 = RBuscaDefecto3(0)
                Else
                    VDefecto3 = ""
                End If
                        
                        
                'COMPARA SI SON DEL MISMO TIPO
                If (VDefecto1 = VDefecto2) And (VDefecto1 = VDefecto3) And (VDefecto2 = VDefecto3) Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto1 = VDefecto2 And VDefecto1 <> "" And VDefecto2 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text)
                ElseIf VDefecto1 = VDefecto3 And VDefecto1 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(14).Text)
                ElseIf VDefecto2 = VDefecto3 And VDefecto2 <> "" And VDefecto3 <> "" Then
                    VSumaDefectos = Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                Else
                    VSumaDefectos = Txttexto.Item(14).Text
                End If
                        
                If RBuscaDefecto3.RecordCount > 0 Then
                    'BUSCAMOS EN LA FICHA TECNICA EL CODIGO DE ATRIBUTO QUE TIENE
                    Set RBuscaFichaTecnica = Db.OpenRecordset("Select A.Menores, A.Mayores, A.Criticos From FichaTecnica as F, Atributos As A Where F.Esp_Tec = '" & Txttexto.Item(1).Text & "' And F.Atributos = A.Codigo")
                        If RBuscaFichaTecnica.RecordCount > 0 Then
                            
                            If RBuscaDefecto3(0) = 0 Then
                                If VSumaDefectos > RBuscaFichaTecnica(0) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto3(0) = 1 Then
                                If VSumaDefectos > RBuscaFichaTecnica(1) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            ElseIf RBuscaDefecto3(0) = 2 Then
                                If VSumaDefectos > RBuscaFichaTecnica(2) Then
                                    CboCal.Text = "R"
                                Else
                                    CboCal.Text = "A"
                                End If
                            End If
                            
                    End If
                End If
                            'SUMA LA CANTIDAD DE DEFECTOS
                            VSumaDefectos2 = Val(Txttexto.Item(10).Text) + Val(Txttexto.Item(12).Text) + Val(Txttexto.Item(14).Text)
                            If VSumaDefectos2 = 0 Then
                                CboCal.Text = "A"
                            End If
                    
        'FECHA DE EXTERNA
        ElseIf Index = 19 Then
            Txttexto.Item(19).Text = Format(Txttexto.Item(19).Text, "dd/mm/yyyy")
        End If ' TERMINA INDICES
        
        'DESPUES DE INGRESAR LA LINEA CAMBIA EL FOCUS AL BATCH
        If Index = 0 Then
            Txttexto.Item(5).SetFocus
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
