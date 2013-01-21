VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form BultosMateriaPrima 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento De Bultos De Materia Prima"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "BultosMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8055
      Left            =   0
      TabIndex        =   66
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   8640
         Picture         =   "BultosMateriaPrima.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   72
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   71
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   70
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   5040
         TabIndex        =   67
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   2040
            TabIndex        =   69
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Data DataConsultas 
         Caption         =   "consultas"
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
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "BultosMateriaPrima.frx":293C
         Height          =   6735
         Left            =   120
         OleObjectBlob   =   "BultosMateriaPrima.frx":2958
         TabIndex        =   73
         ToolTipText     =   "Signo '+' o Doble Click Para Seleccionar"
         Top             =   1080
         Width           =   11535
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.Data DataBultos 
      Caption         =   "Bultos De Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\erick\Amapro Metalenvases\metalenvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   320
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleEntradasMateriaPrima"
      Top             =   7800
      Width           =   11655
   End
   Begin TabDlg.SSTab TabBultos 
      Height          =   7215
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12726
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "BultosMateriaPrima.frx":3333
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBultos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "BultosMateriaPrima.frx":364D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridBultos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "BultosMateriaPrima.frx":3A9F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin MSDBGrid.DBGrid DGridBultos 
         Bindings        =   "BultosMateriaPrima.frx":3EF1
         Height          =   6375
         Left            =   -74880
         OleObjectBlob   =   "BultosMateriaPrima.frx":3F0A
         TabIndex        =   26
         Top             =   720
         Width           =   11655
      End
      Begin VB.Frame FrameBusquedadeDatos 
         Caption         =   "Busqueda de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6375
         Left            =   -74880
         TabIndex        =   39
         Top             =   720
         Width           =   11655
         Begin VB.OptionButton OptBusqueda 
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
            Height          =   1335
            Index           =   7
            Left            =   10200
            Picture         =   "BultosMateriaPrima.frx":6C41
            Style           =   1  'Graphical
            TabIndex        =   62
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Numero De Bulto/Paleta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   6
            Left            =   8760
            Picture         =   "BultosMateriaPrima.frx":750B
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Numero De Serie"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   5
            Left            =   7320
            Picture         =   "BultosMateriaPrima.frx":7DD5
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   360
            Width           =   1300
         End
         Begin VB.TextBox TxtBusqueda2 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8400
            TabIndex        =   29
            Top             =   3600
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Bodega Y Codigo Materia Prima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   4
            Left            =   5880
            Picture         =   "BultosMateriaPrima.frx":80DF
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Bodega Disponible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   3
            Left            =   4440
            Picture         =   "BultosMateriaPrima.frx":89A9
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton OptBusqueda 
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
            Height          =   1335
            Index           =   2
            Left            =   3000
            Picture         =   "BultosMateriaPrima.frx":8DEB
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   360
            Width           =   1300
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8400
            TabIndex        =   30
            Top             =   3960
            Width           =   2535
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Numero Bulto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   1
            Left            =   1560
            Picture         =   "BultosMateriaPrima.frx":90FD
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   360
            Width           =   1300
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Codigo Materia Prima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Index           =   0
            Left            =   120
            Picture         =   "BultosMateriaPrima.frx":9407
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   1300
         End
         Begin VB.Label LblBusqueda2 
            Caption         =   "Codigo Materia Prima"
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
            Left            =   6480
            TabIndex        =   58
            Top             =   3600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.Label LblBusqueda 
            Alignment       =   1  'Right Justify
            Caption         =   "Codigo Materia Prima"
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
            Left            =   6240
            TabIndex        =   40
            Top             =   3960
            Width           =   2055
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   7
            Left            =   8400
            TabIndex        =   32
            Top             =   5160
            Width           =   2535
            Caption         =   "Seleccionar Todos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   6
            Left            =   8400
            TabIndex        =   31
            Top             =   4440
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "BultosMateriaPrima.frx":9CD1
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
         End
      End
      Begin VB.Frame FrameBultos 
         Caption         =   "Datos del Bulto"
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
         Height          =   4935
         Left            =   360
         TabIndex        =   33
         Top             =   1320
         Width           =   11175
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "OrdenProduccion"
            DataSource      =   "DataBultos"
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
            Index           =   3
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1560
            Width           =   1920
         End
         Begin VB.ComboBox CboEstado 
            DataField       =   "Estado"
            DataSource      =   "DataBultos"
            Height          =   315
            ItemData        =   "BultosMateriaPrima.frx":9FEB
            Left            =   8760
            List            =   "BultosMateriaPrima.frx":9FF5
            TabIndex        =   7
            Text            =   "I"
            Top             =   1200
            Width           =   615
         End
         Begin VB.ComboBox CboCalidad 
            DataField       =   "Calidad"
            DataSource      =   "DataBultos"
            Height          =   288
            ItemData        =   "BultosMateriaPrima.frx":9FFF
            Left            =   1800
            List            =   "BultosMateriaPrima.frx":A00C
            TabIndex        =   5
            Text            =   "A"
            Top             =   2280
            Width           =   615
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "PesoEntrada"
            DataSource      =   "DataBultos"
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
            Index           =   23
            Left            =   8760
            TabIndex        =   14
            Top             =   2640
            Width           =   2280
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Peso"
            DataSource      =   "DataBultos"
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
            Index           =   22
            Left            =   8760
            TabIndex        =   18
            Top             =   4080
            Width           =   2280
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "CantidadSalida"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   19
            Left            =   8760
            TabIndex        =   16
            Top             =   3360
            Width           =   2280
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "SaldoDisponibilidad"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   20
            Left            =   8760
            TabIndex        =   17
            Top             =   3720
            Width           =   2280
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "CantidadTraslado"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   18
            Left            =   8760
            TabIndex        =   15
            Top             =   3000
            Width           =   2280
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "BodegaDisponibilidad"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   6
            Left            =   1800
            MaxLength       =   3
            TabIndex        =   4
            Top             =   1920
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   11
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   6
            Top             =   4440
            Width           =   9240
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NumeroIngreso"
            DataSource      =   "DataBultos"
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
            Left            =   1800
            TabIndex        =   2
            Top             =   1200
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "BobinaBoleta"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   16
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   12
            Top             =   4080
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "FechaBoleta"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   15
            Left            =   1800
            TabIndex        =   11
            Top             =   3720
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "BultoBoleta"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   14
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   10
            Top             =   3360
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "OrdenBoleta"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   13
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   9
            Top             =   3000
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "NumeroUnicoSerieBoleta"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   12
            Left            =   1800
            MaxLength       =   20
            TabIndex        =   8
            Top             =   2640
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Cantidad"
            DataSource      =   "DataBultos"
            Height          =   285
            Index           =   17
            Left            =   8760
            MaxLength       =   30
            TabIndex        =   13
            Top             =   2280
            Width           =   2280
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Codigo"
            DataSource      =   "DataBultos"
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   1
            Top             =   840
            Width           =   1920
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Documento"
            DataSource      =   "DataBultos"
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   0
            Top             =   480
            Width           =   1920
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Index           =   9
            Left            =   120
            TabIndex        =   65
            Top             =   1560
            Width           =   1500
         End
         Begin VB.Label LblEstado 
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
            Left            =   9480
            TabIndex        =   64
            Top             =   1200
            Width           =   1575
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Height          =   195
            Index           =   1
            Left            =   7200
            TabIndex        =   63
            Top             =   1200
            Width           =   600
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Peso De Entrada"
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
            Index           =   18
            Left            =   7200
            TabIndex        =   59
            Top             =   2640
            Width           =   1452
         End
         Begin VB.Label LblRecepcion 
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
            TabIndex        =   54
            Top             =   480
            Width           =   7095
         End
         Begin VB.Label LblBodegaDisponible 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   252
            Left            =   3960
            TabIndex        =   53
            Top             =   1920
            Width           =   7092
         End
         Begin VB.Label LblMateriaPrima 
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
            TabIndex        =   52
            Top             =   840
            Width           =   7095
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Peso Actual"
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
            Index           =   25
            Left            =   7200
            TabIndex        =   51
            Top             =   4200
            Width           =   1035
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Salidas"
            Height          =   195
            Index           =   23
            Left            =   7200
            TabIndex        =   50
            Top             =   3480
            Width           =   510
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Saldo"
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
            Left            =   7200
            TabIndex        =   49
            Top             =   3840
            Width           =   495
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Inicio"
            Height          =   195
            Index           =   21
            Left            =   7200
            TabIndex        =   48
            Top             =   3120
            Width           =   375
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bodega Disponib."
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
            Index           =   17
            Left            =   120
            TabIndex        =   47
            Top             =   1920
            Width           =   1512
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bobina"
            Height          =   192
            Index           =   14
            Left            =   120
            TabIndex        =   46
            Top             =   4200
            Width           =   492
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "# De Ingreso"
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
            Left            =   120
            TabIndex        =   45
            Top             =   1200
            Width           =   1125
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Height          =   192
            Index           =   12
            Left            =   120
            TabIndex        =   44
            Top             =   2280
            Width           =   648
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   43
            Top             =   4560
            Width           =   1065
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            Height          =   192
            Index           =   8
            Left            =   120
            TabIndex        =   42
            Top             =   3840
            Width           =   456
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Orden ó Lote"
            Height          =   192
            Index           =   6
            Left            =   120
            TabIndex        =   41
            Top             =   3120
            Width           =   936
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Bulto ó Paleta"
            Height          =   192
            Index           =   5
            Left            =   120
            TabIndex        =   38
            Top             =   3480
            Width           =   996
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Numero Serie"
            Height          =   192
            Index           =   4
            Left            =   120
            TabIndex        =   37
            Top             =   2760
            Width           =   960
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad Entrada"
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
            Index           =   3
            Left            =   7200
            TabIndex        =   36
            Top             =   2280
            Width           =   1488
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Codigo Mat. Prima"
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
            TabIndex        =   35
            Top             =   840
            Width           =   1560
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   480
            Width           =   1065
         End
      End
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   7320
      Width           =   1800
      Caption         =   "Agregar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Accelerator     =   65
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   1
      Left            =   2040
      TabIndex        =   21
      Top             =   7320
      Width           =   1800
      Caption         =   "Editar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Picture         =   "BultosMateriaPrima.frx":A019
      Accelerator     =   69
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   2
      Left            =   3960
      TabIndex        =   22
      Top             =   7320
      Width           =   1800
      VariousPropertyBits=   25
      Caption         =   "Grabar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Accelerator     =   71
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   3
      Left            =   5880
      TabIndex        =   23
      Top             =   7320
      Width           =   1800
      VariousPropertyBits=   25
      Caption         =   "Cancelar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Accelerator     =   67
      FontEffects     =   1073750017
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   4
      Left            =   7800
      TabIndex        =   24
      Top             =   7320
      Width           =   1800
      Caption         =   "Borrar"
      PicturePosition =   196613
      Size            =   "3175;661"
      Accelerator     =   66
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin MSForms.CommandButton CmdBotones 
      Height          =   375
      Index           =   5
      Left            =   9720
      TabIndex        =   25
      Top             =   7320
      Width           =   1800
      Caption         =   "Salida"
      PicturePosition =   196613
      Size            =   "3175;661"
      Picture         =   "BultosMateriaPrima.frx":A55B
      Accelerator     =   83
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
End
Attribute VB_Name = "BultosMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer

Dim RBuscaRecepcion As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaDefecto1 As Recordset
Dim RBuscaDefecto2 As Recordset
Dim RBuscaDefecto3 As Recordset
Dim RBuscaBodegaEntrada As Recordset
Dim RBuscaBodegaDisponible As Recordset






Private Sub CboEstado_Change()
        If CboEstado.Text = "I" Then
            LblEstado.Caption = "INSPECIONADO"
        ElseIf CboEstado.Text = "N" Then
            LblEstado.Caption = "NO INSPECIONADO"
        End If
End Sub

Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataBultos.Recordset
    
        'AGREGAR
        If Index = 0 Then
                .AddNew
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
        'EDITAR
        ElseIf Index = 1 Then
                        .Edit
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
        'GRABAR
        ElseIf Index = 2 Then
                    
                .Update
                
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
                CmdBotones.Item(0).SetFocus
        'CANCELAR
        ElseIf Index = 3 Then
                .CancelUpdate
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
        ElseIf Index = 4 Then ' BORRAR
        
                VMensaje = MsgBox("Esta seguro de borrar el registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                If VMensaje = vbYes Then
                    .Delete
                    .MoveLast
                            If Err.Number > 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                End If
        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    If OptBusqueda.Item(0).Value = True Then
                        DataBultos.RecordSource = ("Select * From DetalleEntradasMateriaPrima where Codigo = '" & Txtbusqueda.Text & "' Order By NumeroIngreso")
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataBultos.RecordSource = ("Select * From DetalleEntradasMateriaPrima where NumeroIngreso = " & Txtbusqueda.Text)
                    ElseIf OptBusqueda.Item(2).Value = True Then
                        If IsNumeric(Txtbusqueda.Text) Then
                            DataBultos.RecordSource = ("Select * From DetalleEntradasMateriaPrima where Documento = " & Txtbusqueda.Text & " Order By NumeroIngreso")
                        Else
                            MsgBox "Transaccion Debe Ser Numerica", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                    ElseIf OptBusqueda.Item(3).Value = True Then
                        DataBultos.RecordSource = ("Select * From DetalleEntradasMateriaPrima where BodegaDisponibilidad = '" & Txtbusqueda.Text & "' Order By NumeroIngreso")
                    ElseIf OptBusqueda.Item(4).Value = True Then
                        DataBultos.RecordSource = ("Select * From DetalleEntradasMateriaPrima where BodegaDisponibilidad = '" & Txtbusqueda.Text & "' AND Codigo = '" & TxtBusqueda2.Text & "' Order By NumeroIngreso")
                    ElseIf OptBusqueda.Item(5).Value = True Then
                        DataBultos.RecordSource = ("Select * From DetalleEntradasMateriaPrima where NumeroUnicoSerieBoleta Like '*" & Txtbusqueda.Text & "*' Order By NumeroIngreso")
                    ElseIf OptBusqueda.Item(6).Value = True Then
                        DataBultos.RecordSource = ("Select * From DetalleEntradasMateriaPrima where BultoBoleta like '*" & Txtbusqueda.Text & "*' Order By NumeroIngreso")
                    ElseIf OptBusqueda.Item(7).Value = True Then
                        DataBultos.RecordSource = ("Select * From DetalleEntradasMateriaPrima where OrdenProduccion = '" & Txtbusqueda.Text & "' Order By NumeroIngreso")
                    End If
                    DataBultos.Refresh
                    DGridBultos.Refresh
                    TabBultos.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataBultos.RecordSource = "Select * From DetalleEntradasMateriaPrima"
                    DataBultos.Refresh
                    DGridBultos.Refresh
                    TabBultos.Tab = 1
        End If
    End With
    

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameBultos.Enabled = True
         DataBultos.Visible = False
         DGridBultos.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameBultos.Enabled = False
         DataBultos.Visible = True
         DGridBultos.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub CmdSale_Click()
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_DblClick()
        Txttexto.Item(6).Text = DBGridConsultas.Columns(0)
        Txttexto.Item(6).SetFocus
        TxtConsultas.Text = ""
        FrameConsultas.Visible = False

End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                Txttexto.Item(6).Text = DBGridConsultas.Columns(0)
                Txttexto.Item(6).SetFocus
                TxtConsultas.Text = ""
                FrameConsultas.Visible = False
        End If
End Sub

Private Sub DGridBultos_HeadClick(ByVal ColIndex As Integer)
        DataBultos.RecordSource = "Select * From DetalleEntradasMateriaPrima Order By " & DGridBultos.Columns(ColIndex).DataField
        DataBultos.Refresh
        DGridBultos.Refresh
End Sub

Private Sub Form_Load()
    DataBultos.ConnectionString = GTipoProveedor
    DataBultos.Refresh

    DataConsultas.ConnectionString = GTipoProveedor
    DataConsultas.Refresh
End Sub


Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            LblBusqueda.Caption = "Codigo Materia Prima"
    ElseIf Index = 1 Then
            LblBusqueda.Caption = "Numero Ingreso"
    ElseIf Index = 2 Then
            LblBusqueda.Caption = "No. Recepcion"
    ElseIf Index = 3 Then
            LblBusqueda.Caption = "Bodega Disponible"
    ElseIf Index = 4 Then
            LblBusqueda.Caption = "Bodega Disponible"
    ElseIf Index = 5 Then
            LblBusqueda.Caption = "Numero Serie"
    ElseIf Index = 6 Then
            LblBusqueda.Caption = "Bulto o Paleta"
    ElseIf Index = 7 Then
            LblBusqueda.Caption = "Orden"
    End If
            Txtbusqueda.SetFocus
            
            'OPCION DE BODEGA Y CODIGO
            If Index = 4 Then
                LblBusqueda2.Visible = True
                TxtBusqueda2.Visible = True
            Else
                LblBusqueda2.Visible = False
                TxtBusqueda2.Visible = False
            End If
End Sub

Private Sub TxtBusqueda_GotFocus()
    Txtbusqueda.SelStart = 0
    Txtbusqueda.SelLength = Len(Txtbusqueda.Text)
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub TxtBusqueda2_GotFocus()
        TxtBusqueda2.SelStart = 0
        TxtBusqueda2.SelLength = Len(Txtbusqueda.Text)
End Sub

Private Sub TxtBusqueda2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtConsultas_Change()
    'SI VA A BUSCAR POR CODIGO
        If OptCod.Value = True Then
            If OptPalIni.Value = True Then
                    DataConsultas.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where CodigoBodega Like '" & TxtConsultas.Text & "*' Order by CodigoBodega")
            Else
                    DataConsultas.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where CodigoBodega Like '*" & TxtConsultas.Text & "*' Order by CodigoBodega")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDes.Value = True Then
            If OptPalIni.Value = True Then
                    DataConsultas.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where Descripcion Like '" & TxtConsultas.Text & "*' Order by CodigoBodega")
            Else
                    DataConsultas.RecordSource = ("Select CodigoBodega, Descripcion from BodegasMateriaPrima Where Descripcion Like '*" & TxtConsultas.Text & "*' Order by CodigoBodega")
            End If
        End If
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            DBGridConsultas.Columns(1).Width = "4000"
    
End Sub

Private Sub TxtConsultas_GotFocus()
        TxtConsultas.SelStart = 0
        TxtConsultas.SelLength = Len(TxtConsultas.Text)
End Sub

Private Sub TxtConsultas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        'RECEPCION
        If Index = 0 Then
            'REVISA SI ES NUMERICO
            If IsNumeric(Txttexto.Item(0).Text) Then
                Set RBuscaRecepcion = Db.OpenRecordset("Select FechaEntrada From EncabezadoEntradasMateriaPrima Where Documento = " & Txttexto.Item(0).Text)
                    If RBuscaRecepcion.RecordCount > 0 Then
                        LblRecepcion.Caption = RBuscaRecepcion!FechaEntrada
                    Else
                        LblRecepcion.Caption = ""
                    End If
            End If
        'CODIGO MATERIA PRIMA
        ElseIf Index = 1 Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & Txttexto.Item(1).Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblMateriaPrima.Caption = ""
                End If
        'BODEGA DISPONIBLE
        ElseIf Index = 6 Then
            Set RBuscaBodegaEntrada = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & Txttexto.Item(6).Text & "'")
                If RBuscaBodegaEntrada.RecordCount > 0 Then
                    LblBodegaDisponible.Caption = RBuscaBodegaEntrada!Descripcion
                Else
                    LblBodegaDisponible.Caption = ""
                End If
        
        End If
                
        
                
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        DataConsultas.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima"
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        DBGridConsultas.Columns(1).Width = "4000"
        FrameConsultas.Visible = True
        TxtConsultas.SetFocus

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
    Txttexto.Item(Index).SelStart = 0
    Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
                SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
                DataConsultas.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima"
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                DBGridConsultas.Columns(1).Width = "4000"
                FrameConsultas.Visible = True
                TxtConsultas.SetFocus
        End If
End Sub
