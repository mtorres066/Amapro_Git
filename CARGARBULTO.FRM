VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CargarBulto 
   BackColor       =   &H000000FF&
   Caption         =   "Cargar Bulto"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "CargarBulto.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8055
      Left            =   0
      TabIndex        =   46
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   8640
         Picture         =   "CargarBulto.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   50
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   48
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   49
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   5040
         TabIndex        =   47
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   1800
            TabIndex        =   53
            Top             =   360
            Width           =   1575
         End
      End
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
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "CargarBulto.frx":074C
         Height          =   6855
         Left            =   120
         OleObjectBlob   =   "CargarBulto.frx":0768
         TabIndex        =   51
         Top             =   1080
         Width           =   9255
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   720
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   1800
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   555
      Left            =   6360
      Picture         =   "CargarBulto.frx":1143
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   1500
   End
   Begin TabDlg.SSTab TabIngresos 
      Height          =   6975
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CargarBulto.frx":1585
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameNumerosIngresos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CargarBulto.frx":189F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridNumerosIngresos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CargarBulto.frx":1CF1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "Label2(8)"
      Tab(2).Control(2)=   "Label2(9)"
      Tab(2).Control(3)=   "LblBusCodMatPri"
      Tab(2).Control(4)=   "FrameOpciones"
      Tab(2).Control(5)=   "TxtBuscar"
      Tab(2).Control(6)=   "DtpFecIni"
      Tab(2).Control(7)=   "DtpFecFin"
      Tab(2).Control(8)=   "CmdBuscar"
      Tab(2).Control(9)=   "CmdActualizar"
      Tab(2).Control(10)=   "TxtBusCodMatPri"
      Tab(2).ControlCount=   11
      Begin VB.TextBox TxtBusCodMatPri 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67680
         TabIndex        =   29
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Actualizar"
         Height          =   735
         Left            =   -69720
         Picture         =   "CargarBulto.frx":2143
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Datos"
         Height          =   735
         Left            =   -67680
         Picture         =   "CargarBulto.frx":244D
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   5760
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -67680
         TabIndex        =   27
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24772611
         CurrentDate     =   37301
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -67680
         TabIndex        =   26
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24772611
         CurrentDate     =   37301
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -67680
         TabIndex        =   28
         ToolTipText     =   " "
         Top             =   3960
         Width           =   1935
      End
      Begin VB.Frame FrameOpciones 
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
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   -74880
         TabIndex        =   35
         Top             =   720
         Width           =   5205
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha Y Materia Prima"
            Height          =   1305
            Index           =   2
            Left            =   3480
            Picture         =   "CargarBulto.frx":288F
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Numero Ingreso y Codigo Materia Prima"
            Height          =   1305
            Index           =   0
            Left            =   120
            Picture         =   "CargarBulto.frx":3159
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha"
            Height          =   1305
            Index           =   1
            Left            =   1800
            Picture         =   "CargarBulto.frx":3463
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame FrameNumerosIngresos 
         Caption         =   "Datos De Numero Ingreso"
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
         TabIndex        =   33
         Top             =   720
         Width           =   9195
         Begin VB.TextBox TxtCanSal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "CantidadSalida"
            DataSource      =   "DataNumerosIngresos"
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
            Left            =   3840
            Locked          =   -1  'True
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   5760
            Width           =   1455
         End
         Begin VB.TextBox TxtBodSal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "BodegaSalida"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   5760
            Width           =   615
         End
         Begin MSMask.MaskEdBox MskConIni 
            DataField       =   "ContadorInicial"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   7560
            TabIndex        =   15
            Top             =   5760
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Index           =   7
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   11
            Top             =   2760
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Index           =   6
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   10
            Top             =   2400
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Index           =   5
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   2040
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   1680
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   7
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   600
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Left            =   7080
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox TxtFicTec 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "FichaTecnica"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1815
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            TabIndex        =   0
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Text1 
            DataField       =   "UsuarioEditar"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Index           =   1
            Left            =   240
            MaxLength       =   10
            TabIndex        =   44
            Top             =   2280
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            DataField       =   "UsuarioAgregar"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Index           =   0
            Left            =   240
            MaxLength       =   10
            TabIndex        =   43
            Top             =   2640
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            DataField       =   "Linea"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   2
            Top             =   1080
            Width           =   1815
         End
         Begin MSMask.MaskEdBox MskHor 
            DataField       =   "Hora"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            TabIndex        =   1
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Format          =   "HH:mm"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtCod 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Codigo"
            DataSource      =   "DataNumerosIngresos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1200
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   3120
            Width           =   6975
         End
         Begin VB.TextBox TxtNumIng 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "NumeroIngreso"
            DataSource      =   "DataNumerosIngresos"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1185
            Left            =   2040
            TabIndex        =   13
            ToolTipText     =   " "
            Top             =   4440
            Width           =   6975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Existencia"
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
            Left            =   2880
            TabIndex        =   69
            Top             =   5760
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Disponible En Bodega"
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
            Index           =   17
            Left            =   120
            TabIndex        =   68
            Top             =   5760
            Width           =   1890
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Contador Inicial"
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
            Left            =   6120
            TabIndex        =   66
            Top             =   5760
            Width           =   1350
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
            Height          =   255
            Left            =   120
            TabIndex        =   65
            Top             =   1800
            Width           =   3735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Caja"
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
            Left            =   5520
            TabIndex        =   64
            Top             =   2760
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nylon Strech"
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
            Left            =   5520
            TabIndex        =   63
            Top             =   2400
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sello Solvente"
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
            Left            =   5520
            TabIndex        =   62
            Top             =   2040
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Barniz Polvo"
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
            Left            =   5520
            TabIndex        =   61
            Top             =   1680
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Barniz Liquido"
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
            Left            =   5520
            TabIndex        =   60
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Alambre Cobre"
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
            Left            =   5520
            TabIndex        =   59
            Top             =   960
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   6
            Left            =   5520
            TabIndex        =   58
            Top             =   600
            Width           =   1200
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   5
            Left            =   5520
            TabIndex        =   57
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   56
            Top             =   1440
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   39
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Codigo Materia Prima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   795
            Index           =   2
            Left            =   120
            TabIndex        =   38
            Top             =   3240
            Width           =   2025
         End
         Begin VB.Label Label2 
            Caption         =   "Numero Ingreso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   795
            Index           =   1
            Left            =   120
            TabIndex        =   37
            Top             =   4560
            Width           =   1275
         End
         Begin VB.Label Label2 
            Caption         =   "Hora"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   720
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGridNumerosIngresos 
         Bindings        =   "CargarBulto.frx":376D
         Height          =   6105
         Left            =   -74880
         OleObjectBlob   =   "CargarBulto.frx":378F
         TabIndex        =   22
         Top             =   720
         Width           =   9225
      End
      Begin VB.Label LblBusCodMatPri 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -69600
         TabIndex        =   45
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Index           =   9
         Left            =   -68880
         TabIndex        =   42
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Index           =   8
         Left            =   -68880
         TabIndex        =   41
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero Ingreso"
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
         Left            =   -70560
         TabIndex        =   36
         Top             =   3960
         Width           =   2775
      End
   End
   Begin VB.Data DataNumerosIngresos 
      BackColor       =   &H80000014&
      Caption         =   "Numeros Ingresos"
      Connect         =   "Access"
      DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "NumerosIngresosMPaProducir"
      Top             =   7680
      Width           =   9315
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   555
      Left            =   7920
      MouseIcon       =   "CargarBulto.frx":4B98
      Picture         =   "CargarBulto.frx":4FDA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   555
      Left            =   4800
      MouseIcon       =   "CargarBulto.frx":541C
      Picture         =   "CargarBulto.frx":585E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3240
      MouseIcon       =   "CargarBulto.frx":5D90
      Picture         =   "CargarBulto.frx":61D2
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1680
      MouseIcon       =   "CargarBulto.frx":6704
      Picture         =   "CargarBulto.frx":6B46
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   555
      Left            =   120
      MouseIcon       =   "CargarBulto.frx":7078
      Picture         =   "CargarBulto.frx":74BA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1500
   End
End
Attribute VB_Name = "CargarBulto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim Mensaje As String
Dim Buscar As String

Dim RBuscaNumeroIngreso As Recordset
Dim RBuscaNumeroIngresoEntradas As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaExistencia As Recordset
Dim RBuscaLinea As Recordset
Dim RBuscaFichaTecnica As Recordset
Dim RBuscaFicha As Recordset

Dim VCodigoMateriaPrima As String
Dim VHojalata As String
Dim VNumeroIngreso As Double
Dim VBodega As String
Dim VCantidadMateriaPrima As Double

Dim BCamposVisibles As Boolean
Dim BLinea As Boolean
Dim BMateriaPrima As Boolean


Sub botones()
    If Bandera = True Then
         FrameNumerosIngresos.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataNumerosIngresos.Visible = False
         FrameOpciones.Visible = False
         DBGridNumerosIngresos.Visible = False
    Else
         FrameNumerosIngresos.Enabled = False
         CmdAgregar.Enabled = True
         
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataNumerosIngresos.Visible = True
         FrameOpciones.Visible = True
         DBGridNumerosIngresos.Visible = True
    End If
End Sub

Private Sub CmdActualizar_Click()
    DataNumerosIngresos.RecordSource = "Select * From NumerosIngresosMPaProducir"
    DataNumerosIngresos.Refresh
    DBGridNumerosIngresos.Refresh
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
       'AGREGA UN REGISTRO
        DataNumerosIngresos.Recordset.AddNew
        'SI HAY ERRORES
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        
        'BOTONES DISPONIBLES
        Bandera = True
        botones
     
        'ASIGNA FECHA Y HORA
        MskFec.Text = Date
        MskHor.Text = Time
        'ASIGNA USUARIO
        Text1.Item(0).Text = GUsuario
        TxtLin.SetFocus
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            
            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If Mensaje = vbOK Then
   
                         VHojalata = TxtTexto.Item(0).Text
                         VCodigoMateriaPrima = TxtCod.Text
                         VBodega = TxtBodSal.Text
                         VCantidadMateriaPrima = TxtCanSal.Text
                         VNumeroIngreso = TxtNumIng.Text
                            
                        'GRABA DATOS
                        DataNumerosIngresos.Recordset.Delete
                        
                        If Err <> 0 Then
                             MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Else
                        
                             'SI ESTA USANDO UN BULTO DE HOJALATA NO REBAJA EL BULTO
                             If VCodigoMateriaPrima = VHojalata Then
                             
                             'CON CUALQUIER OTRO TIPO DE MATERIA PRIMA SI REBAJA DE INVENTARIO LA CANTIDAD QUE VIENE
                             Else
                                     'BUSCA LA EXISTENCIA Y LAS SALIDAS DE LA MATERIA PRIMA Y RESTA LA EXISTENCIA Y SUMA SALIDAS
                                     Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Salidas From InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and Bodega = '" & VBodega & "'")
                                        If RBuscaMateriaPrima.RecordCount > 0 Then
                                                RBuscaMateriaPrima.Edit
                                                    RBuscaMateriaPrima!Existencia = (RBuscaMateriaPrima!Existencia + VCantidadMateriaPrima)
                                                    RBuscaMateriaPrima!Salidas = (RBuscaMateriaPrima!Salidas - VCantidadMateriaPrima)
                                                RBuscaMateriaPrima.Update
                                        End If
                                     
                                     'BUSCA EL NUMERO DE INGRESO CON CODIGO DE MATERIA PRIMA EN EL DETALLE DE LAS ENTRADAS
                                      Set RBuscaNumeroIngresoEntradas = Db.OpenRecordset("Select Disponibilidad, SaldoDisponibilidad, CantidadSalida From DetalleEntradasMateriaPrima Where Codigo = '" & VCodigoMateriaPrima & "' And NumeroIngreso = " & VNumeroIngreso)
                                          If RBuscaNumeroIngresoEntradas.RecordCount > 0 Then
                                                 'ASIGNA LA DISPONIBILIDAD AL BULTO O NUMERO DE INGRESO Y CONTROLA EL SALDO DEL BULTO
                                                 RBuscaNumeroIngresoEntradas.Edit
                                                         RBuscaNumeroIngresoEntradas!CantidadSalida = RBuscaNumeroIngresoEntradas!CantidadSalida - VCantidadMateriaPrima
                                                         RBuscaNumeroIngresoEntradas!SaldoDisponibilidad = RBuscaNumeroIngresoEntradas!SaldoDisponibilidad + VCantidadMateriaPrima
                                                         'SI EL SALDO QUE QUEDA ES CERO ENTONCES EL BULTO YA NO EXISTE
                                                         If RBuscaNumeroIngresoEntradas!SaldoDisponibilidad > 0 Then
                                                             RBuscaNumeroIngresoEntradas!Disponibilidad = "EXISTE"
                                                         End If
                                                 RBuscaNumeroIngresoEntradas.Update
                                                 If Err <> 0 Then
                                                 End If
                                          End If
                             End If
                             
                              Bandera = False
                              botones
                              CmdAgregar.SetFocus
                        End If

                
                DataNumerosIngresos.Recordset.MoveLast
            End If
  
            If DataNumerosIngresos.Recordset.EOF Then
                DataNumerosIngresos.Recordset.MoveLast
                If Err = 3021 Then
                    Mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdBuscar_Click()
        'NUMERO INGRESO Y CODIGO MATERIA PRIMA
        If OptBusqueda.Item(0).Value = True Then
            If Not IsNumeric(TxtBuscar.Text) Then
                    MsgBox "El Numero De Ingreso Debe Ser Numerico", vbOKOnly + vbExclamation, "Informacion"
                    TxtBuscar.SetFocus
                    Exit Sub
            End If
            DataNumerosIngresos.RecordSource = "Select * From NumerosIngresosMPaProducir Where NumeroIngreso = " & TxtBuscar.Text & " And Codigo = '" & TxtBusCodMatPri.Text & "'"
            TabIngresos.Tab = 0
        'FECHAS
        ElseIf OptBusqueda.Item(1).Value = True Then
            DataNumerosIngresos.RecordSource = "Select * From NumerosIngresosMPaProducir Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#"
            TabIngresos.Tab = 1
        'FECHA Y MATERIA PRIMA
        ElseIf OptBusqueda.Item(2).Value = True Then
            DataNumerosIngresos.RecordSource = "Select * From NumerosIngresosMPaProducir Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Codigo = '" & TxtBuscar.Text & "'"
            TabIngresos.Tab = 1
        End If
        DataNumerosIngresos.Refresh
        DBGridNumerosIngresos.Refresh
        
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
        
        DataNumerosIngresos.Recordset.CancelUpdate
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        
        Bandera = False
        botones

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
        'EDITA EL REGISTRO
        DataNumerosIngresos.Recordset.Edit
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        
        'BOTONES DISPONIBLES
        Bandera = True
        botones
        
        'ASIGNA USUARIO
        Text1.Item(1).Text = GUsuario
                
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   
    VHojalata = TxtTexto.Item(0).Text
    VCodigoMateriaPrima = TxtCod.Text
    VBodega = TxtBodSal.Text
    VCantidadMateriaPrima = TxtCanSal.Text
    VNumeroIngreso = TxtNumIng.Text
    
    'REVISA LA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
    End If
          
    'REVISA EL CODIGO DE LA MATERIA PRIMA
    If TxtCod.Text = "" Then
        MsgBox "Codigo De Materia Prima No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    'REVISA LAS LINEAS 03 O 04 PARA QUE PUEDA TRABAJAR CON LINEAS 01 O 02
    If TxtLin.Text = "03" Then
        MsgBox "Si Quiere Trabajar Con Linea 03, Utilize Linea 01", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    If TxtLin.Text = "04" Then
        MsgBox "Si Quiere Trabajar Con Linea 04, Utilize Linea 02", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    'REVISA SI EXISTE EL BULTO
    Set RBuscaNumeroIngreso = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCod.Text & "'")
        If RBuscaNumeroIngreso.RecordCount > 0 Then
        Else
            MsgBox "Numero Ingreso No Existe Para Esta Materia Prima", vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
       
   'GRABA DATOS
   DataNumerosIngresos.Recordset.Update
   
   If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
   Else
         Bandera = False
         botones
         CmdAgregar.SetFocus
   End If

   
End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next

MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDocIng.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
                CrReportes.SelectionFormula = "{NumerosIngresosMPaProducir.NumeroIngreso} = " & TxtNumIng.Text & " And {NumerosIngresosMPaProducir.Codigo} = '" & TxtCod.Text & "'"
                CrReportes.ReportFileName = App.Path & "\FormatoNumerosIngresosMPaProducir.rpt"
                CrReportes.Action = 1
                
                If Err <> 0 Then
                     MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
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
    'LINEAS
    If BLinea = True Then
        TxtLin.Text = DBGridConsultas.Columns(0).Text
        TxtLin.SetFocus
    'MATERIAS PRIMAS
    ElseIf BMateriaPrima = True Then
        TxtCod.Text = DBGridConsultas.Columns(0).Text
        TxtCod.SetFocus
    End If
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        'LINEAS
        If BLinea = True Then
            TxtLin.Text = DBGridConsultas.Columns(0).Text
            TxtLin.SetFocus
        'MATERIAS PRIMAS
        ElseIf BMateriaPrima = True Then
            TxtCod.Text = DBGridConsultas.Columns(0).Text
            TxtCod.SetFocus
        End If
            FrameConsultas.Visible = False
    End If
End Sub

Private Sub dbgridnumerosingresos_HeadClick(ByVal ColIndex As Integer)
    DataNumerosIngresos.RecordSource = ("Select * from NumerosIngresosMPaProducir order by " & DBGridNumerosIngresos.Columns(ColIndex).DataField)
    DataNumerosIngresos.Refresh
    DBGridNumerosIngresos.Refresh
    
End Sub

Private Sub Form_Load()
    DataNumerosIngresos.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
End Sub

Private Sub MskConIni_GotFocus()
        MskConIni.SelStart = 0
        MskConIni.SelLength = Len(MskConIni.Text)
End Sub

Private Sub MskConIni_KeyPress(KeyAscii As Integer)
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

Private Sub OptBusqueda_Click(Index As Integer)
        If Index = 0 Then
            Label2.Item(8).Visible = False
            Label2.Item(9).Visible = False
            Lbletiqueta.Caption = "Numero Ingreso"
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            DtpFecIni.Visible = False
            DtpFecFin.Visible = False
            LblBusCodMatPri.Visible = True
            TxtBusCodMatPri.Visible = True
        ElseIf Index = 1 Then
            Label2.Item(8).Visible = True
            Label2.Item(9).Visible = True
            Lbletiqueta.Caption = ""
            TxtBuscar.Visible = False
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            LblBusCodMatPri.Visible = False
            TxtBusCodMatPri.Visible = False
        ElseIf Index = 2 Then
            Label2.Item(8).Visible = True
            Label2.Item(9).Visible = True
            Lbletiqueta.Caption = "Codigo Materia Prima"
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            LblBusCodMatPri.Visible = False
            TxtBusCodMatPri.Visible = False
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

Private Sub TxtBusCodMatPri_GotFocus()
        TxtBusCodMatPri.SelStart = 0
        TxtBusCodMatPri.SelLength = Len(TxtBusCodMatPri.Text)
End Sub

Private Sub TxtBusCodMatPri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub TxtCanSal_GotFocus()
        TxtCanSal.SelStart = 0
        TxtCanSal.SelLength = Len(TxtCanSal.Text)
End Sub

Private Sub TxtCanSal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtConsultas_Change()
    'LINEA
    If BLinea = True Then
        'DESCRIPCION
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec And L.Descrip Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec And L.Descrip Like '*" & TxtConsultas.Text & "*'"
            End If
        'CODIGO
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec And L.Linea Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec And L.Linea Like '*" & TxtConsultas.Text & "*'"
            End If
        End If
    'MATERIA PRIMA
    ElseIf BMateriaPrima = True Then
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion, Bodega From InventarioMateriaPrima Where Descripcion Like '" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima, Bodega"
            Else
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion, Bodega From InventarioMateriaPrima Where Descripcion Like '*" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima, Bodega"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion, Bodega From InventarioMateriaPrima Where CodigoMateriaPrima Like '" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima, Bodega"
            Else
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion, Bodega From InventarioMateriaPrima Where CodigoMateriaPrima Like '*" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima, Bodega"
            End If
        End If
    End If
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    AnchoColumnas
    
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


Private Sub TxtFicTec_Change()
    'BUSCAMOS LOS CODIGOS DE MATERIA PRIMA DE ACUERDO A LA FICHA TECNICA
        Set RBuscaFicha = Db.OpenRecordset("Select * From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
        If RBuscaFicha.RecordCount > 0 Then
            LblFichaTecnica.Caption = RBuscaFicha!Descrip
                    TxtTexto.Item(0).Text = RBuscaFicha!Platina
                    If IsNull(RBuscaFicha!Fondo) Then
                        TxtTexto.Item(1).Text = RBuscaFicha!Tapa
                    Else
                        TxtTexto.Item(1).Text = RBuscaFicha!Fondo
                    End If
                    TxtTexto.Item(2).Text = RBuscaFicha!CodAla
                    TxtTexto.Item(3).Text = RBuscaFicha!CodBarLiq
                    TxtTexto.Item(4).Text = RBuscaFicha!CodBarPol
                    TxtTexto.Item(5).Text = RBuscaFicha!CodSelSol
                    TxtTexto.Item(6).Text = RBuscaFicha!CodNylStr
                    If IsNull(RBuscaFicha!CodCaja) Then
                        TxtTexto.Item(7).Text = ""
                    Else
                        TxtTexto.Item(7).Text = RBuscaFicha!CodCaja
                    End If
            
        Else
                    LblFichaTecnica.Caption = ""
                    TxtTexto.Item(0).Text = ""
                    TxtTexto.Item(1).Text = ""
                    TxtTexto.Item(2).Text = ""
                    TxtTexto.Item(3).Text = ""
                    TxtTexto.Item(4).Text = ""
                    TxtTexto.Item(5).Text = ""
                    TxtTexto.Item(6).Text = ""
                    TxtTexto.Item(7).Text = ""
                    
        End If
End Sub

Private Sub TxtLin_DblClick()
    BLinea = True
    BMateriaPrima = False
    DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec"
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    AnchoColumnas
    FrameConsultas.Visible = True
    TxtConsultas.SetFocus
    
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
            BMateriaPrima = False
            DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec"
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            AnchoColumnas
            FrameConsultas.Visible = True
            TxtConsultas.SetFocus
    End If
End Sub

Private Sub TxtLin_LostFocus()
    'BUSCAMOS LA LINEA PARA JALAR LA FICHA TECNICA
    Set RBuscaLinea = Db.OpenRecordset("Select Esp_Tec From Lineas Where Linea = '" & TxtLin.Text & "'")
    If RBuscaLinea.RecordCount > 0 Then
        'BUSCAMOS LOS CODIGOS DE MATERIA PRIMA DE ACUERDO A LA FICHA TECNICA
        Set RBuscaFichaTecnica = Db.OpenRecordset("Select * From FichaTecnica Where Esp_Tec = '" & RBuscaLinea!Esp_tec & "'")
            'SI ENCUENTRA LA FICHA TECNICA
            If RBuscaFichaTecnica.RecordCount > 0 Then
                    'ASIGNA EL CODIGO DE FICHA TECNICA
                    TxtFicTec.Text = RBuscaFichaTecnica!Esp_tec
                    
                    TxtTexto.Item(0).Text = RBuscaFichaTecnica!Platina
                    If IsNull(RBuscaFichaTecnica!Fondo) Then
                        TxtTexto.Item(1).Text = RBuscaFichaTecnica!Tapa
                    Else
                        TxtTexto.Item(1).Text = RBuscaFichaTecnica!Fondo
                    End If
                    TxtTexto.Item(2).Text = RBuscaFichaTecnica!CodAla
                    TxtTexto.Item(3).Text = RBuscaFichaTecnica!CodBarLiq
                    TxtTexto.Item(4).Text = RBuscaFichaTecnica!CodBarPol
                    TxtTexto.Item(5).Text = RBuscaFichaTecnica!CodSelSol
                    TxtTexto.Item(6).Text = RBuscaFichaTecnica!CodNylStr
                    If IsNull(RBuscaFichaTecnica!CodCaja) Then
                        TxtTexto.Item(7).Text = ""
                    Else
                        TxtTexto.Item(7).Text = RBuscaFichaTecnica!CodCaja
                    End If
            End If
    Else
        MsgBox "Linea No Existe", vbOKOnly + vbInformation, "Informacion"
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
End Sub


Sub AnchoColumnas()
    If BLinea = True Then
        DBGridConsultas.Columns(0).Width = "500"
        DBGridConsultas.Columns(1).Width = "3000"
        DBGridConsultas.Columns(2).Width = "1200"
        DBGridConsultas.Columns(3).Width = "4000"
    Else
        DBGridConsultas.Columns(0).Width = "1200"
        DBGridConsultas.Columns(1).Width = "4000"
        DBGridConsultas.Columns(2).Width = "1200"
    End If
    
        
End Sub

Private Sub TxtNumIng_LostFocus()
            If IsNumeric(TxtNumIng.Text) Then
                'BUSCA EL NUMERO DE INGRESO Y ASIGNA LA BODEGA, CODIGO Y CANTIDAD DE ACUERDO COMO ENTRO A LA BODEGA
                Set RBuscaNumeroIngreso = Db.OpenRecordset("Select BodegaDisponibilidad, SaldoDisponibilidad From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCod.Text & "'")
                'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
                If RBuscaNumeroIngreso.RecordCount > 0 Then
                    TxtBodSal.Text = RBuscaNumeroIngreso!BodegaDisponibilidad
                    TxtCanSal.Text = RBuscaNumeroIngreso!SaldoDisponibilidad
                    MskConIni.SetFocus
                Else
                    TxtBodSal.Text = ""
                    TxtCanSal.Text = "0"
                    MskConIni.SetFocus
                End If
            End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    TxtCod.Text = TxtTexto.Item(Index).Text
    TxtNumIng.SetFocus
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
        TxtCod.Text = TxtTexto.Item(Index).Text
        TxtNumIng.SetFocus
    End If
End Sub
