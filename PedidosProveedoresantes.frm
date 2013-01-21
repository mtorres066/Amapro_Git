VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form PedidosProveedores 
   BackColor       =   &H00008000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pedidos De Materia Prima A PROVEEDORES"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "PedidosProveedoresantes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda De Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   7680
      TabIndex        =   22
      Top             =   8160
      Visible         =   0   'False
      Width           =   11655
      Begin VB.Data DataBusqueda 
         Caption         =   "Busqueda"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2760
         Visible         =   0   'False
         Width           =   3495
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Sale De Busqueda"
         Top             =   360
         Width           =   855
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   5880
         TabIndex        =   15
         Top             =   240
         Width           =   3855
         Begin VB.OptionButton OptOpcion2 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   16
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptOpcion2 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   5295
      End
      Begin VB.OptionButton OptOpcion1 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   32
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptOpcion1 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   4680
      Top             =   7920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.Data DataPedidos 
      Caption         =   "Pedidos De Materia Prima a Proveedores"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "PedidosProveedores"
      Top             =   7320
      Width           =   11655
   End
   Begin TabDlg.SSTab TabPedidos 
      Height          =   7095
      Left            =   120
      TabIndex        =   35
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12515
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePedidos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Busqueda"
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).ControlCount=   1
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
         Height          =   6135
         Left            =   -74880
         TabIndex        =   41
         Top             =   720
         Width           =   11295
         Begin VB.Frame FrameTipoBusqueda2 
            Caption         =   "Tipo De Busqueda"
            Height          =   1575
            Left            =   7320
            TabIndex        =   19
            Top             =   360
            Visible         =   0   'False
            Width           =   3495
            Begin VB.OptionButton OptTipBus2 
               Caption         =   "Palabra Inicial"
               Height          =   1095
               Index           =   1
               Left            =   1800
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton OptTipBus2 
               Caption         =   "Palabra Inicial"
               Height          =   1095
               Index           =   0
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   27
               Top             =   360
               Value           =   -1  'True
               Width           =   1575
            End
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas De Entrega De Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Index           =   3
            Left            =   4680
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   480
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DTPFecFin 
            Height          =   255
            Left            =   8040
            TabIndex        =   30
            Top             =   3240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24772611
            CurrentDate     =   37281
         End
         Begin MSComCtl2.DTPicker DTPFecIni 
            Height          =   255
            Left            =   8040
            TabIndex        =   29
            Top             =   2760
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24772611
            CurrentDate     =   37281
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas De Pedido Y Proveedor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Index           =   2
            Left            =   3240
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   480
            Width           =   1335
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8040
            TabIndex        =   31
            ToolTipText     =   "Signo '+' O Doble Click Para Ver Lista"
            Top             =   3840
            Width           =   2535
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas De Pedido Y Materia Prima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Index           =   1
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas De Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Index           =   0
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   480
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label lblFieldLabel 
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
            Height          =   195
            Index           =   13
            Left            =   6840
            TabIndex        =   52
            Top             =   2760
            Width           =   1110
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   12
            Left            =   6960
            TabIndex        =   51
            Top             =   3240
            Width           =   1005
         End
         Begin VB.Label LblBusqueda 
            Alignment       =   1  'Right Justify
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
            Left            =   5880
            TabIndex        =   42
            Top             =   3840
            Width           =   2055
         End
      End
      Begin VB.Frame FramePedidos 
         Caption         =   "Datos del Pedido"
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
         Height          =   6015
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   11175
         Begin MSMask.MaskEdBox MskSalEnt 
            DataField       =   "SaldoPorEntregar"
            DataSource      =   "DataPedidos"
            Height          =   285
            Left            =   7920
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   5520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanEnt 
            DataField       =   "CantidadEntregada"
            DataSource      =   "DataPedidos"
            Height          =   285
            Left            =   3360
            TabIndex        =   8
            Top             =   5520
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "DiasDeAtraso"
            DataSource      =   "DataPedidos"
            Enabled         =   0   'False
            Height          =   285
            Index           =   2
            Left            =   7920
            MaxLength       =   10
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   4800
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MskCanPed 
            DataField       =   "CantidadPedido"
            DataSource      =   "DataPedidos"
            Height          =   285
            Left            =   1920
            TabIndex        =   2
            Top             =   2280
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "DiasPedido"
            DataSource      =   "DataPedidos"
            Enabled         =   0   'False
            Height          =   285
            Index           =   7
            Left            =   3360
            MaxLength       =   10
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   4800
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "UsuarioEditar"
            DataSource      =   "DataPedidos"
            Height          =   285
            Index           =   10
            Left            =   8160
            MaxLength       =   50
            TabIndex        =   13
            Top             =   4320
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "UsuarioAgregar"
            DataSource      =   "DataPedidos"
            Height          =   285
            Index           =   9
            Left            =   6600
            MaxLength       =   50
            TabIndex        =   12
            Top             =   4320
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataPedidos"
            Height          =   285
            Index           =   8
            Left            =   1920
            MaxLength       =   50
            TabIndex        =   5
            Top             =   3360
            Width           =   7080
         End
         Begin MSMask.MaskEdBox MskFecEntTot 
            DataField       =   "FechaEntregaTotal"
            DataSource      =   "DataPedidos"
            Height          =   285
            Left            =   7920
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   5160
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecEnt 
            DataField       =   "FechaParaEntregar"
            DataSource      =   "DataPedidos"
            Height          =   285
            Left            =   3360
            TabIndex        =   7
            Top             =   5160
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecPed 
            DataField       =   "FechaPedido"
            DataSource      =   "DataPedidos"
            Height          =   285
            Left            =   1920
            TabIndex        =   0
            Top             =   1560
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Codigo"
            DataSource      =   "DataPedidos"
            Height          =   285
            Index           =   5
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   3
            ToolTipText     =   "Signo '+' O Doble Click Para Ver Lista"
            Top             =   2640
            Width           =   1260
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Proveedor"
            DataSource      =   "DataPedidos"
            Height          =   285
            Index           =   6
            Left            =   1920
            MaxLength       =   10
            TabIndex        =   4
            Top             =   3000
            Width           =   1260
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataPedidos"
            Height          =   285
            Index           =   0
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   1
            Top             =   1920
            Width           =   1260
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   3240
            Top             =   1680
            Width           =   480
         End
         Begin VB.Label LblUniMed 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
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
            TabIndex        =   20
            Top             =   2280
            Width           =   5775
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "Dias De Atraso"
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
            Left            =   5880
            TabIndex        =   21
            Top             =   4800
            Width           =   1290
         End
         Begin VB.Label LblMateriaPrima 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
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
            TabIndex        =   50
            Top             =   2640
            Width           =   5775
         End
         Begin VB.Label LblProveedor 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label1"
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
            TabIndex        =   49
            Top             =   3000
            Width           =   5775
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   48
            Top             =   3360
            Width           =   1065
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "Fecha De Entregado"
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
            Left            =   5880
            TabIndex        =   47
            Top             =   5160
            Width           =   1770
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "Fecha Para Entregar"
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
            Left            =   840
            TabIndex        =   46
            Top             =   5160
            Width           =   1770
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Pedido"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   45
            Top             =   1560
            Width           =   990
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "Dias De Entrega"
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
            Left            =   840
            TabIndex        =   44
            Top             =   4800
            Width           =   1410
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "Saldo Por Entregar"
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
            Left            =   5880
            TabIndex        =   43
            Top             =   5520
            Width           =   1620
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   40
            Top             =   3000
            Width           =   885
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H00008000&
            Caption         =   "Cantidad Entregada"
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
            Left            =   840
            TabIndex        =   39
            Top             =   5520
            Width           =   1695
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Materia Prima"
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
            Left            =   240
            TabIndex        =   38
            Top             =   2640
            Width           =   1170
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Pedido"
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
            Index           =   2
            Left            =   240
            TabIndex        =   37
            Top             =   2280
            Width           =   1410
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
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
            Left            =   240
            TabIndex        =   36
            Top             =   1920
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00008000&
            BackStyle       =   1  'Opaque
            Height          =   1215
            Left            =   120
            Top             =   4680
            Width           =   10935
         End
      End
   End
End
Attribute VB_Name = "PedidosProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim Bandera As Boolean
Dim VMensaje As Integer

Dim RBuscaBodega As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaProveedor As Recordset
Dim RBuscaProveedorDias As Recordset
Dim RBuscaMaximo As Recordset


Dim BBodega As Boolean
Dim BMateriaPrima As Boolean
Dim BProveedor As Boolean
Dim BMateriaPrima2 As Boolean


Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next

    With DataPedidos.Recordset
    
        'AGREGAR
        If Index = 0 Then
                .AddNew
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                MskFecPed.SetFocus
                MskFecPed.Text = Date
                'USUARIO AGREGAR
                TxtTexto.Item(9).Text = GUsuario
                
                'BUSCA EL DOCUMENTO MAXIMO QUE EXISTE DE DOCUMENTOS Y LE SUMA UNO
                'Set RBuscaMaximo = Db.OpenRecordset("Select Max(documento) From PedidosMateriaPrimaProveedores")
                'If IsNull(RBuscaMaximo(0)) Then
                '        TxtTexto.Item(0).Text = "1"
                'Else
                '        TxtTexto.Item(0).Text = RBuscaMaximo(0) + 1
                'End If
        'EDITAR
        ElseIf Index = 1 Then
                        .Edit
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                MskFecPed.SetFocus
                'USUARIO EDITAR
                TxtTexto.Item(10).Text = GUsuario
        'GRABAR
        ElseIf Index = 2 Then
        
                        'REVISA LA FECHA
                        If Not IsDate(MskFecPed.Text) Then
                                MsgBox "Fecha De Pedido Incorrecta", vbOKOnly + vbInformation, "Informacion"
                                Exit Sub
                        End If
                                            
                        'REVISA SI EXISTE LA MATERIA PRIMA
                        Set RBuscaMateriaPrima = Db.OpenRecordset("Select * From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(5).Text & "'")
                        If RBuscaMateriaPrima.RecordCount > 0 Then
                        Else
                            MsgBox "El Codigo Materia Prima No Existe En Correlativos Maximos De Materia Prima", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                                  
                        'SALDO DE PEDIDO ES IGUAL A CANTIDAD DE PEDIDO MENOS CANTIDAD ENTREGADA
                        MskSalEnt.Text = Val(MskCanPed.Text) - Val(MskCanEnt.Text)
                        
                        'GRABA DATOS
                        .Update
                        If Err.Number <> 0 And Err.Number <> 3022 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        ElseIf Err.Number = 3022 Then
                                MsgBox "Este Pedido Con Este Codigo De Materia Prima Ya Existe ", vbInformation, "Informacion"
                                Exit Sub
                        End If
                Bandera = False
                botones
                CmdBotones.Item(0).SetFocus
        'CANCELAR
        ElseIf Index = 3 Then
                .CancelUpdate
                        If Err.Number <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = False
                botones
        ElseIf Index = 4 Then ' BORRAR
                VMensaje = MsgBox("Esta seguro de borrar el registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                If vbYes Then
                    .Delete
                    .MoveLast
                            If Err.Number <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                End If
        ElseIf Index = 5 Then 'BUSCAR
                    mensaje = InputBox("Pedido a Buscar")
                    .FindFirst ("Documento = '" & mensaje & "'")
        ElseIf Index = 6 Then 'IMPRIMIR
                MousePointer = 11
                    CrReportes.SelectionFormula = "{PedidosMateriaPrimaProveedores.Documento} = '" & TxtTexto.Item(0) & "' And {PedidosMateriaPrimaProveedores.Codigo} = '" & TxtTexto.Item(5) & "'"
                    CrReportes.ReportFileName = App.Path & "\PedidosMateriaPrimaProveedores.rpt"
                    CrReportes.Action = 1
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    End If
                MousePointer = 0
        ElseIf Index = 7 Then ' SALIDA
                Unload Me
        ElseIf Index = 8 Then 'SELECCIONAR DATOS
                    'FECHAS DE PEDIDO
                    If OptBusqueda.Item(0).Value = True Then
                        DataPedidos.RecordSource = ("Select * From PedidosMateriaPrimaProveedores where FechaPedido >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And FechaPedido <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order By FechaPedido")
                    'FECHAS DE PEDIDO Y MATERIA PRIMA
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        If OptTipBus2.Item(0).Value = True Then
                            DataPedidos.RecordSource = ("Select * From PedidosMateriaPrimaProveedores where FechaPedido >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And FechaPedido <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Codigo Like '" & TxtBusqueda.Text & "*' Order By FechaPedido")
                        Else
                            DataPedidos.RecordSource = ("Select * From PedidosMateriaPrimaProveedores where FechaPedido >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And FechaPedido <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Codigo Like '*" & TxtBusqueda.Text & "*' Order By FechaPedido")
                        End If
                    'FECHAS DE PEDIDO Y PROVEEDOR
                    ElseIf OptBusqueda.Item(2).Value = True Then
                        If OptTipBus2.Item(0).Value = True Then
                            DataPedidos.RecordSource = ("Select * From PedidosMateriaPrimaProveedores where FechaPedido >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And FechaPedido <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Proveedor Like '" & TxtBusqueda.Text & "*' Order By FechaPedido")
                        Else
                            DataPedidos.RecordSource = ("Select * From PedidosMateriaPrimaProveedores where FechaPedido >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And FechaPedido <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Proveedor Like '*" & TxtBusqueda.Text & "*' Order By FechaPedido")
                        End If
                    'FECHAS DE ENTREGA
                    ElseIf OptBusqueda.Item(3).Value = True Then
                        DataPedidos.RecordSource = ("Select * From PedidosMateriaPrimaProveedores where FechaParaEntregar >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# And FechaParaEntregar <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order By FechaPedido")
                    End If
                    DataPedidos.Refresh
                    DGridPedidos.Refresh
                    TabPedidos.Tab = 1
        ElseIf Index = 9 Then 'ACTUALIZAR
                    DataPedidos.RecordSource = "Select * From PedidosMateriaPrimaProveedores Order By FechaPedido"
                    DataPedidos.Refresh
                    DGridPedidos.Refresh
                    TabPedidos.Tab = 1
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
         CmdBotones.Item(6).Enabled = False
         CmdBotones.Item(7).Enabled = False
         FramePedidos.Enabled = True
         DataPedidos.Visible = False
         DGridPedidos.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         CmdBotones.Item(6).Enabled = True
         CmdBotones.Item(7).Enabled = True
         FramePedidos.Enabled = False
         DataPedidos.Visible = True
         DGridPedidos.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub


Private Sub DBGridBusqueda_DblClick()
        'BODEGAS
        If BBodega = True Then
                TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0)
                TxtTexto.Item(1).SetFocus
        'MATERIA PRIMA
        ElseIf BMateriaPrima = True Then
                TxtTexto.Item(5).Text = DBGridBusqueda.Columns(0)
                TxtTexto.Item(5).SetFocus
        'PROVEEDOR
        ElseIf BProveedor = True Then
                TxtTexto.Item(6).Text = DBGridBusqueda.Columns(0)
                TxtTexto.Item(6).SetFocus
        'MATERIA PRIMA DE BUSQUEDA
        ElseIf BMateriaPrima2 = True Then
                TxtBusqueda.Text = DBGridBusqueda.Columns(0)
                TxtBusqueda.SetFocus
        End If
        FrameBusqueda.Visible = False
        TxtBuscar.Text = ""
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        
        If KeyAscii = 43 Then
            'BODEGAS
            If BBodega = True Then
                    TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0)
                    TxtTexto.Item(1).SetFocus
            'MATERIA PRIMA
            ElseIf BMateriaPrima = True Then
                    TxtTexto.Item(5).Text = DBGridBusqueda.Columns(0)
                    TxtTexto.Item(5).SetFocus
            'PROVEEDOR
            ElseIf BProveedor = True Then
                    TxtTexto.Item(6).Text = DBGridBusqueda.Columns(0)
                    TxtTexto.Item(6).SetFocus
            'MATERIA PRIMA DE BUSQUEDA
            ElseIf BMateriaPrima2 = True Then
                    TxtBusqueda.Text = DBGridBusqueda.Columns(0)
                    TxtBusqueda.SetFocus
            End If
            FrameBusqueda.Visible = False
            TxtBuscar.Text = ""
        End If
End Sub

Private Sub DGridPedidos_HeadClick(ByVal ColIndex As Integer)
    DataPedidos.RecordSource = "Select * from PedidosMateriaPrimaProveedores Order by " & DGridPedidos.Columns(ColIndex).DataField
    DataPedidos.Refresh
    DGridPedidos.Refresh
End Sub

Private Sub Form_Activate()
    'MATERIA PRIMA
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(5).Text & "'")
            If RBuscaMateriaPrima.RecordCount > 0 Then
                LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                If IsNull(RBuscaMateriaPrima!UnidadMedida) Then
                    LblUniMed.Caption = ""
                Else
                    LblUniMed.Caption = RBuscaMateriaPrima!UnidadMedida
                End If
            Else
                LblMateriaPrima.Caption = ""
                LblUniMed.Caption = ""
            End If
    'PROVEEDOR
    
            Set RBuscaProveedor = Db.OpenRecordset("Select Proveedor From Proveedores Where CodigoProveedor = '" & TxtTexto.Item(6).Text & "'")
            If RBuscaProveedor.RecordCount > 0 Then
                LblProveedor.Caption = RBuscaProveedor!Proveedor
            Else
                LblProveedor.Caption = ""
            End If
    

End Sub

Private Sub Form_Load()
    DataPedidos.Connect = GConnect
    DataBusqueda.Connect = GConnect
    
    DataPedidos.DatabaseName = BasedeDatos
    DataBusqueda.DatabaseName = BasedeDatos
End Sub

Private Sub MskCanEnt_GotFocus()
        MskCanEnt.SelStart = 0
        MskCanEnt.SelLength = Len(MskCanEnt.Text)
End Sub

Private Sub MskCanEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
        'SALDO DE PEDIDO ES IGUAL A CANTIDAD DE PEDIDO MENOS CANTIDAD ENTREGADA
        MskSalEnt.Text = Val(MskCanPed.Text) - Val(MskCanEnt.Text)
    End If
End Sub

Private Sub MskCanPed_GotFocus()
    MskCanPed.SelStart = 0
    MskCanPed.SelLength = Len(MskCanPed.Text)
End Sub

Private Sub MskCanPed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskCanPed_LostFocus()
            'SALDO DE PEDIDO ES IGUAL A CANTIDAD DE PEDIDO MENOS CANTIDAD ENTREGADA
            MskSalEnt.Text = Val(MskCanPed.Text) - Val(MskCanEnt.Text)
End Sub

Private Sub MskFecEnt_GotFocus()
        MskFecEnt.SelStart = 0
        MskFecEnt.SelLength = Len(MskFecEnt.Text)
End Sub

Private Sub MskFecEnt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
             If IsDate(MskFecEnt.Text) Then
                TxtTexto.Item(7).Text = DateValue(MskFecEnt.Text) - DateValue(MskFecPed.Text)
             End If
        End If
End Sub

Private Sub MskFecEnt_LostFocus()
             If IsDate(MskFecEnt.Text) Then
                TxtTexto.Item(7).Text = DateValue(MskFecEnt.Text) - DateValue(MskFecPed.Text)
             End If
End Sub

Private Sub MskFecPed_GotFocus()
        MskFecPed.SelStart = 0
        MskFecPed.SelLength = Len(MskFecPed.Text)
End Sub

Private Sub MskFecPed_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
                SendKeys "{tab}"
        End If
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            FrameTipoBusqueda2.Visible = False
            LblBusqueda.Caption = ""
            TxtBusqueda.Visible = False
    ElseIf Index = 1 Then
            FrameTipoBusqueda2.Visible = True
            LblBusqueda.Caption = "Codigo Materia Prima"
            TxtBusqueda.Visible = True
            TxtBusqueda.SetFocus
    ElseIf Index = 2 Then
            FrameTipoBusqueda2.Visible = True
            LblBusqueda.Caption = "Codigo Proveedor"
            TxtBusqueda.Visible = True
            TxtBusqueda.SetFocus
    ElseIf Index = 3 Then
        FrameTipoBusqueda2.Visible = False
            LblBusqueda.Caption = ""
            TxtBusqueda.Visible = False
    End If
                
End Sub

Private Sub TabPedidos_Click(PreviousTab As Integer)
    If TabPedidos.Tab = 2 Then
        DTPFecIni.Value = Date
        DTPFecFin.Value = Date
        
    End If
End Sub

Private Sub TxtBuscar_Change()
    'BODEGA
    If BBodega = True Then
            'DESCRIPCION
            If OptOpcion1.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptOpcion2.Item(0).Value = True Then
                    DataBusqueda.RecordSource = ("Select * From BodegasMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*'")
                'CUALQUIER PALABRA
                ElseIf OptOpcion2.Item(1).Value = True Then
                    DataBusqueda.RecordSource = ("Select * From BodegasMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*'")
                End If
            'CODIGO
            ElseIf OptOpcion1.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptOpcion2.Item(0).Value = True Then
                    DataBusqueda.RecordSource = ("Select * From BodegasMateriaPrima Where CodigoBodega Like '" & TxtBuscar.Text & "*'")
                'CUALQUIER PALABRA
                ElseIf OptOpcion2.Item(1).Value = True Then
                    DataBusqueda.RecordSource = ("Select * From BodegasMateriaPrima Where CodigoBodega Like '*" & TxtBuscar.Text & "*'")
                End If
            End If
            
    'MATERIA PRIMA
    ElseIf (BMateriaPrima = True Or BMateriaPrima2 = True) Then
            'DESCRIPCION
            If OptOpcion1.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptOpcion2.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptOpcion2.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*'"
                End If
            'CODIGO
            ElseIf OptOpcion1.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptOpcion2.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '" & TxtBuscar.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptOpcion2.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & TxtBuscar.Text & "*'"
                End If
            End If
    End If
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
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

Private Sub TxtBusqueda_DblClick()
        'MATERIA PRIMA
        If OptBusqueda.Item(1).Value = True Then
            BBodega = False
            BMateriaPrima = False
            BProveedor = False
            BMateriaPrima2 = True
            DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
        'PROVEEDORES
        ElseIf OptBusqueda.Item(2).Value = True Then
            BBodega = False
            BMateriaPrima = False
            BProveedor = False
            BMateriaPrima2 = False
            DataBusqueda.RecordSource = "Select CodigoProveedor, Proveedor, TipoDeProveedor From Proveedores"
        End If
            FrameBusqueda.Visible = True
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
            TxtBuscar.SetFocus
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        'MATERIA PRIMA
        If OptBusqueda.Item(1).Value = True Then
            BBodega = False
            BMateriaPrima = False
            BProveedor = False
            BMateriaPrima2 = True
            DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
        'PROVEEDORES
        ElseIf OptBusqueda.Item(2).Value = True Then
            BBodega = False
            BMateriaPrima = False
            BProveedor = False
            BMateriaPrima2 = True
            DataBusqueda.RecordSource = "Select CodigoProveedor, Proveedor, TipoDeProveedor From Proveedores"
        End If
            FrameBusqueda.Visible = True
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
            TxtBuscar.SetFocus
    End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
    
    'MATERIA PRIMA
    If Index = 5 Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(5).Text & "'")
            If RBuscaMateriaPrima.RecordCount > 0 Then
                LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                LblUniMed.Caption = RBuscaMateriaPrima!UnidadMedida
            Else
                LblMateriaPrima.Caption = ""
                LblUniMed.Caption = ""
            End If
    'PROVEEDOR
    ElseIf Index = 6 Then
            Set RBuscaProveedor = Db.OpenRecordset("Select Proveedor From Proveedores Where CodigoProveedor = '" & TxtTexto.Item(6).Text & "'")
            If RBuscaProveedor.RecordCount > 0 Then
                LblProveedor.Caption = RBuscaProveedor!Proveedor
            Else
                LblProveedor.Caption = ""
            End If
    End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    'BODEGAS
    If Index = 1 Then
            BBodega = True
            BMateriaPrima = False
            BProveedor = False
            BMateriaPrima2 = False
            DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima"
    'MATERIA PRIMA
    ElseIf Index = 5 Then
            BBodega = False
            BMateriaPrima = True
            BProveedor = False
            BMateriaPrima2 = False
            DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
    'PROVEEDORES
    ElseIf Index = 6 Then
            BBodega = False
            BMateriaPrima = False
            BProveedor = True
            BMateriaPrima2 = False
            DataBusqueda.RecordSource = "Select CodigoProveedor, Proveedor From Proveedores"
    End If
    'SI EL INDICE ES DIFERENTE A ESTOS NO HACE NADA
    If (Index = 1 Or Index = 5 Or Index = 6) Then
            FrameBusqueda.Visible = True
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"
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
            'BODEGAS
            If Index = 1 Then
                    BBodega = True
                    BMateriaPrima = False
                    BProveedor = False
                    BMateriaPrima2 = False
                    DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima"
            'MATERIA PRIMA
            ElseIf Index = 5 Then
                    BBodega = False
                    BMateriaPrima = True
                    BProveedor = False
                    BMateriaPrima2 = False
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
            'PROVEEDORES
            ElseIf Index = 6 Then
                    BBodega = False
                    BMateriaPrima = False
                    BProveedor = True
                    BMateriaPrima2 = False
                    DataBusqueda.RecordSource = "Select CodigoProveedor, Proveedor From Proveedores"
            End If
            
            'SI EL INDICE ES DIFERENTE A ESTOS NO HACE NADA
            If (Index = 1 Or Index = 5 Or Index = 6) Then
                    FrameBusqueda.Visible = True
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "4000"
                    TxtBuscar.SetFocus
            End If
            
        End If
End Sub

Private Sub TxtTexto_LostFocus(Index As Integer)
    'MATERIA PRIMA
    If Index = 5 Then
        'BUSCA LOS DIAS DE ENTREGA DEL PROVEEDOR
        'Set RBuscaProveedorDias = Db.OpenRecordset("Select DiasDeEntrega From Proveedores where CodigoProveedor = '" & TxtTexto.Item(6).Text & "'")
        'If RBuscaProveedorDias.RecordCount Then
        '    TxtTexto.Item(7).Text = RBuscaProveedorDias!DiasDeEntrega
        '    If Not IsDate(MskFecPed.Text) Then
        '            MsgBox "Ingrese La Fecha De Pedido Para Calcular La Fecha De Entrega", vbOKOnly + vbInformation, "Informacion"
        '    Else
        '            MskFecEnt.Text = DateValue(MskFecPed.Text) + TxtTexto.Item(7).Text
        '    End If
        'Else
        '    MskFecEnt.Text = MskFecPed.Text
        '    TxtTexto.Item(7).Text = "0"
        'End If
        
        'BUSCA MATERIA PRIMA
        Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(5).Text & "'")
            If RBuscaMateriaPrima.RecordCount > 0 Then
                LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                LblUniMed.Caption = RBuscaMateriaPrima!UnidadMedida
            Else
                LblMateriaPrima.Caption = ""
                LblUniMed.Caption = ""
            End If

    End If
    
    
End Sub
