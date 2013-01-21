VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form CerrarPedidoClientes 
   BackColor       =   &H00008000&
   Caption         =   "Cerrar Pedidos De Clientes"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "CerrarPedidoClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda De Datos"
      Height          =   8415
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   11655
      Begin VB.OptionButton OptOpcion1 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptOpcion1 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   3975
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         Height          =   735
         Left            =   4200
         TabIndex        =   27
         Top             =   240
         Width           =   3855
         Begin VB.OptionButton OptOpcion2 
            Caption         =   "Palabra Inicial"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptOpcion2 
            Caption         =   "Cualquier Palabra"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   1
            Left            =   1920
            TabIndex        =   29
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   8160
         Picture         =   "CerrarPedidoClientes.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "sale de busqueda"
         Top             =   360
         Width           =   855
      End
      Begin VB.Data DataBusqueda 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1680
         Visible         =   0   'False
         Width           =   2415
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "CerrarPedidoClientes.frx":293C
         Height          =   7215
         Left            =   120
         OleObjectBlob   =   "CerrarPedidoClientes.frx":2957
         TabIndex        =   26
         ToolTipText     =   "doble click o signo '+' para seleccionar"
         Top             =   1080
         Width           =   11415
      End
   End
   Begin VB.Data DataCerrarPedidos 
      Caption         =   "Cerrar Pedido De Materia Prima a Proveedores"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CerrarPedidoClientes"
      Top             =   8040
      Width           =   11775
   End
   Begin TabDlg.SSTab TabDepartamentos 
      Height          =   7935
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13996
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CerrarPedidoClientes.frx":3331
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "CmdBotones(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdBotones(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdBotones(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdBotones(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdBotones(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FrameCerrarPedidos"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CerrarPedidoClientes.frx":364B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridCerrarPedidos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "CerrarPedidoClientes.frx":3A9D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).ControlCount=   1
      Begin MSDBGrid.DBGrid DGridCerrarPedidos 
         Bindings        =   "CerrarPedidoClientes.frx":3EEF
         Height          =   7095
         Left            =   -74880
         OleObjectBlob   =   "CerrarPedidoClientes.frx":3F0F
         TabIndex        =   12
         ToolTipText     =   "click en encabezado columna para indexar"
         Top             =   720
         Width           =   11535
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
         Height          =   6975
         Left            =   -74880
         TabIndex        =   13
         Top             =   720
         Width           =   11535
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "No. Pedido"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   2
            Left            =   3000
            Picture         =   "CerrarPedidoClientes.frx":5325
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   360
            Width           =   1335
         End
         Begin MSComCtl2.DTPicker DtpFecFin 
            Height          =   255
            Left            =   7320
            TabIndex        =   18
            Top             =   3480
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37501
         End
         Begin MSComCtl2.DTPicker DtpFecIni 
            Height          =   255
            Left            =   7320
            TabIndex        =   17
            Top             =   3120
            Visible         =   0   'False
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   24707075
            CurrentDate     =   37501
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            TabIndex        =   19
            Top             =   4080
            Width           =   2535
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fechas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Index           =   1
            Left            =   1560
            Picture         =   "CerrarPedidoClientes.frx":562F
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptBusqueda 
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
            Height          =   975
            Index           =   0
            Left            =   120
            Picture         =   "CerrarPedidoClientes.frx":84A9
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   6
            Left            =   8760
            TabIndex        =   20
            Top             =   5280
            Width           =   2535
            Caption         =   "Seleccionar Datos"
            PicturePosition =   196613
            Size            =   "4471;1085"
            Picture         =   "CerrarPedidoClientes.frx":DAAB
            FontEffects     =   1073741825
            FontHeight      =   165
            FontCharSet     =   0
            FontPitchAndFamily=   2
            ParagraphAlign  =   3
            FontWeight      =   700
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
            Height          =   195
            Left            =   6120
            TabIndex        =   33
            Top             =   3120
            Visible         =   0   'False
            Width           =   1110
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
            Left            =   6120
            TabIndex        =   34
            Top             =   3480
            Visible         =   0   'False
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
            Left            =   3960
            TabIndex        =   40
            Top             =   4080
            Width           =   2055
         End
         Begin MSForms.CommandButton CmdBotones 
            Height          =   615
            Index           =   7
            Left            =   8760
            TabIndex        =   21
            Top             =   6120
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
      End
      Begin VB.Frame FrameCerrarPedidos 
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
         Height          =   6375
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   11535
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "TipoDocumento"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   5
            Left            =   5400
            MaxLength       =   20
            TabIndex        =   2
            ToolTipText     =   "20 digitos"
            Top             =   360
            Width           =   2340
         End
         Begin VB.TextBox TxtDatos2 
            BackColor       =   &H0080C0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1935
            Left            =   360
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   44
            Top             =   4320
            Width           =   10815
         End
         Begin VB.TextBox TxtDatos 
            BackColor       =   &H0000C000&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   360
            MultiLine       =   -1  'True
            TabIndex        =   43
            Top             =   2280
            Width           =   10815
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "FechaOperacion"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   3
            Left            =   9240
            TabIndex        =   31
            Top             =   360
            Visible         =   0   'False
            Width           =   960
         End
         Begin MSMask.MaskEdBox MskCan 
            DataField       =   "Cantidad"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Left            =   2280
            TabIndex        =   6
            Top             =   1800
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   49152
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecRec 
            DataField       =   "FechaRecepcion"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Left            =   2280
            TabIndex        =   3
            Top             =   720
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "UsuarioAgregar"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   4
            Left            =   10320
            MaxLength       =   10
            TabIndex        =   32
            Top             =   360
            Visible         =   0   'False
            Width           =   840
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Pedido"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   2
            Left            =   2280
            MaxLength       =   12
            TabIndex        =   5
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   1440
            Width           =   1500
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Codigo"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   1
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   4
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   1095
            Width           =   1500
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            DataField       =   "Documento"
            DataSource      =   "DataCerrarPedidos"
            Height          =   285
            Index           =   0
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   1
            Top             =   360
            Width           =   1500
         End
         Begin VB.Label lblFieldLabel 
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
            Index           =   7
            Left            =   3840
            TabIndex        =   47
            Top             =   360
            Width           =   1410
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Datos Del Pedido"
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
            Left            =   9720
            TabIndex        =   46
            Top             =   1920
            Width           =   1500
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Datos De Cierres De Pedido"
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
            Left            =   8760
            TabIndex        =   45
            Top             =   4080
            Width           =   2400
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
            Left            =   3840
            TabIndex        =   42
            Top             =   1080
            Width           =   7335
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Despacho"
            Height          =   195
            Index           =   4
            Left            =   480
            TabIndex        =   41
            Top             =   720
            Width           =   1230
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Pedido"
            Height          =   195
            Index           =   3
            Left            =   480
            TabIndex        =   39
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad Enviada"
            Height          =   195
            Index           =   2
            Left            =   480
            TabIndex        =   38
            Top             =   1800
            Width           =   1260
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Codigo "
            Height          =   195
            Index           =   1
            Left            =   480
            TabIndex        =   37
            Top             =   1080
            Width           =   540
         End
         Begin VB.Label lblFieldLabel 
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
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   36
            Top             =   360
            Width           =   1335
         End
      End
      Begin MSForms.CommandButton CmdBotones 
         Height          =   615
         Index           =   5
         Left            =   9360
         TabIndex        =   11
         Top             =   7200
         Width           =   2205
         Caption         =   "Salida"
         PicturePosition =   196613
         Size            =   "3881;1085"
         Picture         =   "CerrarPedidoClientes.frx":DEFD
         Accelerator     =   83
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdBotones 
         Height          =   615
         Index           =   4
         Left            =   7080
         TabIndex        =   10
         Top             =   7200
         Width           =   2205
         Caption         =   "Borrar"
         PicturePosition =   196613
         Size            =   "3881;1085"
         Picture         =   "CerrarPedidoClientes.frx":FF7F
         Accelerator     =   66
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdBotones 
         Height          =   615
         Index           =   3
         Left            =   4800
         TabIndex        =   9
         Top             =   7200
         Width           =   2205
         VariousPropertyBits=   25
         Caption         =   "Cancelar"
         PicturePosition =   196613
         Size            =   "3881;1085"
         Picture         =   "CerrarPedidoClientes.frx":104C1
         Accelerator     =   67
         FontEffects     =   1073750017
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdBotones 
         Height          =   615
         Index           =   2
         Left            =   2520
         TabIndex        =   8
         Top             =   7200
         Width           =   2205
         VariousPropertyBits=   25
         Caption         =   "Grabar"
         PicturePosition =   196613
         Size            =   "3881;1085"
         Accelerator     =   71
         FontEffects     =   1073750017
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdBotones 
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   7200
         Width           =   2205
         Caption         =   "Agregar"
         PicturePosition =   196613
         Size            =   "3881;1085"
         Picture         =   "CerrarPedidoClientes.frx":10A03
         Accelerator     =   65
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "CerrarPedidoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer

Dim VFechaEntrada As Date
Dim VDiasDeAtraso As Long
Dim VNumeroPedido As String
Dim VCodigo As String
Dim VCantidadEntrada As Single
Dim VCantidadCerrarPedido As Single

Dim RBuscaPedido As Recordset
Dim RBuscaCantidadPedido As Recordset
Dim RBuscaPedido2 As Recordset
Dim RBuscaPedido3 As Recordset
Dim RBuscaRecepcion As Recordset
Dim RBuscaCantidadRecepcion As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaSaldoEncabezado As Recordset
Dim RBuscaSaldoDetalle As Recordset
Dim RBuscaCierrePedidos As Recordset

Dim BMateriaPrima As Boolean
Dim BPedido As Boolean



Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    With DataCerrarPedidos.Recordset
    
        'AGREGAR
        If Index = 0 Then
                .AddNew
                        If Err.Number > 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                Bandera = True
                botones
                MskFecRec.Text = Date
                TxtTexto.Item(4).Text = GUsuario
                TxtTexto.Item(3).Text = Date
                TxtTexto.Item(0).SetFocus
                
        'GRABAR
        ElseIf Index = 2 Then
        
                'REVISA SI ES NUMERICO LA CANTIDAD DE ENTRADA
                If Not IsNumeric(MskCan.Text) Then
                        MsgBox "Cantidad De Entrada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
                
                'REVISA SI ES FECHA VALIDA
                If Not IsDate(MskFecRec.Text) Then
                        MsgBox "Fecha De Despacho Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
                
                
                'REVISA SI EXISTE EL PEDIDO
                Set RBuscaPedido2 = Db.OpenRecordset("Select * From EncabezadoPedidosClientes Where Documento = '" & TxtTexto.Item(2).Text & "'")
                If RBuscaPedido2.RecordCount > 0 Then
                Else
                         MsgBox "Pedido No Existe", vbOKOnly + vbInformation, "Informacion"
                End If
                
                'REVISA SI EXISTE EL PEDIDO Y CON LA MATERIA PRIMA
                Set RBuscaPedido3 = Db.OpenRecordset("Select * From DetallePedidosClientes Where Documento = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaPedido3.RecordCount > 0 Then
                Else
                         MsgBox "Este Pedido No Corresponde A Esta Materia Prima", vbOKOnly + vbInformation, "Informacion"
                End If
                
                'BUSCA EL SALDO DEL PEDIDO
                Set RBuscaSaldoEncabezado = Db.OpenRecordset("Select SaldoPorEntregar From DetallePedidosClientes Where Documento = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaSaldoEncabezado.RecordCount > 0 Then
                    If Val(MskCan.Text) > Val(RBuscaSaldoEncabezado!SaldoPorEntregar) Then
                        MsgBox "La Cantidad No Puede Ser Mayor Al Saldo Del Pedido", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                End If
                
                'BUSCA SI EL CODIGO PERTENECE A LA RECEPCION
                'Set RBuscaRecepcion = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where Documento = '" & TxtTexto.Item(0).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                '    If RBuscaRecepcion.RecordCount > 0 Then
                '    Else
                '        MsgBox "El Codigo Materia Prima No Ingreso Por Este Numero De Recepcion " & TxtTexto.Item(0).Text, vbOKOnly + vbExclamation, "Advertencia"
                        'Txttexto.Item(1).SetFocus
                        'Exit Sub
                '    End If
                                    
                'REVISA SI LA CANTIDAD QUE ESTAN INGRESANDO NO ES MAYOR QUE LE SALDO QUE LE FALTA A LA RECEPCION DEPENDIENDO DEL CODIGO
                '________________________________________________________________________________________________________
                            'BUSCA LA CANTIDAD QUE TRAE LA RECEPCION
                            'Set RBuscaRecepcion = Db.OpenRecordset("Select Sum(Cantidad) From DetalleEntradasMateriaPrima Where Documento = '" & TxtTexto.Item(0).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                            '    If RBuscaRecepcion.RecordCount > 0 Then
                            '        If IsNull(RBuscaRecepcion(0)) Then
                            '            VCantidadEntrada = 0
                            '        Else
                            '            VCantidadEntrada = RBuscaRecepcion(0)
                            '        End If
                            '    Else
                            '        VCantidadEntrada = 0
                            '    End If
                           '
                            'BUSCA LA CANTIDAD QUE TIENE EL CIERRE DE PEDIDOS
                            'Set RBuscaPedido = Db.OpenRecordset("Select Sum(Cantidad) From CerrarPedidoClientes Where Documento = '" & TxtTexto.Item(0).Text & "' And CodigoMateriaPrima = '" & TxtTexto.Item(1).Text & "'")
                            '    If RBuscaPedido.RecordCount > 0 Then
                            '        If IsNull(RBuscaPedido(0)) Then
                            '            VCantidadCerrarPedido = 0
                            '        Else
                            '            VCantidadCerrarPedido = RBuscaPedido(0)
                            '        End If
                            '    Else
                            '        VCantidadCerrarPedido = 0
                            '    End If
                            'SI LA CANTIDAD INGRESADA ES MAYOR AL SALDO DE LA CANTIDAD DE RECEPCION MENOS LO INGRESADO EN EL CIERRE DE PEDIDOS
                            'If Val(MskCan.Text) > (Val(VCantidadEntrada) - Val(VCantidadCerrarPedido)) Then
                            '    MsgBox "La Cantidad Es Mayor Que El Saldo Que Ingreso En La Recepcion, Se Va A Grabar El Registro Pero Verifique", vbExclamation, "Advertencia"
                                'MskCan.SetFocus
                                'Exit Sub
                            'End If
                '________________________________________________________________________________________________________
                
                VCantidadEntrada = MskCan.Text
                VFechaEntrada = MskFecRec.Text
                VNumeroPedido = TxtTexto.Item(2).Text
                VCodigo = TxtTexto.Item(1).Text
                
                'GRABA DATOS
                .Update
                
                        'If Err = 3022 Then
                        '        MsgBox "El Codigo De Materia Prima Ya Esta Ingresado Con Esta Recepcion Para Este Pedido", vbOKOnly + vbExclamation, "Verifique"
                        '        Exit Sub
                        If Err.Number <> 0 Then 'And Err <> 3022 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                Exit Sub
                        End If
                
                '---------- PEDIDO ------------------------------------------------------
                'BUSCA EL DOCUMENTO DE PEDIDO PARA SUMAR LA CANTIDAD ENTREGADA Y RESTAR EL SALDO POR ENTREGAR DE LA MATERIA PRIMA
                Set RBuscaPedido = Db.OpenRecordset("Select CantidadEntregada, SaldoPorEntregar, FechaParaEntregar, FechaEntregaTotal, DiasDeAtraso From DetallePedidosClientes Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'")
                    If RBuscaPedido.RecordCount > 0 Then
                        'EDITA EL REGISTRO DE PEDIDO Y ACTUALIZA DATOS
                        RBuscaPedido.Edit
                            RBuscaPedido!CantidadEntregada = Val(RBuscaPedido!CantidadEntregada) + Val(VCantidadEntrada)
                            RBuscaPedido!SaldoPorEntregar = Val(RBuscaPedido!SaldoPorEntregar) - Val(VCantidadEntrada)
                                            
                            'SI EL SALDO POR ENTREGAR YA ESTA EN CERO O MENOR QUE CERO ACTUALIZA LA FECHA DE ENTREGA Y CALCULA
                            'LOS DIAS DE ATRASO
                            If RBuscaPedido!SaldoPorEntregar <= 0 Then
                                'CAMBIA LA FECHA DE ENTREGA TOTAL POR LA ACTUAL DEL ULTIMO INGRESO
                                RBuscaPedido!FechaEntregaTotal = VFechaEntrada
                                            
                                'CALCULA LOS DIAS DE ATRASO
                                VDiasDeAtraso = (DateValue(RBuscaPedido!FechaParaEntregar) - DateValue(VFechaEntrada))
                                                
                                'SI LA VARIABLE VDIASDEATRASO ES MENOR QUE CERO ES PORQUE ENTREGO EL PEDIDO ANTES DE LA FECHA
                                If VDiasDeAtraso < 0 Then
                                    VDiasDeAtraso = VDiasDeAtraso * -1
                                Else
                                    VDiasDeAtraso = 0
                                End If
                                                
                                'MODIFICA LOS DIAS DE ATRASO EN EL PEDIDO
                                RBuscaPedido!DiasDeAtraso = VDiasDeAtraso
                            Else
                                If IsNull(RBuscaPedido!FechaEntregaTotal) Then
                                Else
                                    RBuscaPedido!FechaEntregaTotal = ""
                                    RBuscaPedido!DiasDeAtraso = "0"
                                End If
                            End If
                        'GRABA DATOS
                        RBuscaPedido.Update
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
        
                If GBorrar = False Then
                      MsgBox "Usted No Tiene Acceso a Esta Funcion, Consulte al Encargado", vbOKOnly + vbInformation, "Informacion"
                      Exit Sub
                End If
        
                VCantidadEntrada = MskCan.Text
                VFechaEntrada = MskFecRec.Text
                VNumeroPedido = TxtTexto.Item(2).Text
                VCodigo = TxtTexto.Item(1).Text
        
                VMensaje = MsgBox("Esta seguro de borrar el registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                If VMensaje = vbYes Then
                    .Delete
                    .MoveNext
                    
                            '---------- PEDIDO ------------------------------------------------------
                            'BUSCA EL DOCUMENTO DE PEDIDO PARA SUMAR LA CANTIDAD ENTREGADA Y RESTAR EL SALDO POR ENTREGAR DE LA MATERIA PRIMA
                            Set RBuscaPedido = Db.OpenRecordset("Select CantidadEntregada, SaldoPorEntregar, FechaParaEntregar, FechaEntregaTotal, DiasDeAtraso From DetallePedidosClientes Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'")
                                If RBuscaPedido.RecordCount > 0 Then
                                    'EDITA EL REGISTRO DE PEDIDO Y ACTUALIZA DATOS
                                    RBuscaPedido.Edit
                                        RBuscaPedido!CantidadEntregada = RBuscaPedido!CantidadEntregada - VCantidadEntrada
                                        RBuscaPedido!SaldoPorEntregar = RBuscaPedido!SaldoPorEntregar + VCantidadEntrada
                                                        
                                        'SI EL SALDO POR ENTREGAR ES MAYOR QUE CERO CAMBIA LA FECHA
                                        If RBuscaPedido!SaldoPorEntregar > 0 Then
                                            'CAMBIA LA FECHA DE ENTREGA TOTAL
                                            RBuscaPedido!FechaEntregaTotal = ""
                                            'MODIFICA LOS DIAS DE ATRASO EN EL PEDIDO
                                            RBuscaPedido!DiasDeAtraso = 0
                                        End If
                                    'GRABA DATOS
                                    RBuscaPedido.Update
                                End If
                    
                            If Err.Number = 3021 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbInformation, "Error"
                                    Exit Sub
                            End If
                End If
        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    'RECEPCION
                    If OptBusqueda.Item(0).Value = True Then
                        DataCerrarPedidos.RecordSource = ("Select * From CerrarPedidoClientes where Documento = '" & TxtBusqueda.Text & "' Order By FechaRecepcion")
                    'FECHAS
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataCerrarPedidos.RecordSource = ("Select * From CerrarPedidoClientes where FechaRecepcion >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaRecepcion <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# Order By FechaRecepcion")
                    'PEDIDO
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        DataCerrarPedidos.RecordSource = ("Select * From CerrarPedidoClientes where Pedido = '" & TxtBusqueda.Text & "' Order By FechaRecepcion")
                    End If
                    DataCerrarPedidos.Refresh
                    DGridCerrarPedidos.Refresh
                    TabDepartamentos.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    DataCerrarPedidos.RecordSource = "Select * From CerrarPedidoClientes Order By FechaRecepcion"
                    DataCerrarPedidos.Refresh
                    DGridCerrarPedidos.Refresh
                    TabDepartamentos.Tab = 1
        End If
    End With
    

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameCerrarPedidos.Enabled = True
         DataCerrarPedidos.Visible = False
         DGridCerrarPedidos.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameCerrarPedidos.Enabled = False
         DataCerrarPedidos.Visible = True
         DGridCerrarPedidos.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub



Private Sub DataCerrarPedidos_Reposition()
'BUSCA EL ENCABEZADO DE PEDIDO
        Set RBuscaSaldoEncabezado = Db.OpenRecordset("Select EP.Fecha, P.Descripcion, EP.Observaciones From EncabezadoPedidosClientes as EP, Clientes as P Where EP.Documento = '" & TxtTexto.Item(2).Text & "' And EP.Cliente = P.CodigoCliente")
                   If RBuscaSaldoEncabezado.RecordCount > 0 Then
                       TxtDatos.Text = ""
                       TxtDatos.Text = TxtDatos.Text & "Fecha Pedido     " & RBuscaSaldoEncabezado(0) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Cliente        " & RBuscaSaldoEncabezado(1) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Observaciones    " & RBuscaSaldoEncabezado(2) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "         Pedido           Entregado              Saldo     Dias      Entrega    Entregado   Atraso" & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "______________________________________________________________________________________________________" & vbCrLf
                       'BUSCA EL DETALLE DEL PEDIDO
                       Set RBuscaSaldoDetalle = Db.OpenRecordset("Select * From DetallePedidosClientes Where Documento = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                           Do Until RBuscaSaldoDetalle.EOF
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!CantidadPedido) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!CantidadEntregada) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!SaldoPorEntregar) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatInteger5(RBuscaSaldoDetalle!DiasPedido) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & RBuscaSaldoDetalle!FechaParaEntregar & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & RBuscaSaldoDetalle!FechaEntregaTotal & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatInteger5(RBuscaSaldoDetalle!DiasDeAtraso) & Space(3) & vbCrLf
                               RBuscaSaldoDetalle.MoveNext
                           Loop
                   Else
                       TxtDatos.Text = ""
                   End If
                   
        'BUSCA TODOS LOS CIERRES QUE TIENE EL PEDIDO
        Set RBuscaCierrePedidos = Db.OpenRecordset("Select Documento, FechaOperacion, FechaRecepcion, Cantidad From CerrarPedidoClientes Where Pedido = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
            If RBuscaCierrePedidos.RecordCount > 0 Then
                    TxtDatos2.Text = ""
                    TxtDatos2.Text = TxtDatos2.Text & "Documento      Fecha Operacion     Fecha Despacho              Cantidad" & vbCrLf
                    TxtDatos2.Text = TxtDatos2.Text & "___________________________________________________________________________________________________" & vbCrLf
                        Do Until RBuscaCierrePedidos.EOF
                                TxtDatos2.Text = TxtDatos2.Text & FormatString15(RBuscaCierrePedidos!Documento) & Space(5) & RBuscaCierrePedidos!FechaOperacion & Space(10) & RBuscaCierrePedidos!FechaRecepcion & Space(5) & FormatSingle(RBuscaCierrePedidos!Cantidad) & vbCrLf
                            RBuscaCierrePedidos.MoveNext
                        Loop
            Else
                    TxtDatos2.Text = ""
            End If
                   
        
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BMateriaPrima = True Then
            TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
            'MskCan.Text = DBGridBusqueda.Columns(1).Text
            TxtTexto.Item(1).SetFocus
        ElseIf BPedido = True Then
            TxtTexto.Item(2).Text = DBGridBusqueda.Columns(1).Text
            TxtTexto.Item(2).SetFocus
        End If
            FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        
        If KeyAscii = 43 Then
            If BMateriaPrima = True Then
                TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
                'MskCan.Text = DBGridBusqueda.Columns(1).Text
                TxtTexto.Item(1).SetFocus
            ElseIf BPedido = True Then
                TxtTexto.Item(2).Text = DBGridBusqueda.Columns(1).Text
                TxtTexto.Item(2).SetFocus
            End If
                FrameBusqueda.Visible = False
        End If
End Sub

Private Sub DGridCerrarPedidos_HeadClick(ByVal ColIndex As Integer)
        DataCerrarPedidos.RecordSource = "Select * from CerrarPedidoClientes Order by " & DGridCerrarPedidos.Columns(ColIndex).DataField
        DataCerrarPedidos.Refresh
        DGridCerrarPedidos.Refresh
End Sub

Private Sub Form_Activate()
'BUSCA EL ENCABEZADO DE PEDIDO
        Set RBuscaSaldoEncabezado = Db.OpenRecordset("Select EP.Fecha, P.Descripcion, EP.Observaciones From EncabezadoPedidosClientes as EP, Clientes as P Where EP.Documento = '" & TxtTexto.Item(2).Text & "' And EP.Cliente = P.CodigoCliente")
                   If RBuscaSaldoEncabezado.RecordCount > 0 Then
                       TxtDatos.Text = ""
                       TxtDatos.Text = TxtDatos.Text & "Fecha Pedido     " & RBuscaSaldoEncabezado(0) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Cliente          " & RBuscaSaldoEncabezado(1) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Observaciones    " & RBuscaSaldoEncabezado(2) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "         Pedido           Entregado              Saldo     Dias      Entrega    Entregado   Atraso" & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "___________________________________________________________________________________________________" & vbCrLf
                       'BUSCA EL DETALLE DEL PEDIDO
                       Set RBuscaSaldoDetalle = Db.OpenRecordset("Select * From DetallePedidosClientes Where Documento = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                           Do Until RBuscaSaldoDetalle.EOF
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!CantidadPedido) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!CantidadEntregada) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!SaldoPorEntregar) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatInteger5(RBuscaSaldoDetalle!DiasPedido) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & RBuscaSaldoDetalle!FechaParaEntregar & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & RBuscaSaldoDetalle!FechaEntregaTotal & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatInteger5(RBuscaSaldoDetalle!DiasDeAtraso) & Space(3) & vbCrLf
                               RBuscaSaldoDetalle.MoveNext
                           Loop
                   Else
                       TxtDatos.Text = ""
                   End If
                   
                'BUSCA TODOS LOS CIERRES QUE TIENE EL PEDIDO
                Set RBuscaCierrePedidos = Db.OpenRecordset("Select Documento, FechaOperacion, FechaRecepcion, Cantidad From CerrarPedidoClientes Where Pedido = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                    If RBuscaCierrePedidos.RecordCount > 0 Then
                            TxtDatos2.Text = ""
                            TxtDatos2.Text = TxtDatos2.Text & "Documento      Fecha Operacion     Fecha Despacho              Cantidad" & vbCrLf
                            TxtDatos2.Text = TxtDatos2.Text & "___________________________________________________________________________________________________" & vbCrLf
                                Do Until RBuscaCierrePedidos.EOF
                                        TxtDatos2.Text = TxtDatos2.Text & FormatString15(RBuscaCierrePedidos!Documento) & Space(5) & RBuscaCierrePedidos!FechaOperacion & Space(10) & RBuscaCierrePedidos!FechaRecepcion & Space(5) & FormatSingle(RBuscaCierrePedidos!Cantidad) & vbCrLf
                                    RBuscaCierrePedidos.MoveNext
                                Loop
                    Else
                            TxtDatos2.Text = ""
                    End If
                    
End Sub

Private Sub Form_Load()
    DataCerrarPedidos.Connect = GConnect
    DataBusqueda.Connect = GConnect
    
    DataCerrarPedidos.DatabaseName = BasedeDatos
    DataBusqueda.DatabaseName = BasedeDatos
End Sub

Private Sub MskCan_GotFocus()
        MskCan.SelStart = 0
        MskCan.SelLength = Len(MskCan.Text)
End Sub

Private Sub MskCan_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskFecRec_GotFocus()
            MskFecRec.SelStart = 0
            MskFecRec.SelLength = Len(MskFecRec.Text)
End Sub

Private Sub MskFecRec_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If Index = 0 Then
            LblFecIni.Visible = False
            LblFecFin.Visible = False
            DtpFecIni.Visible = False
            DtpFecFin.Visible = False
            LblBusqueda.Caption = "No. Documento"
            TxtBusqueda.Visible = True
            TxtBusqueda.SetFocus
    ElseIf Index = 1 Then
            LblFecIni.Visible = True
            LblFecFin.Visible = True
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            LblBusqueda.Caption = ""
            TxtBusqueda.Visible = False
            DtpFecIni.SetFocus
    ElseIf Index = 2 Then
            LblFecIni.Visible = False
            LblFecFin.Visible = False
            DtpFecIni.Visible = False
            DtpFecFin.Visible = False
            LblBusqueda.Caption = "No. Pedido"
            TxtBusqueda.Visible = True
            TxtBusqueda.SetFocus
    End If
    
End Sub

Private Sub TabDepartamentos_Click(PreviousTab As Integer)
        DtpFecIni.Value = Date
        DtpFecFin.Value = Date
End Sub

Private Sub TxtBuscar_Change()
            
    'MATERIA PRIMA
    If BMateriaPrima = True Then
            'DESCRIPCION
            If OptOpcion1.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptOpcion2.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '" & TxtBuscar.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptOpcion2.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '*" & TxtBuscar.Text & "*'"
                End If
            'CODIGO
            ElseIf OptOpcion1.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptOpcion2.Item(0).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '" & TxtBuscar.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptOpcion2.Item(1).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & TxtBuscar.Text & "*'"
                End If
            End If
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.Columns(1).Width = "4000"

    End If

End Sub

Private Sub TxtBuscar_GotFocus()
        TxtBuscar.SelStart = 0
        TxtBuscar.SelLength = Len(TxtBuscar.Text)
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{TAB}"
        End If
End Sub

Private Sub TxtBusqueda_GotFocus()
        TxtBusqueda.SelStart = 0
        TxtBusqueda.SelLength = Len(TxtBusqueda.Text)
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
    If Index = 1 Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblMateriaPrima.Caption = ""
                End If

    ElseIf Index = 2 Then
    'BUSCA EL ENCABEZADO DE PEDIDO
        Set RBuscaSaldoEncabezado = Db.OpenRecordset("Select EP.Fecha, P.Descripcion, EP.Observaciones From EncabezadoPedidosClientes as EP, Clientes as P Where EP.Documento = '" & TxtTexto.Item(2).Text & "' And EP.Cliente = P.CodigoCliente")
                   If RBuscaSaldoEncabezado.RecordCount > 0 Then
                       TxtDatos.Text = ""
                       TxtDatos.Text = TxtDatos.Text & "Fecha Pedido     " & RBuscaSaldoEncabezado(0) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Cliente            " & RBuscaSaldoEncabezado(1) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Observaciones    " & RBuscaSaldoEncabezado(2) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "         Pedido           Entregado              Saldo     Dias      Entrega    Entregado   Atraso" & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "__________________________________________________________________________________________________" & vbCrLf
                       'BUSCA EL DETALLE DEL PEDIDO
                       Set RBuscaSaldoDetalle = Db.OpenRecordset("Select * From DetallePedidosClientes Where Documento = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                           Do Until RBuscaSaldoDetalle.EOF
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!CantidadPedido) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!CantidadEntregada) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!SaldoPorEntregar) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatInteger5(RBuscaSaldoDetalle!DiasPedido) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & RBuscaSaldoDetalle!FechaParaEntregar & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & RBuscaSaldoDetalle!FechaEntregaTotal & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatInteger5(RBuscaSaldoDetalle!DiasDeAtraso) & Space(3) & vbCrLf
                               RBuscaSaldoDetalle.MoveNext
                           Loop
                   Else
                       TxtDatos.Text = ""
                   End If
                   
                   'BUSCA TODOS LOS CIERRES QUE TIENE EL PEDIDO
                    Set RBuscaCierrePedidos = Db.OpenRecordset("Select Documento, FechaOperacion, FechaRecepcion, Cantidad From CerrarPedidoClientes Where Pedido = '" & TxtTexto.Item(2).Text & "' And Codigo = '" & TxtTexto.Item(1).Text & "'")
                        If RBuscaCierrePedidos.RecordCount > 0 Then
                                TxtDatos2.Text = ""
                                TxtDatos2.Text = TxtDatos2.Text & "Documento      Fecha Operacion     Fecha Despacho              Cantidad" & vbCrLf
                                TxtDatos2.Text = TxtDatos2.Text & "___________________________________________________________________________________________________" & vbCrLf
                                    Do Until RBuscaCierrePedidos.EOF
                                            TxtDatos2.Text = TxtDatos2.Text & FormatString15(RBuscaCierrePedidos!Documento) & Space(5) & RBuscaCierrePedidos!FechaOperacion & Space(10) & RBuscaCierrePedidos!FechaRecepcion & Space(5) & FormatSingle(RBuscaCierrePedidos!Cantidad) & vbCrLf
                                        RBuscaCierrePedidos.MoveNext
                                    Loop
                        Else
                                TxtDatos2.Text = ""
                        End If
                    
    End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    'BUSCA Y AGRUPA TODOS LOS CODIGOS DE MATERIA PRIMA QUE VINO EN LA RECEPCION DE BODEGA (pendiente)
    'AHORA SOLO BUSCA CODIGO DE MATERIA PRIMA
    If Index = 1 Then
        'DataBusqueda.RecordSource = "Select Codigo, Sum(Cantidad) From DetalleEntradasMateriaPrima Where Documento = '" & TxtTexto.Item(0).Text & "' Group By Codigo"
        DataBusqueda.RecordSource = ("Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima")
        BMateriaPrima = True
        BPedido = False
    'BUSCA TODOS LOS PEDIDOS DE LA MATERIA PRIMA
    ElseIf Index = 2 Then
        DataBusqueda.RecordSource = "Select P.Fecha, P.Documento, DP.CantidadPedido, DP.CantidadEntregada, DP.SaldoPorEntregar, Pr.Descripcion From EncabezadoPedidosClientes As P, Clientes as Pr, DetallePedidosClientes as DP Where DP.Codigo = '" & TxtTexto.Item(1).Text & "' And DP.SaldoPorEntregar > 0 And P.Documento = DP.Documento And P.Cliente = Pr.CodigoCliente Order By Fecha"
        BMateriaPrima = False
        BPedido = True
    End If
        DataBusqueda.Refresh
        DBGridBusqueda.Refresh
        
                    If Index = 1 Then
                        'DBGridBusqueda.Columns(0).Width = 1500
                        'DBGridBusqueda.Columns(1).Width = 1500
                        'DBGridBusqueda.Columns(0).Caption = "Codigo"
                        'DBGridBusqueda.Columns(1).Caption = "Total"
                        'DBGridBusqueda.Columns(1).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(0).Width = 1500
                        DBGridBusqueda.Columns(1).Width = 4000
                    ElseIf Index = 2 Then
                        DBGridBusqueda.Columns(0).Width = 1000
                        DBGridBusqueda.Columns(1).Width = 1200
                        DBGridBusqueda.Columns(2).Width = 1200
                        DBGridBusqueda.Columns(3).Width = 1200
                        DBGridBusqueda.Columns(4).Width = 1200
                        DBGridBusqueda.Columns(5).Width = 2500
                        DBGridBusqueda.Columns(0).Caption = "Fecha"
                        DBGridBusqueda.Columns(1).Caption = "Pedido"
                        DBGridBusqueda.Columns(2).Caption = "Inicio"
                        DBGridBusqueda.Columns(3).Caption = "Entregado"
                        DBGridBusqueda.Columns(4).Caption = "Saldo"
                        DBGridBusqueda.Columns(5).Caption = "Descripcion"
                        DBGridBusqueda.Columns(2).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(3).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(4).NumberFormat = "#,###,##0"
                    End If
        FrameBusqueda.Visible = True
        TxtBuscar.SetFocus
        
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
                'BUSCA Y AGRUPA TODOS LOS CODIGOS DE MATERIA PRIMA QUE VINO EN LA RECEPCION DE BODEGA
                'AHORA SOLO VA A BUSCAR TODOS LOS CODIGOS
                If Index = 1 Then
                    'DataBusqueda.RecordSource = "Select Codigo, Sum(Cantidad) From DetalleEntradasMateriaPrima Where Documento = '" & TxtTexto.Item(0).Text & "' Group By Codigo"
                    DataBusqueda.RecordSource = ("Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima")
                    BMateriaPrima = True
                    BPedido = False
                'BUSCA TODOS LOS PEDIDOS DE LA MATERIA PRIMA
                ElseIf Index = 2 Then
                    DataBusqueda.RecordSource = "Select P.Fecha, P.Documento, DP.CantidadPedido, DP.CantidadEntregada, DP.SaldoPorEntregar, Pr.Descripcion From EncabezadoPedidosClientes As P, Clientes as Pr, DetallePedidosClientes as DP Where DP.Codigo = '" & TxtTexto.Item(1).Text & "' And DP.SaldoPorEntregar > 0 And P.Documento = DP.Documento And P.Cliente = Pr.CodigoCliente Order By Fecha"
                    BMateriaPrima = False
                    BPedido = True
                End If
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    
                    If Index = 1 Then
                        'DBGridBusqueda.Columns(0).Width = 1500
                        'DBGridBusqueda.Columns(1).Width = 1500
                        'DBGridBusqueda.Columns(0).Caption = "Codigo"
                        'DBGridBusqueda.Columns(1).Caption = "Total"
                        'DBGridBusqueda.Columns(1).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(0).Width = 1500
                        DBGridBusqueda.Columns(1).Width = 4000
                    ElseIf Index = 2 Then
                        DBGridBusqueda.Columns(0).Width = 1000
                        DBGridBusqueda.Columns(1).Width = 1200
                        DBGridBusqueda.Columns(2).Width = 1200
                        DBGridBusqueda.Columns(3).Width = 1200
                        DBGridBusqueda.Columns(4).Width = 1200
                        DBGridBusqueda.Columns(5).Width = 2500
                        DBGridBusqueda.Columns(0).Caption = "Fecha"
                        DBGridBusqueda.Columns(1).Caption = "Pedido"
                        DBGridBusqueda.Columns(2).Caption = "Inicio"
                        DBGridBusqueda.Columns(3).Caption = "Entregado"
                        DBGridBusqueda.Columns(4).Caption = "Saldo"
                        DBGridBusqueda.Columns(5).Caption = "Descripcion"
                        DBGridBusqueda.Columns(2).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(3).NumberFormat = "#,###,##0"
                        DBGridBusqueda.Columns(4).NumberFormat = "#,###,##0"
                    End If
                        FrameBusqueda.Visible = True
                        TxtBuscar.SetFocus
        End If
        
End Sub
