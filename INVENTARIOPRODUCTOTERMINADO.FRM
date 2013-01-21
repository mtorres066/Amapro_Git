VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form InventarioProductoTerminado 
   BackColor       =   &H00800000&
   Caption         =   "Inventario De Producto Terminado"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9015
   Icon            =   "InventarioProductoTerminado.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab TabForma 
      Height          =   7215
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   12726
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "InventarioProductoTerminado.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePlatinas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameBusqueda"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "InventarioProductoTerminado.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridInventario"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "InventarioProductoTerminado.frx":1036
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "CmdBuscar"
      Tab(2).Control(2)=   "CmdActualizar"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "TxtBuscar"
      Tab(2).Control(5)=   "FrameTiposDeBusqueda"
      Tab(2).ControlCount=   6
      Begin VB.Frame FrameBusqueda 
         Caption         =   "Busqueda de Datos"
         Height          =   7095
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Visible         =   0   'False
         Width           =   9015
         Begin VB.CommandButton CmdSale 
            Height          =   615
            Left            =   8040
            Picture         =   "InventarioProductoTerminado.frx":1488
            Style           =   1  'Graphical
            TabIndex        =   48
            ToolTipText     =   "Sale De Consulta"
            Top             =   240
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
            Left            =   2400
            Options         =   0
            ReadOnly        =   0   'False
            RecordsetType   =   1  'Dynaset
            RecordSource    =   ""
            Top             =   2400
            Visible         =   0   'False
            Width           =   3375
         End
         Begin MSDBGrid.DBGrid DBGridBusqueda 
            Bindings        =   "InventarioProductoTerminado.frx":18CA
            Height          =   6615
            Left            =   120
            OleObjectBlob   =   "InventarioProductoTerminado.frx":18E5
            TabIndex        =   28
            Top             =   240
            Width           =   8535
         End
      End
      Begin VB.Frame FrameTiposDeBusqueda 
         Caption         =   "Tipos De Busqueda"
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
         Height          =   1815
         Left            =   -69360
         TabIndex        =   46
         Top             =   840
         Width           =   3135
         Begin VB.OptionButton OptOpcion2 
            Caption         =   "Cualquier Inicial"
            Height          =   1215
            Index           =   1
            Left            =   1680
            Picture         =   "InventarioProductoTerminado.frx":22BF
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   360
            Width           =   1220
         End
         Begin VB.OptionButton OptOpcion2 
            Caption         =   "Palabra Inicial"
            Height          =   1215
            Index           =   0
            Left            =   240
            Picture         =   "InventarioProductoTerminado.frx":2409
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   1220
         End
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -69120
         TabIndex        =   25
         ToolTipText     =   " "
         Top             =   3360
         Width           =   2925
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
         ForeColor       =   &H000000FF&
         Height          =   1815
         Left            =   -74880
         TabIndex        =   32
         Top             =   840
         Width           =   4125
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Cliente"
            Height          =   1185
            Index           =   2
            Left            =   2760
            Picture         =   "InventarioProductoTerminado.frx":2553
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   " "
            Top             =   360
            Width           =   1220
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Bodega"
            Height          =   1185
            Index           =   1
            Left            =   1440
            Picture         =   "InventarioProductoTerminado.frx":2E1D
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   " "
            Top             =   360
            Width           =   1220
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Codigo Producto"
            Height          =   1185
            Index           =   0
            Left            =   120
            Picture         =   "InventarioProductoTerminado.frx":325F
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   1220
         End
      End
      Begin VB.Frame FramePlatinas 
         Caption         =   "Datos de Inventario Producto Terminado"
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
         TabIndex        =   30
         Top             =   720
         Width           =   8715
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Cuerpos"
            DataSource      =   "DataInventario"
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
            Index           =   14
            Left            =   2040
            TabIndex        =   5
            ToolTipText     =   " "
            Top             =   2400
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Minimo"
            DataSource      =   "DataInventario"
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
            Index           =   13
            Left            =   2040
            TabIndex        =   4
            ToolTipText     =   " "
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Descripcion"
            DataSource      =   "DataInventario"
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
            Index           =   12
            Left            =   2040
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   960
            Width           =   6495
         End
         Begin MSMask.MaskEdBox MskCos 
            DataField       =   "Costo"
            DataSource      =   "DataInventario"
            Height          =   285
            Left            =   2040
            TabIndex        =   3
            Top             =   1680
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Bodega"
            DataSource      =   "DataInventario"
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
            Index           =   10
            Left            =   2040
            MaxLength       =   3
            TabIndex        =   2
            ToolTipText     =   " "
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "UsuarioEditar"
            DataSource      =   "DataInventario"
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
            Left            =   7200
            MaxLength       =   10
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   3480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "UsuarioAgregar"
            DataSource      =   "DataInventario"
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
            Index           =   8
            Left            =   5760
            MaxLength       =   10
            TabIndex        =   42
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   3480
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "UnidadMedida"
            DataSource      =   "DataInventario"
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
            Left            =   5520
            MaxLength       =   30
            TabIndex        =   12
            ToolTipText     =   " "
            Top             =   5760
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Existencia"
            DataSource      =   "DataInventario"
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
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   11
            TabStop         =   0   'False
            ToolTipText     =   " "
            Top             =   5760
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Traslados"
            DataSource      =   "DataInventario"
            Height          =   285
            Index           =   4
            Left            =   5520
            MaxLength       =   10
            TabIndex        =   10
            ToolTipText     =   " "
            Top             =   5040
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Salidas"
            DataSource      =   "DataInventario"
            Height          =   285
            Index           =   3
            Left            =   5520
            MaxLength       =   10
            TabIndex        =   9
            ToolTipText     =   " "
            Top             =   4680
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Entradas"
            DataSource      =   "DataInventario"
            Height          =   285
            Index           =   2
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   8
            ToolTipText     =   " "
            Top             =   5040
            Width           =   1215
         End
         Begin VB.TextBox TxtTexto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Inicio"
            DataSource      =   "DataInventario"
            Height          =   285
            Index           =   1
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   7
            ToolTipText     =   " "
            Top             =   4680
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MskFecInv 
            DataField       =   "FechaInventario"
            DataSource      =   "DataInventario"
            Height          =   285
            Left            =   2040
            TabIndex        =   6
            Top             =   2760
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "CodigoProducto"
            DataSource      =   "DataInventario"
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
            Left            =   2040
            MaxLength       =   12
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   600
            Width           =   1815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cuerpos X Tarima O Bulto"
            Height          =   195
            Index           =   15
            Left            =   120
            TabIndex        =   52
            Top             =   2400
            Width           =   1830
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Existencia Minima"
            Height          =   195
            Index           =   14
            Left            =   120
            TabIndex        =   51
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Index           =   13
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   1020
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Costo x Unidad Medida"
            Height          =   195
            Index           =   12
            Left            =   120
            TabIndex        =   49
            Top             =   1680
            Width           =   1650
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo De Bodega"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   45
            Top             =   1320
            Width           =   1350
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
            Left            =   3360
            TabIndex        =   44
            Top             =   1320
            Width           =   5175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Movimientos"
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
            TabIndex        =   41
            Top             =   4320
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Unidad De Medida"
            Height          =   195
            Index           =   8
            Left            =   4080
            TabIndex        =   40
            Top             =   5760
            Width           =   1335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha De Inventario"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   39
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   38
            Top             =   5760
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H008080FF&
            Caption         =   "Traslados"
            Height          =   255
            Index           =   4
            Left            =   4680
            TabIndex        =   37
            Top             =   5040
            Width           =   975
         End
         Begin VB.Label Label1 
            BackColor       =   &H008080FF&
            Caption         =   "Salidas"
            Height          =   255
            Index           =   3
            Left            =   4680
            TabIndex        =   36
            Top             =   4680
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H008080FF&
            Caption         =   "Entradas"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   35
            Top             =   5040
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H008080FF&
            Caption         =   "Inicio"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   34
            Top             =   4680
            Width           =   615
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Producto"
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
            TabIndex        =   31
            Top             =   600
            Width           =   1425
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H008080FF&
            BackStyle       =   1  'Opaque
            Height          =   855
            Left            =   120
            Top             =   4560
            Width           =   8415
         End
      End
      Begin MSDBGrid.DBGrid DBGridInventario 
         Bindings        =   "InventarioProductoTerminado.frx":3569
         Height          =   6345
         Left            =   -74880
         OleObjectBlob   =   "InventarioProductoTerminado.frx":3586
         TabIndex        =   19
         ToolTipText     =   "Click En Encabezado De Columna Para Indexar"
         Top             =   720
         Width           =   8745
      End
      Begin MSForms.CommandButton CmdActualizar 
         Height          =   615
         Left            =   -69720
         TabIndex        =   27
         Top             =   4920
         Width           =   3495
         Caption         =   "Actualizar"
         PicturePosition =   196613
         Size            =   "6165;1085"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdBuscar 
         Height          =   615
         Left            =   -69720
         TabIndex        =   26
         Top             =   4080
         Width           =   3495
         Caption         =   "Buscar O Seleccionar Datos"
         PicturePosition =   196613
         Size            =   "6165;1085"
         Picture         =   "InventarioProductoTerminado.frx":56F2
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin VB.Label Lbletiqueta 
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
         Left            =   -72000
         TabIndex        =   33
         Top             =   3360
         Width           =   2775
      End
   End
   Begin VB.Data DataInventario 
      BackColor       =   &H80000014&
      Caption         =   "Inventario De Producto Terminado"
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
      RecordSource    =   "InventarioProductoTerminado"
      Top             =   8160
      Width           =   8775
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   7320
      MouseIcon       =   "InventarioProductoTerminado.frx":5B44
      Picture         =   "InventarioProductoTerminado.frx":5F86
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7320
      Width           =   1545
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5880
      MouseIcon       =   "InventarioProductoTerminado.frx":63C8
      Picture         =   "InventarioProductoTerminado.frx":680A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7320
      Width           =   1300
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4440
      MouseIcon       =   "InventarioProductoTerminado.frx":6D3C
      Picture         =   "InventarioProductoTerminado.frx":717E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7320
      Width           =   1300
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   3000
      MouseIcon       =   "InventarioProductoTerminado.frx":76B0
      Picture         =   "InventarioProductoTerminado.frx":7AF2
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7320
      Width           =   1300
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   1560
      MouseIcon       =   "InventarioProductoTerminado.frx":8024
      Picture         =   "InventarioProductoTerminado.frx":8466
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7320
      Width           =   1300
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   120
      MouseIcon       =   "InventarioProductoTerminado.frx":8998
      Picture         =   "InventarioProductoTerminado.frx":8DDA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7320
      Width           =   1300
   End
End
Attribute VB_Name = "InventarioProductoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BTiposMateriaPrima As Boolean
Dim BBodegas As Boolean

Dim RBuscaBodega As Recordset
Dim RBuscaDescripcion As Recordset

Sub botones()
    If Bandera = True Then
         FramePlatinas.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         TxtTexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataInventario.Visible = False
         FrameOpciones.Visible = False
         DBGridInventario.Visible = False
    Else
         FramePlatinas.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataInventario.Visible = True
         FrameOpciones.Visible = True
         DBGridInventario.Visible = True
    End If
End Sub

Private Sub CmdActualizar_Click()
    DataInventario.RecordSource = ("Select * from InventarioProductoTerminado")
    DataInventario.Refresh
    DBGridInventario.Refresh
    TabForma.Tab = 1
End Sub



Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
        With DataInventario.Recordset
                'AGREGAR
                If Index = 0 Then
                    .AddNew
                                If Err <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                    Exit Sub
                                End If
                                Bandera = True
                                botones
                                TxtTexto.Item(0).SetFocus
                                'SI ESTA EDITANDO AGREGAMOS EL USUARIO AL CAMPO USUARIO AGREGAR
                                TxtTexto.Item(8).Text = GUsuario
                'EDITAR
                ElseIf Index = 1 Then
                        .Edit
                                'SI HAY ERRORES
                                If Err <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                    Exit Sub
                                End If
                                Bandera = True
                                botones
                                TxtTexto.Item(0).SetFocus
                                'SI ESTA EDITANDO AGREGAMOS EL USUARIO AL CAMPO USUARIO EDITAR
                                TxtTexto.Item(9).Text = GUsuario
                'GRABAR
                ElseIf Index = 2 Then
                        .Update
                             'SI SE DUPLICA LA LLAVE
                             If Err = 3022 Then
                                  MsgBox "Codigo de Producto Terminado ya Existe", vbOKOnly + vbInformation, "Informacion"
                                  TxtTexto.Item(0).SetFocus
                             ElseIf Err <> 0 And Err <> 3022 Then
                                  MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                             Else
                                  Bandera = False
                                  botones
                                  CmdBotones.Item(0).SetFocus
                            End If
                'CANCELAR
                ElseIf Index = 3 Then
                        .CancelUpdate
                                'SI HAY ERRORES
                                If Err <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                End If
                                Bandera = False
                                botones
                'BORRAR
                ElseIf Index = 4 Then
                            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
                
                            If mensaje = vbOK Then
                                'BORRA EL REGISTRO
                                DataInventario.Recordset.Delete
                                    If Err <> 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                        Exit Sub
                                    End If
                                'SE MUEVE AL ULTIMO REGISTRO
                                DataInventario.Recordset.MoveLast
                            End If
                            'SI ESTA EN EL FINAL DE ARCHIVO
                            If DataInventario.Recordset.EOF Then
                                DataInventario.Recordset.MoveLast
                                If Err = 3021 Then
                                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                                End If
                            End If
                'SALIDA
                ElseIf Index = 5 Then
                            Unload Me
                End If
                
        End With
End Sub

Private Sub CmdBuscar_Click()
        'CODIGO DE PRODUCTO
        If OptOpcion.Item(0).Value = True Then
            If OptOpcion.Item(0).Value = True Then
                DataInventario.RecordSource = ("Select * from InventarioProductoTerminado where CodigoProducto like '" & TxtBuscar.Text & "*'")
            Else
                DataInventario.RecordSource = ("Select * from InventarioProductoTerminado where CodigoProducto like '*" & TxtBuscar.Text & "*'")
            End If
        'BODEGA
        ElseIf OptOpcion.Item(1).Value = True Then
            If OptOpcion2.Item(0).Value = True Then
                DataInventario.RecordSource = ("Select * from InventarioProductoTerminado where Bodega like '" & TxtBuscar.Text & "*'")
            Else
                DataInventario.RecordSource = ("Select * from InventarioProductoTerminado where Bodega like '*" & TxtBuscar.Text & "*'")
            End If
        'CLIENTE
        ElseIf OptOpcion.Item(2).Value = True Then
            If OptOpcion2.Item(0).Value = True Then
                DataInventario.RecordSource = ("Select * from InventarioProductoTerminado where CodigoCliente like '" & TxtBuscar.Text & "*'")
            Else
                DataInventario.RecordSource = ("Select * from InventarioProductoTerminado where CodigoCliente like '*" & TxtBuscar.Text & "*'")
            End If
        End If
            DataInventario.Refresh
            DBGridInventario.Refresh
            TabForma.Tab = 1
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub



Private Sub DBGridBusqueda_DblClick()
    'TIPOS DE MATERIA PRIMA
    If BTiposMateriaPrima = True Then
        TxtTexto.Item(11).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(11).SetFocus
    'BODEGAS
    ElseIf BBodegas = True Then
        TxtTexto.Item(10).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(10).SetFocus
    End If
        FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        'TIPOS DE MATERIA PRIMA
        If BTiposMateriaPrima = True Then
            TxtTexto.Item(11).Text = DBGridBusqueda.Columns(0)
            TxtTexto.Item(11).SetFocus
        'BODEGAS
        ElseIf BBodegas = True Then
            TxtTexto.Item(10).Text = DBGridBusqueda.Columns(0)
            TxtTexto.Item(10).SetFocus
        End If
            FrameBusqueda.Visible = False
    End If

End Sub

Private Sub dbgridInventario_HeadClick(ByVal ColIndex As Integer)
    DataInventario.RecordSource = ("Select * from InventarioProductoTerminado order by " & DBGridInventario.Columns(ColIndex).DataField)
    DataInventario.Refresh
    DBGridInventario.Refresh
End Sub

Private Sub Form_Load()
   DataInventario.DatabaseName = BasedeDatos
   DataBusqueda.DatabaseName = BasedeDatos
End Sub

Private Sub MskCos_GotFocus()
    MskCos.SelStart = 0
    MskCos.SelLength = Len(MskCos.Text)
End Sub

Private Sub MskCos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskFecInv_GotFocus()
    MskFecInv.SelStart = 0
    MskFecInv.SelLength = Len(MskFecInv.Text)
End Sub

Private Sub MskFecInv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub OptOpcion_Click(Index As Integer)
    If Index = 0 Then
        Lbletiqueta.Caption = "Codigo De Platina"
    ElseIf Index = 1 Then
        Lbletiqueta.Caption = "Codigo De Bodega"
    ElseIf Index = 2 Then
        Lbletiqueta.Caption = "Codigo De Descripcion"
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

Private Sub TxtTexto_Change(Index As Integer)
    'BODEGAS
    If Index = 10 Then
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasProductoTerminado Where CodigoBodega = '" & TxtTexto.Item(10).Text & "'")
        If RBuscaBodega.RecordCount > 0 Then
            LblBodega.Caption = RBuscaBodega(0)
        Else
            LblBodega.Caption = ""
        End If
    End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    'BUSCA BODEGAS
    If Index = 10 Then
        BBodegas = True
        DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasProductoTerminado"
        FrameBusqueda.Visible = True
        DataBusqueda.Refresh
        DBGridBusqueda.Refresh
        DBGridBusqueda.SetFocus
    End If
        DBGridBusqueda.Columns(0).Width = "1000"
        DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
    TxtTexto.Item(Index).SelStart = 0
    TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
'SI PRECIONA LA TECLA DE ENTER
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    'SI PRECIONA LA TECLA DEL SIGNO +
    If KeyAscii = 43 Then
        'BUSCA BODEGAS
        If Index = 10 Then
            BBodegas = True
            DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasProductoTerminado"
            FrameBusqueda.Visible = True
            DataBusqueda.Refresh
            DBGridBusqueda.Refresh
            DBGridBusqueda.SetFocus
        End If
            DBGridBusqueda.Columns(0).Width = "1000"
            DBGridBusqueda.Columns(1).Width = "4000"
    End If
End Sub


Sub CalculaExistencia()
    TxtTexto.Item(5).Text = Val(TxtTexto.Item(1)) + Val(TxtTexto.Item(2))
    TxtTexto.Item(5).Text = TxtTexto.Item(5) - Val(TxtTexto.Item(3)) - Val(TxtTexto.Item(4))
End Sub

Private Sub TxtTexto_LostFocus(Index As Integer)
    'INICIO, ENTRADAS, SALIDAS, TRASLADOS
    If (Index = 1 Or Index = 2 Or Index = 3 Or Index = 4) Then
        CalculaExistencia
    End If
    
    'CODIGO DE MATERIA PRIMA
    If Index = 0 Then
            Set RBuscaDescripcion = Db.OpenRecordset("Select Descripcion From InventarioProductoTerminado Where CodigoProducto = '" & TxtTexto.Item(0).Text & "'")
            If RBuscaDescripcion.RecordCount > 0 Then
                TxtTexto.Item(12).Text = RBuscaDescripcion!Descripcion
                TxtTexto.Item(12).Enabled = False
            Else
                TxtTexto.Item(12).Text = ""
                TxtTexto.Item(12).Enabled = True
            End If
            
    End If
End Sub
