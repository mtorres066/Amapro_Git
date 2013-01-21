VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form TrasladosProductoTerminado 
   BackColor       =   &H00FF8080&
   Caption         =   "Traslados Producto Terminado"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "TrasladosProductoTerminado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
      Caption         =   "Busqueda De Datos"
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
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
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
         Left            =   6360
         TabIndex        =   6
         Top             =   240
         Width           =   4335
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
            TabIndex        =   8
            Top             =   360
            Width           =   1695
         End
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
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   10800
         Picture         =   "TrasladosProductoTerminado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataBuscar 
         Caption         =   "Productos"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1920
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin MSDBGrid.DBGrid DBGridBuscar 
         Bindings        =   "TrasladosProductoTerminado.frx":293C
         Height          =   7335
         Left            =   120
         OleObjectBlob   =   "TrasladosProductoTerminado.frx":2955
         TabIndex        =   4
         ToolTipText     =   "Doble Click o Esc Para Seleccionar"
         Top             =   960
         Width           =   11535
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4935
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab TabTraslados 
      Height          =   8055
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Encabezado"
      TabPicture(0)   =   "TrasladosProductoTerminado.frx":332D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameEncabezado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "TrasladosProductoTerminado.frx":377F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridDetalleDespachos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "FrameDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSDBGrid.DBGrid DBGridDetalleDespachos 
         Bindings        =   "TrasladosProductoTerminado.frx":3A99
         Height          =   4695
         Left            =   -74760
         OleObjectBlob   =   "TrasladosProductoTerminado.frx":3ABC
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2520
         Width           =   11415
      End
      Begin VB.Frame FrameDetalle 
         Caption         =   "Detalle Traslados De Producto Terminado"
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
         Height          =   7215
         Left            =   -74880
         TabIndex        =   46
         Top             =   720
         Width           =   11685
         Begin VB.Frame FrameDetalleCompras 
            Enabled         =   0   'False
            Height          =   1455
            Left            =   120
            TabIndex        =   58
            Top             =   240
            Width           =   11415
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
               Left            =   2040
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   360
               Width           =   7575
            End
            Begin VB.TextBox TxtCodPro 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "FichaTecnica"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   47
               Top             =   360
               Width           =   1815
            End
            Begin VB.TextBox TxtDocDet 
               Appearance      =   0  'Flat
               DataField       =   "Documento"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   5160
               TabIndex        =   62
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox TxtBod 
               Appearance      =   0  'Flat
               DataField       =   "Bodega"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   6720
               MaxLength       =   3
               TabIndex        =   54
               Top             =   1080
               Width           =   435
            End
            Begin VB.TextBox TxtTar 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Tarima"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   2760
               TabIndex        =   49
               Top             =   720
               Width           =   1035
            End
            Begin VB.TextBox TxtLin 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Linea"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   6720
               MaxLength       =   2
               TabIndex        =   51
               Top             =   720
               Width           =   435
            End
            Begin VB.TextBox TxtBat 
               Appearance      =   0  'Flat
               DataField       =   "Batch"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   4680
               TabIndex        =   53
               Top             =   1080
               Width           =   1185
            End
            Begin VB.TextBox TxtCueTar 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
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
               Left            =   120
               TabIndex        =   60
               TabStop         =   0   'False
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox TxtCal 
               Appearance      =   0  'Flat
               DataField       =   "Calidad"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   2760
               MaxLength       =   1
               TabIndex        =   52
               Top             =   1080
               Width           =   1035
            End
            Begin MSMask.MaskEdBox MskFecPro 
               DataField       =   "FechaProduccion"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   4680
               TabIndex        =   50
               Top             =   720
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   8438015
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TxtCanPro 
               DataField       =   "Cantidad"
               DataSource      =   "DataDetalleDespachos"
               Height          =   285
               Left            =   9720
               TabIndex        =   48
               Top             =   360
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,###,###"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               BackColor       =   &H80000004&
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
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   75
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label2 
               BackColor       =   &H80000004&
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
               Left            =   2040
               TabIndex        =   74
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000004&
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
               Index           =   0
               Left            =   9720
               TabIndex        =   73
               Top             =   120
               Width           =   855
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
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
               Height          =   195
               Index           =   7
               Left            =   2040
               TabIndex        =   72
               Top             =   720
               Width           =   585
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               Caption         =   "Fecha "
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
               Left            =   4080
               TabIndex        =   71
               Top             =   720
               Width           =   600
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
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
               Index           =   9
               Left            =   6120
               TabIndex        =   70
               Top             =   720
               Width           =   480
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
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
               Height          =   195
               Index           =   10
               Left            =   4080
               TabIndex        =   69
               Top             =   1080
               Width           =   510
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
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
               Index           =   5
               Left            =   6000
               TabIndex        =   68
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
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
               Height          =   195
               Index           =   21
               Left            =   2040
               TabIndex        =   67
               Top             =   1080
               Width           =   645
            End
            Begin VB.Label LblLin 
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
               Left            =   7200
               TabIndex        =   66
               Top             =   720
               Width           =   4095
            End
            Begin VB.Label LblBod 
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
               Left            =   7200
               TabIndex        =   65
               Top             =   1080
               Width           =   4095
            End
         End
         Begin VB.CommandButton CmdAgregar2 
            Caption         =   "A&gregar"
            Height          =   495
            Left            =   240
            Picture         =   "TrasladosProductoTerminado.frx":4EAE
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   6600
            Visible         =   0   'False
            Width           =   1700
         End
         Begin VB.CommandButton CmdGrabar2 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3840
            Picture         =   "TrasladosProductoTerminado.frx":53E0
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   6600
            Visible         =   0   'False
            Width           =   1700
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
            Left            =   9240
            Picture         =   "TrasladosProductoTerminado.frx":5912
            TabIndex        =   63
            Top             =   6600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton CmdCancelar2 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5640
            Picture         =   "TrasladosProductoTerminado.frx":5E44
            Style           =   1  'Graphical
            TabIndex        =   59
            Top             =   6600
            Visible         =   0   'False
            Width           =   1700
         End
         Begin VB.CommandButton CmdBorrar2 
            Caption         =   "B&orrar"
            Height          =   495
            Left            =   7440
            Picture         =   "TrasladosProductoTerminado.frx":6376
            Style           =   1  'Graphical
            TabIndex        =   61
            Top             =   6600
            Visible         =   0   'False
            Width           =   1700
         End
         Begin VB.CommandButton CmdEditar2 
            Caption         =   "Editar"
            Height          =   495
            Left            =   2040
            Picture         =   "TrasladosProductoTerminado.frx":68A8
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   6600
            Visible         =   0   'False
            Width           =   1700
         End
      End
      Begin VB.Frame FrameEncabezado 
         Caption         =   "Encabezado Traslados De Producto Terminado"
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
         Height          =   3855
         Left            =   120
         TabIndex        =   10
         Top             =   1800
         Width           =   11655
         Begin VB.Frame FrameCompras 
            Enabled         =   0   'False
            Height          =   2775
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   11415
            Begin VB.TextBox TxtDocIng 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               DataField       =   "Documento"
               DataSource      =   "DataDespachos"
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
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   22
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtLib 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Liberado"
               DataSource      =   "DataDespachos"
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtReq 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Requerido"
               DataSource      =   "DataDespachos"
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   30
               TabStop         =   0   'False
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TxtBatch 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               DataField       =   "Batch"
               DataSource      =   "DataDespachos"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1560
               TabIndex        =   26
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox TxtEstado 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               DataField       =   "Estado"
               DataSource      =   "DataDespachos"
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               DataField       =   "Observaciones"
               DataSource      =   "DataDespachos"
               Height          =   285
               Index           =   6
               Left            =   1560
               MaxLength       =   50
               TabIndex        =   28
               Top             =   2400
               Width           =   6855
            End
            Begin VB.TextBox TxtTipDoc 
               Appearance      =   0  'Flat
               DataField       =   "TipoDeDocumento"
               DataSource      =   "DataDespachos"
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   24
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox TxtBodTra 
               Appearance      =   0  'Flat
               DataField       =   "BodegaATrasladar"
               DataSource      =   "DataDespachos"
               Height          =   285
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   25
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox TxtNumDoc 
               Appearance      =   0  'Flat
               DataField       =   "NumeroDocumento"
               DataSource      =   "DataDespachos"
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
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   23
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox TxtLinea 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               DataField       =   "Linea"
               DataSource      =   "DataDespachos"
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
               Left            =   1560
               MaxLength       =   2
               TabIndex        =   27
               Top             =   2040
               Width           =   1215
            End
            Begin MSMask.MaskEdBox MskFec 
               DataField       =   "Fecha"
               DataSource      =   "DataDespachos"
               Height          =   285
               Left            =   1560
               TabIndex        =   21
               Top             =   240
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   45
               Top             =   240
               Width           =   540
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
               Left            =   2880
               TabIndex        =   44
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label6 
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
               Index           =   3
               Left            =   120
               TabIndex        =   43
               Top             =   960
               Width           =   1410
            End
            Begin VB.Label Label6 
               Caption         =   "No. Batch"
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
               Index           =   13
               Left            =   120
               TabIndex        =   42
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Clasificacion"
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
               Left            =   8520
               TabIndex        =   41
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Requerido"
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
               Left            =   8520
               TabIndex        =   40
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Liberado"
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
               Left            =   8520
               TabIndex        =   39
               Top             =   960
               Width           =   750
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   4
               Left            =   120
               TabIndex        =   38
               Top             =   2400
               Width           =   1275
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
               Left            =   2880
               TabIndex        =   37
               Top             =   960
               Width           =   5535
            End
            Begin VB.Label Label6 
               Caption         =   "Bodega a Trasladar"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Index           =   6
               Left            =   120
               TabIndex        =   36
               Top             =   1200
               Width           =   960
            End
            Begin VB.Label LblBodTra 
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
               Left            =   2880
               TabIndex        =   35
               Top             =   1320
               Width           =   5535
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "# Documento"
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
               Index           =   14
               Left            =   120
               TabIndex        =   34
               Top             =   600
               Width           =   1155
            End
            Begin VB.Label Label6 
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
               Index           =   15
               Left            =   120
               TabIndex        =   33
               Top             =   2040
               Width           =   480
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
               Height          =   255
               Left            =   2880
               TabIndex        =   32
               Top             =   2040
               Width           =   5535
            End
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar"
            Height          =   600
            Left            =   120
            Picture         =   "TrasladosProductoTerminado.frx":6DDA
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3120
            Width           =   1300
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "&Grabar"
            Enabled         =   0   'False
            Height          =   600
            Left            =   2760
            Picture         =   "TrasladosProductoTerminado.frx":730C
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   3120
            Width           =   1300
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   600
            Left            =   4080
            Picture         =   "TrasladosProductoTerminado.frx":783E
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   3120
            Width           =   1300
         End
         Begin VB.CommandButton CmdBorrar 
            Caption         =   "&Borrar"
            Height          =   600
            Left            =   5400
            Picture         =   "TrasladosProductoTerminado.frx":7D70
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   3120
            Width           =   1300
         End
         Begin VB.CommandButton CmdSalida 
            Appearance      =   0  'Flat
            Height          =   600
            Left            =   10560
            Picture         =   "TrasladosProductoTerminado.frx":82A2
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Salida"
            Top             =   3120
            Width           =   945
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar Documento"
            Height          =   600
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   3120
            Width           =   1300
         End
         Begin VB.CommandButton CmdEditar 
            Caption         =   "&Editar"
            Height          =   600
            Left            =   1440
            Picture         =   "TrasladosProductoTerminado.frx":A314
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3120
            Width           =   1300
         End
         Begin VB.CommandButton CmdImprimir 
            Caption         =   "&Imprimir"
            Height          =   600
            Left            =   9240
            Picture         =   "TrasladosProductoTerminado.frx":A846
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   3120
            Width           =   1300
         End
         Begin VB.CommandButton CmdBuscar2 
            Caption         =   "Siguiente Documento"
            Height          =   600
            Left            =   8040
            TabIndex        =   17
            Top             =   3120
            Width           =   1215
         End
      End
   End
   Begin VB.Data DataDetalleDespachos 
      Caption         =   "Detalle Despachos Producto Terminado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\erick\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleTrasladosProductoTerminado"
      Top             =   8160
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Data DataDespachos 
      Caption         =   "Despachos De Producto Terminado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\erick\Amapro Nuevo Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EncabezadoTrasladosProductoTerminado"
      Top             =   8040
      Width           =   11655
   End
End
Attribute VB_Name = "TrasladosProductoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As Long
Dim VDocumentoDetalle As Long
Dim VCantidad As Long
Dim VCodigoProducto As String
Dim VCantidadProducto As Long
Dim VBodega As String
Dim VBatch As Long
Dim VClasificacion As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim BTipoDeDocumento As Boolean
Dim BProducto As Boolean
Dim BBodegaDetalle As Boolean
Dim BTransportistas As Boolean
Dim BEditar As Boolean

Dim VLinea As String

Dim RBuscaProducto As Recordset
Dim RMaximo As Recordset
Dim RBuscaBodega As Recordset
Dim RBuscaDetalle As Recordset
Dim RBuscaEncabezado As Recordset
Dim RBuscaEntradasProductoTerminado As Recordset
Dim RBuscaDocumento As Recordset
Dim RBuscaTarima As Recordset
Dim RCuentaTarimas As Recordset
Dim RBuscaTipoDocumento As Recordset
Dim RBuscaLinea As Recordset


Sub Botones1()
    If Bandera = True Then
         FrameCompras.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdBuscar2.Enabled = False
         CmdSalida.Enabled = False
         CmdImprimir.Enabled = False
         DataDespachos.Visible = False
    Else
         FrameCompras.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdBuscar2.Enabled = True
         CmdSalida.Enabled = True
         CmdImprimir.Enabled = True
         DataDespachos.Visible = True
         
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleCompras.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
         CmdEditar2.Enabled = False
    Else
         FrameDetalleCompras.Enabled = False
         CmdAgregar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
         CmdImprimir.Enabled = True
         CmdBuscar.Enabled = True
         CmdEditar2.Enabled = True
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
    'AGREGA DATOS
    DataDetalleDespachos.Recordset.AddNew
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
    Bandera2 = True
    Botones2
    DBGridDetalleDespachos.Enabled = False
    TxtDocDet.Text = VDocumento
    
    'ASIGNA LA BODEGA INGRESADA EN EL ENCABEZADO
    'TxtBod.Text = VBodega
    'TxtBat.Text = VBatch
    
    TxtCodPro.SetFocus
    TxtDesPro.Text = ""
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            If GBorrar = True Then
                    'NO HACE NADA PORQUE SI TIENE ACCESO A BORRAR
            ElseIf TxtEstado.Text = "LIBERADO" Then
                    'VERIFICA SI YA FUE LIBERADA LA ENTRADA
                    MsgBox "Este Documento No Se Puede BORRAR Porque Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
                    Exit Sub
            End If
            VDocumento = TxtDocIng.Text
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                        'BORRA EL ENCABEZADO
                        DataDespachos.Recordset.Delete
                        If Err <> 0 Then
                            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        DataDespachos.Recordset.MoveLast
                MousePointer = 0
            End If
            If DataDespachos.Recordset.EOF Then
                DataDespachos.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
                        
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                                        
                   'BORRA EL DETALLE DE LA ENTRADA
                    DataDetalleDespachos.Recordset.Delete
                    
                    If Err <> 0 Then
                       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                       Exit Sub
                    End If
                    'SELECCIONA TODOS LOS DETALLES DE LA ENTRADAS
                    DataDetalleDespachos.RecordSource = ("Select * from DetalleEntradasProductoTermin where documento = " & VDocumento & " order By Linea, Batch, Tarima")
                    DataDetalleDespachos.Refresh
                    DBGridDetalleDespachos.Refresh
                MousePointer = 0
            End If
  
            If DataDespachos.Recordset.EOF Then
                DataDespachos.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
End Sub

Private Sub CmdBuscar_Click()
    mensaje = InputBox("Documento a Buscar")
    If mensaje = "" Then
    Else
        DataDespachos.Recordset.FindFirst ("NumeroDocumento = '" & mensaje & "'")
    End If
    'SI HAY ERROR
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
End Sub

Private Sub CmdBuscar2_Click()
    'mensaje = InputBox("Documento a Buscar")
    If mensaje = "" Then
    Else
        DataDespachos.Recordset.FindNext ("NumeroDocumento = '" & mensaje & "'")
    End If
    'SI HAY ERROR
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
    'CANCELA LOS CAMBIOS
    DataDespachos.Recordset.CancelUpdate
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error "
        Err.Clear
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    FrameDetalle.Visible = True
    DBGridDetalleDespachos.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    'CANCELA LOS DATOS CAMBIADOS Y GRABA LOS DATOS COMO ESTABAN
    DataDetalleDespachos.Recordset.CancelUpdate
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    DBGridDetalleDespachos.Enabled = True
    Bandera2 = False
    Botones2

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
    BEditar = True

    'VERIFICA SI YA FUE LIBERADA LA ENTRADA
    If GEditar = True Then
        'NO HACE NADA PORQUE SI TIENE ACCESO A EDITAR
    ElseIf TxtEstado.Text = "LIBERADO" Then
        MsgBox "Esta Documento No Se Puede EDITAR Porque Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
        Exit Sub
    End If
    
    'EDITA EL REGISTRO
    DataDespachos.Recordset.Edit
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    MskFec.SetFocus
    
    'GRABA EL USUARIO QUE ESTA EDITANDO
    TxtReq.Text = GUsuario
    
    FrameDetalle.Visible = False
    DBGridDetalleDespachos.Visible = False
    
End Sub


Private Sub CmdEditar2_Click()
    On Error Resume Next
    'AGREGA DATOS
    DataDetalleDespachos.Recordset.Edit
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
    Bandera2 = True
    Botones2
    DBGridDetalleDespachos.Enabled = False
        
    TxtCodPro.SetFocus
    

End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
    
    'GUARDA VARIABLES
    VCantidad = TxtCanPro.Text
    VCodigoProducto = TxtCodPro.Text
        
    'REVISAMOS DATOS
    If Not IsNumeric(TxtCanPro.Text) Then
       MsgBox "Cantidad De Producto Incorrecta", vbOKOnly + vbCritical, "Error"
       TxtCanPro.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS EL BATCH DE DETALLE
    If Not IsNumeric(TxtBat.Text) Then
       MsgBox "Numero De Bath Incorrecto", vbOKOnly + vbCritical, "Error"
       TxtBat.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS LA TARIMA
    If Not IsNumeric(TxtTar.Text) Then
       MsgBox "Numero De Tarima Incorrecto", vbOKOnly + vbCritical, "Error"
       TxtBat.SetFocus
       Exit Sub
    End If
    
    
    'REVISAMOS DATOS
    If Not IsDate(MskFecPro.Text) Then
       MsgBox "Fecha De Produccion Incorrecta", vbOKOnly + vbCritical, "Error"
       MskFecPro.SetFocus
       Exit Sub
    End If
    
        
        'REVISA SI LA TARIMA EXISTE EN LA ENTRADAS DE PRODUCTO TERMINADO
        Set RBuscaEntradasProductoTerminado = Db.OpenRecordset("Select * From DetalleEntradasProductoTermina Where FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTar.Text & " And Linea = '" & Txtlin.Text & "' And FichaTecnica = '" & TxtCodPro.Text & "'")
    
            If RBuscaEntradasProductoTerminado.RecordCount > 0 Then
                'SI LA ENCUENTRA NO HACE NADA
            Else
                    MsgBox "La Tarima No Existe", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
            End If
                
    'GRABA DATOS
    DataDetalleDespachos.Recordset.Update
        
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
        
    Bandera2 = False
    Botones2
         
    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE EL TRASLADO QUE SE ESTA GRABANDO
    DataDetalleDespachos.RecordSource = ("Select * from DetalleEntradasProductoTermin where Documento = " & VDocumento & " Order by Linea, Batch, Tarima")
    DataDetalleDespachos.Refresh
    DBGridDetalleDespachos.Refresh
           
    DBGridDetalleDespachos.Enabled = True
    TxtDesPro.Text = ""
    CmdAgregar2.SetFocus
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
    BEditar = False

    'AGREGA DATOS
    DataDespachos.Recordset.AddNew
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    Bandera = True
    Botones1
    
    'ASIGNA EL USUARIO
    TxtReq.Text = GUsuario
    'ASIGNA LA FECHA ACTUAL
    MskFec.Text = Format(Date, "dd/mm/yyyy")
    MskFec.SetFocus
    'COLOCA EL ESTADO DE LA ENTRADA
    TxtEstado.Text = "NO LIBERADA"
    
    TxtCueTar.Text = ""
    
    'BUSCA EL DOCUMENTO MAXIMO Y LE AGREGA UNO MAS
    Set RMaximo = Db.OpenRecordset("Select max(Documento) from EncabezadoTrasladosProductoTer")
        If RMaximo.RecordCount > 0 Then
            If IsNull(RMaximo(0)) Then
                TxtDocIng.Text = "1"
            Else
                TxtDocIng.Text = Val(RMaximo(0)) + 1
            End If
        End If
        
    FrameDetalle.Visible = False
    DBGridDetalleDespachos.Visible = False

End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next

    'OSEA QUE SI ESTA AGREGANDO UN REGISTRO
    If BEditar = False Then
            'BUSCA SI YA EXISTE EL NUMERO DE DOCUMENTO PARA ESTE TIPO DE DOCUMENTO
            Set RBuscaDocumento = Db.OpenRecordset("Select * From EncabezadoTrasladosProductoTer Where TipoDeDocumento = '" & TxtTipDoc.Text & "' And NumeroDocumento = '" & TxtNumDoc.Text & "'")
                    If RBuscaDocumento.RecordCount > 0 Then
                        MsgBox "Numero Documento Para Este Tipo De Documento Ya Existe", vbOKOnly + vbInformation, "Informacion"
                        TxtTipDoc.SetFocus
                        Exit Sub
                    End If
    End If

MousePointer = 11

    VDocumento = TxtDocIng.Text
    VBatch = TxtBatch.Text
    VLinea = TxtLinea.Text
               
    'REVISA TIPO DE DOCUMENTO
    If TxtTipDoc.Text = "" Then
            MsgBox "Tipo De Documento No Puede Estar Vacia", vbOKOnly + vbInformation, "Informacion"
            TxtTipDoc.SetFocus
            Exit Sub
    End If
    
    'REVISA BODEGA A TRASLADAR
    If TxtBodTra.Text = "" Then
            MsgBox "Bodega a Trasladar No Puede Estar Vacia", vbOKOnly + vbInformation, "Informacion"
            TxtBodTra.SetFocus
            Exit Sub
    End If
        
    'REVISA LINEA
    If TxtLinea.Text = "" Then
            MsgBox "Linea No Puede Estar Vacia", vbOKOnly + vbInformation, "Informacion"
            TxtLinea.SetFocus
            Exit Sub
    End If
    
    'REVISA EL BATCH
    If Not IsNumeric(TxtBatch.Text) Then
            MsgBox "El Batch Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
            TxtBatch.SetFocus
            Exit Sub
    End If
    
               
    'GRABA DATOS
    DataDespachos.Recordset.Update
    
    If Err <> 0 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    
            'BUSCAMOS SI CON ESTE DOCUMENTO YA FUERON SELECCIONADAS ALGUNAS TARIMAS
            'SI YA HAY TARIMAS YA NO AGREGA MAS
            Set RBuscaDocumento = Db.OpenRecordset("Select Documento From DetalleEntradasProductoTermin Where Documento = " & VDocumento)
                'SI ENCUENTRA DATOS
                If RBuscaDocumento.RecordCount > 0 Then
                'NO SELECCIONA NINGUNA TARIMA
                        MsgBox "Ya hay tarimas ingresadas con este documento, si nesesita agregar mas, tiene que hacerlo manual", vbOKOnly + vbInformation, "Informacion"
                Else
                        'INICIALIZA EL RECORDSET PARA AGREGAR DATOS
                        Set RBuscaDetalle = Db.OpenRecordset("Select * From DetalleEntradasProductoTermin")
                        
                        'SELECCIONA TODAS LAS TARIMAS DE ENTRADAS DE PRODUCTO TERMINADO DE ACUERDO AL BATCH
                        'Y QUE EL SALDO SEA MAYOR QUE CERO Y LA LINEA A LA QUE PERTENECEN
                        Set RBuscaEntradasProductoTerminado = Db.OpenRecordset("Select FichaTecnica, Tarima, Linea, FechaProduccion, Batch, Calidad, Bodega, Saldo From DetalleEntradasProductoTermina Where Batch = " & VBatch & " And Saldo > 0 And Linea = '" & VLinea & "'")
                                
                                'CREA UN CICLO CON LOS DATOS DE PRODUCCION DEL BATCH
                                Do Until RBuscaEntradasProductoTerminado.EOF
                                        'AGREGA LAS TARIMAS DE ESE BATCH
                                                RBuscaDetalle.AddNew
                                                    RBuscaDetalle!Documento = VDocumento
                                                    RBuscaDetalle!FichaTecnica = RBuscaEntradasProductoTerminado!FichaTecnica
                                                    RBuscaDetalle!Tarima = RBuscaEntradasProductoTerminado!Tarima
                                                    RBuscaDetalle!FechaProduccion = RBuscaEntradasProductoTerminado!FechaProduccion
                                                    RBuscaDetalle!Linea = RBuscaEntradasProductoTerminado!Linea
                                                    RBuscaDetalle!Batch = VBatch
                                                    RBuscaDetalle!Bodega = RBuscaEntradasProductoTerminado!Bodega
                                                    RBuscaDetalle!Cantidad = RBuscaEntradasProductoTerminado!Saldo
                                                    RBuscaDetalle!Calidad = RBuscaEntradasProductoTerminado!Calidad
                                                RBuscaDetalle.Update
                                    'SE MUEVE AL SIGUIENTE REGISTRO
                                    RBuscaEntradasProductoTerminado.MoveNext
                                Loop
                        
                End If
        
    'SELECCIONA TODOS LOS DETALLES DE EL TRASLADO
    DataDetalleDespachos.RecordSource = ("Select * from DetalleEntradasProductoTermin where Documento = " & VDocumento & " Order By Linea, Batch, Tarima")
    DataDetalleDespachos.Refresh
    DBGridDetalleDespachos.Refresh
            
    'MUEVE EL RECORDSET A EL DOCUMENTO ACTUAL PARA QUE SE ACTUALIZEN LOS CAMBIOS
    DataDespachos.Recordset.FindFirst ("Documento = " & VDocumento)
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Visible = True
    FrameDetalle.Enabled = True
    FrameEncabezado.Enabled = False
    
    DBGridDetalleDespachos.Visible = True
    DBGridDetalleDespachos.AllowDelete = True
    DBGridDetalleDespachos.AllowUpdate = True
    
    'ESCONDE EL DATA
    DataDespachos.Visible = False
        
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = False
    BotonesVisiblesEncabezado
        
    CmdAgregar2.SetFocus
    
    TabTraslados.Tab = 1
    
    
MousePointer = 0
    
End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDocIng.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
                CrReportes.Formulas(0) = "Texto = ''"
                CrReportes.SelectionFormula = "{EncabezadoTrasladosProductoTer.Documento} = " & TxtDocIng.Text
                CrReportes.ReportFileName = App.Path & "\ReporteSalidasProductoTerminado.rpt"
                CrReportes.Action = 1
                
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                End If
MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdTerminar_Click()
If CmdCancelar2.Enabled = True Then
     CmdCancelar2_Click
End If
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle
            
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = False
    'FrameDetalle.Visible = False
    FrameEncabezado.Enabled = True
    
    'VISUALIZA EL DATA
    DataDespachos.Visible = True
            
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesVisiblesEncabezado
    
    DBGridDetalleDespachos.AllowDelete = False
    DBGridDetalleDespachos.AllowUpdate = False
    
    TabTraslados.Tab = 0
End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub


Private Sub datadetalledespachos_Reposition()
        If IsNumeric(TxtDocDet.Text) Then
            'CUENTA CUANTAS TARIMAS TIENE EL DOCUMENTO
            Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From DetalleEntradasProductoTermin Where Documento = " & TxtDocDet.Text)
                If RCuentaTarimas.RecordCount > 0 Then
                    If IsNull(RCuentaTarimas(0)) Then
                        TxtCueTar.Text = "0 Tarimas"
                    Else
                        TxtCueTar.Text = RCuentaTarimas(0) & " Tarimas"
                    End If
                Else
                    TxtCueTar.Text = "0 Tarimas"
                End If
        End If
End Sub


Private Sub datadespachos_Error(DataErr As Integer, Response As Integer)
    On Error Resume Next
        If Err <> 0 Then
            'MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
        End If
End Sub

Private Sub datadespachos_Reposition()
    If IsNumeric(TxtDocIng.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
        DataDetalleDespachos.RecordSource = ("Select * from DetalleEntradasProductoTermin where Documento = " & TxtDocIng.Text & " Order by Linea, Batch, Tarima")
        DataDetalleDespachos.Refresh
        DBGridDetalleDespachos.Refresh
    End If
End Sub


Private Sub DBGridBuscar_DblClick()
    'BODEGA
    If BTipoDeDocumento = True Then
        TxtTipDoc.Text = DBGridBuscar.Columns(0)
        TxtTipDoc.SetFocus
    'BODEGA DETALLE
    ElseIf BBodegaDetalle = True Then
        TxtBod.Text = DBGridBuscar.Columns(0)
        TxtBod.SetFocus
    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        TxtCodPro.Text = DBGridBuscar.Columns(0)
        TxtCodPro.SetFocus
    'BodegasProductoTerminado
    ElseIf BTransportistas = True Then
        TxtBodTra.Text = DBGridBuscar.Columns(0)
        TxtBodTra.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
    'BODEGA
    If BTipoDeDocumento = True Then
        TxtTipDoc.Text = DBGridBuscar.Columns(0)
        TxtTipDoc.SetFocus
    'BODEGA DETALLE
    ElseIf BBodegaDetalle = True Then
        TxtBod.Text = DBGridBuscar.Columns(0)
        TxtBod.SetFocus
    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        TxtCodPro.Text = DBGridBuscar.Columns(0)
        TxtCodPro.SetFocus
    ElseIf BTransportistas = True Then
        TxtBodTra.Text = DBGridBuscar.Columns(0)
        TxtBodTra.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End If

End Sub
Private Sub Form_Activate()
    If IsNumeric(TxtDocIng.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL DESPACHO
        DataDetalleDespachos.RecordSource = ("Select * from DetalleEntradasProductoTermin where Documento = " & TxtDocIng.Text & " Order by Linea, Tarima")
        DataDetalleDespachos.Refresh
        DBGridDetalleDespachos.Refresh
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
End Sub

Private Sub Form_Load()
    DataDespachos.ConnectionString = GTipoProveedor
    DataDetalleDespachos.ConnectionString = GTipoProveedor
    DataBuscar.ConnectionString = GTipoProveedor
    
    DataDespachos.Refresh
    DataDetalleDespachos.Refresh
    DataBuscar.Refresh
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
Private Sub MskFecPro_GotFocus()
    MskFecPro.SelStart = 0
    MskFecPro.SelLength = Len(MskFecPro.Text)
End Sub
Private Sub MskFecPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub Txtbat_GotFocus()
    TxtBat.SelStart = 0
    TxtBat.SelLength = Len(TxtBat.Text)
End Sub
Private Sub Txtbat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub TxtBatch_GotFocus()
    TxtBatch.SelStart = 0
    TxtBatch.SelLength = Len(TxtBatch.Text)
End Sub
Private Sub TxtBatch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TxtBod_Change()
            Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasProductoTerminado Where CodigoBodega = '" & TxtBod.Text & "'")
                If RBuscaBodega.RecordCount > 0 Then
                    LblBod.Caption = RBuscaBodega!Descripcion
                Else
                    LblBod.Caption = ""
                End If

End Sub

Private Sub TxtBod_DblClick()
   BTipoDeDocumento = False
   BProducto = False
   BBodegaDetalle = True
   BTransportistas = False
   FrameBuscar.Visible = True
   TxtBuscar.SetFocus
   DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Order by CodigoBodega")
   DataBuscar.Refresh
   DBGridBuscar.Refresh
   DBGridBuscar.Columns(1).Width = "4000"

End Sub

Private Sub TxtBod_GotFocus()
    TxtBod.SelStart = 0
    TxtBod.SelLength = Len(TxtBod.Text)
End Sub

Private Sub TxtBod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        BTipoDeDocumento = False
        BProducto = False
        BBodegaDetalle = True
        BTransportistas = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Order by CodigoBodega")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
    End If
End Sub

Private Sub TxtLin_Change()
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txtlin.Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                LblLin.Caption = RBuscaLinea!Descrip
            Else
                LblLin.Caption = ""
            End If
End Sub

Private Sub TxtLinea_Change()
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtLinea.Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                LblLinea.Caption = RBuscaLinea!Descrip
            Else
                LblLinea.Caption = ""
            End If

End Sub

Private Sub TxtNumDoc_GotFocus()
        TxtNumDoc.SelStart = 0
        TxtNumDoc.SelLength = Len(TxtNumDoc.Text)
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTipDoc_Change()
    Set RBuscaTipoDocumento = Db.OpenRecordset("Select Descripcion From Documentos Where CodigoDocumento = '" & TxtTipDoc.Text & "'")
    If RBuscaTipoDocumento.RecordCount > 0 Then
        LblDoc.Caption = RBuscaTipoDocumento!Descripcion
    Else
        LblDoc.Caption = ""
    End If
End Sub
Private Sub TxtTipDoc_DblClick()
   BTipoDeDocumento = True
   BProducto = False
   BBodegaDetalle = False
   BTransportistas = False
   FrameBuscar.Visible = True
   TxtBuscar.SetFocus
   DataBuscar.RecordSource = ("Select * from Documentos Order by CodigoDocumento")
   DataBuscar.Refresh
   DBGridBuscar.Refresh
   DBGridBuscar.Columns(1).Width = "4000"

End Sub
Private Sub TxtTipDoc_GotFocus()
    TxtTipDoc.SelStart = 0
    TxtTipDoc.SelLength = Len(TxtTipDoc.Text)
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
        
    If KeyAscii = 43 Then
        BTipoDeDocumento = True
        BProducto = False
        BBodegaDetalle = False
        BTransportistas = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select * from Documentos Order by CodigoDocumento")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
    End If
End Sub
Private Sub Txtbuscar_Change()
    'BODEGA
    If BBodegaDetalle = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Where CodigoBodega Like '" & TxtBuscar.Text & "*' Order by CodigoBodega")
            Else
                    DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Where CodigoBodega Like '*" & TxtBuscar.Text & "*' Order by CodigoBodega")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Where Descripcion Like '" & TxtBuscar.Text & "*' Order by CodigoBodega")
            Else
                    DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Where Descripcion Like '*" & TxtBuscar.Text & "*' Order by CodigoBodega")
            End If
        End If

    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Esp_Tec Like '" & TxtBuscar.Text & "*' And Activa = -1")
            Else
                    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Esp_Tec Like '*" & TxtBuscar.Text & "*' And Activa = -1")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, Envases, Size from FichaTecnica Where Descrip Like '" & TxtBuscar.Text & "*' And Activa = -1")
            Else
                    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, Envases, Size from FichaTecnica Where Descrip Like '*" & TxtBuscar.Text & "*' And Activa = -1")
            End If
        End If
    'Documentos
    ElseIf BTipoDeDocumento = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from Documentos Where CodigoDocumento Like '" & TxtBuscar.Text & "*' Order by CodigoDocumento")
            Else
                    DataBuscar.RecordSource = ("Select * from Documentos Where CodigoDocumento Like '*" & TxtBuscar.Text & "*' Order by CodigoDocumento")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from Documentos Where Descripcion Like '" & TxtBuscar.Text & "*' Order by CodigoDocumento")
            Else
                    DataBuscar.RecordSource = ("Select * from Documentos Where Descripcion Like '*" & TxtBuscar.Text & "*' Order by CodigoDocumento")
            End If
        End If
    
    End If
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
End Sub
Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub
Private Sub TxtCanPro_GotFocus()
    TxtCanPro.SelStart = 0
    TxtCanPro.SelLength = Len(TxtCanPro.Text)
End Sub
Private Sub TxtCanPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub TxtCodPro_Change()
                 Set RBuscaProducto = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCodPro.Text & "'")
                 If RBuscaProducto.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaProducto(0)
                 Else
                        TxtDesPro.Text = ""
                 End If
End Sub

Private Sub TxtCodPro_DblClick()
    BTipoDeDocumento = False
    BProducto = True
    BBodegaDetalle = False
    BTransportistas = False
    FrameBuscar.Visible = True
    TxtBuscar.SetFocus
    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1")
    DataBuscar.Refresh
    DBGridBuscar.Refresh
    DBGridBuscar.Columns(1).Width = "4000"
End Sub

Private Sub TxtCodPro_KeyPress(KeyAscii As Integer)
    'SI PRECIONA ENTER
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    'SI PRECIONA LA TECLA DE SIGNO +
    If KeyAscii = 43 Then
       BTipoDeDocumento = False
       BProducto = True
       BBodegaDetalle = False
       BTransportistas = False
       FrameBuscar.Visible = True
       TxtBuscar.SetFocus
       DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1")
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

Private Sub TxtDocIng_GotFocus()
        TxtDocIng.SelStart = 0
        TxtDocIng.SelLength = Len(TxtDocIng.Text)
End Sub

Private Sub TxtDocing_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtLin_GotFocus()
        Txtlin.SelStart = 0
        Txtlin.SelLength = Len(Txtlin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub TxtLinea_GotFocus()
        TxtLinea.SelStart = 0
        TxtLinea.SelLength = Len(TxtLinea.Text)
End Sub

Private Sub TxtLinea_KeyPress(KeyAscii As Integer)
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

Public Sub BotonesVisiblesEncabezado()
    If Bandera4 = True Then
         CmdAgregar.Visible = True
         CmdEditar.Visible = True
         CmdGrabar.Visible = True
         CmdCancelar.Visible = True
         CmdBorrar.Visible = True
         CmdBuscar.Visible = True
         CmdBuscar2.Visible = True
         CmdImprimir.Visible = True
         CmdSalida.Visible = True
    Else
         CmdAgregar.Visible = False
         CmdEditar.Visible = False
         CmdGrabar.Visible = False
         CmdCancelar.Visible = False
         CmdBorrar.Visible = False
         CmdBuscar.Visible = False
         CmdBuscar2.Visible = False
         CmdImprimir.Visible = False
         CmdSalida.Visible = False
    End If

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        Txttexto.Item(Index).SelStart = 0
        Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub TxtBodTra_Change()
            Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasProductoTerminado Where CodigoBodega = '" & TxtBodTra.Text & "'")
                If RBuscaBodega.RecordCount > 0 Then
                    LblBodTra.Caption = RBuscaBodega!Descripcion
                Else
                    LblBodTra.Caption = ""
                End If
End Sub

Private Sub TxtBodTra_DblClick()
            BTipoDeDocumento = False
            BProducto = False
            BBodegaDetalle = False
            BTransportistas = True
            FrameBuscar.Visible = True
            
            TxtBuscar.SetFocus
            DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Order by CodigoBodega")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"

End Sub

Private Sub TxtBodTra_GotFocus()
        TxtBodTra.SelStart = 0
        TxtBodTra.SelLength = Len(TxtBodTra.Text)
End Sub

Private Sub TxtBodTra_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BTipoDeDocumento = False
            BProducto = False
            BBodegaDetalle = False
            BTransportistas = True
            FrameBuscar.Visible = True
            
            TxtBuscar.SetFocus
            DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Order by CodigoBodega")
            DataBuscar.Refresh
            DBGridBuscar.Refresh
            DBGridBuscar.Columns(1).Width = "4000"
        End If

End Sub

