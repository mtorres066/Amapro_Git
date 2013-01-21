VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form EntradasProductoTerminadoXEvadeva 
   BackColor       =   &H00FF8080&
   Caption         =   "Entradas Producto Terminado"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   Icon            =   "EntradasProductoTerminadoXEvadeva.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11955
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
         Picture         =   "EntradasProductoTerminadoXEvadeva.frx":08CA
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
         Bindings        =   "EntradasProductoTerminadoXEvadeva.frx":293C
         Height          =   7335
         Left            =   120
         OleObjectBlob   =   "EntradasProductoTerminadoXEvadeva.frx":2955
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
      Connect         =   "pwd=metal"
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin TabDlg.SSTab TabEntradas 
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
      TabPicture(0)   =   "EntradasProductoTerminadoXEvadeva.frx":332D
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameEncabezado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "EntradasProductoTerminadoXEvadeva.frx":377F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridDetalleEntradas"
      Tab(1).Control(1)=   "FrameDetalle"
      Tab(1).ControlCount=   2
      Begin MSDBGrid.DBGrid DBGridDetalleEntradas 
         Bindings        =   "EntradasProductoTerminadoXEvadeva.frx":3A99
         Height          =   4575
         Left            =   -74760
         OleObjectBlob   =   "EntradasProductoTerminadoXEvadeva.frx":3ABB
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   2640
         Width           =   11415
      End
      Begin VB.Frame FrameDetalle 
         Caption         =   "Detalle Entradas De Producto Terminado"
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
         TabIndex        =   45
         Top             =   720
         Width           =   11685
         Begin VB.Frame FrameDetalleCompras 
            Enabled         =   0   'False
            Height          =   1575
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   11415
            Begin VB.TextBox TxtBarra 
               Appearance      =   0  'Flat
               DataField       =   "Barra"
               DataSource      =   "DataDetalleEntradas"
               Height          =   288
               Left            =   120
               MaxLength       =   35
               TabIndex        =   78
               Top             =   840
               Visible         =   0   'False
               Width           =   1815
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
               Left            =   3960
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   480
               Width           =   5652
            End
            Begin VB.TextBox TxtCodPro 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "FichaTecnica"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   2040
               MaxLength       =   12
               TabIndex        =   48
               Top             =   480
               Width           =   1812
            End
            Begin VB.TextBox TxtDocDet 
               Appearance      =   0  'Flat
               DataField       =   "Documento"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   5640
               TabIndex        =   63
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox TxtBod 
               Appearance      =   0  'Flat
               DataField       =   "Bodega"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   6720
               MaxLength       =   3
               TabIndex        =   55
               Top             =   1200
               Width           =   435
            End
            Begin VB.TextBox TxtTar 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Tarima"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   2760
               TabIndex        =   50
               Top             =   840
               Width           =   1155
            End
            Begin VB.TextBox TxtLin 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Linea"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   6720
               MaxLength       =   2
               TabIndex        =   52
               Top             =   840
               Width           =   435
            End
            Begin VB.TextBox TxtBat 
               Appearance      =   0  'Flat
               DataField       =   "Batch"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   4680
               TabIndex        =   54
               Top             =   1200
               Width           =   1155
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
               TabIndex        =   62
               Top             =   1200
               Width           =   1815
            End
            Begin VB.TextBox TxtCal 
               Appearance      =   0  'Flat
               DataField       =   "Calidad"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   2760
               MaxLength       =   1
               TabIndex        =   53
               Top             =   1200
               Width           =   1155
            End
            Begin VB.TextBox TxtOrd 
               Appearance      =   0  'Flat
               DataField       =   "OrdenProduccion"
               DataSource      =   "DataDetalleEntradas"
               Height          =   288
               Left            =   120
               MaxLength       =   15
               TabIndex        =   47
               Top             =   480
               Width           =   1812
            End
            Begin MSMask.MaskEdBox MskFecPro 
               DataField       =   "FechaProduccion"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   4680
               TabIndex        =   51
               Top             =   840
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   8438015
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TxtCanPro 
               DataField       =   "Cantidad"
               DataSource      =   "DataDetalleEntradas"
               Height          =   285
               Left            =   9720
               TabIndex        =   49
               Top             =   480
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
               Height          =   252
               Index           =   0
               Left            =   2040
               TabIndex        =   76
               Top             =   240
               Width           =   1572
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
               Height          =   252
               Left            =   3960
               TabIndex        =   75
               Top             =   240
               Width           =   1572
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
               TabIndex        =   74
               Top             =   240
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
               TabIndex        =   73
               Top             =   840
               Width           =   585
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
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
               Index           =   8
               Left            =   4080
               TabIndex        =   72
               Top             =   840
               Width           =   540
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
               Left            =   6000
               TabIndex        =   71
               Top             =   840
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
               TabIndex        =   70
               Top             =   1200
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
               TabIndex        =   69
               Top             =   1200
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
               Index           =   6
               Left            =   2040
               TabIndex        =   68
               Top             =   1200
               Width           =   645
            End
            Begin VB.Label LblLin2 
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
               Left            =   7320
               TabIndex        =   67
               Top             =   840
               Width           =   3975
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
               Left            =   7320
               TabIndex        =   66
               Top             =   1200
               Width           =   3975
            End
            Begin VB.Label Label3 
               BackColor       =   &H80000004&
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
               Height          =   252
               Index           =   1
               Left            =   120
               TabIndex        =   65
               Top             =   240
               Width           =   852
            End
         End
         Begin VB.CommandButton CmdAgregar2 
            Caption         =   "A&gregar"
            Height          =   495
            Left            =   240
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":58C4
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   6600
            Visible         =   0   'False
            Width           =   1700
         End
         Begin VB.CommandButton CmdGrabar2 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   3840
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":5DF6
            Style           =   1  'Graphical
            TabIndex        =   58
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
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":6328
            TabIndex        =   61
            Top             =   6600
            Visible         =   0   'False
            Width           =   1815
         End
         Begin VB.CommandButton CmdCancelar2 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5640
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":685A
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
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":6D8C
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   6600
            Visible         =   0   'False
            Width           =   1700
         End
         Begin VB.CommandButton CmdEditar2 
            Caption         =   "Editar"
            Height          =   495
            Left            =   2040
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":72BE
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   6600
            Visible         =   0   'False
            Width           =   1700
         End
      End
      Begin VB.Frame FrameEncabezado 
         Caption         =   "Encabezado Entradas De Producto Terminado"
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
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   11655
         Begin VB.Frame FrameCompras 
            Enabled         =   0   'False
            Height          =   2895
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   11415
            Begin VB.TextBox TxtDocIng 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               DataField       =   "Documento"
               DataSource      =   "DataEntradas"
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
               Left            =   3960
               TabIndex        =   22
               Top             =   240
               Width           =   1335
            End
            Begin VB.TextBox TxtLib 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Liberado"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtBodega 
               Appearance      =   0  'Flat
               DataField       =   "Bodega"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   25
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox TxtReq 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               DataField       =   "Requerido"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TxtBatch 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               DataField       =   "Batch"
               DataSource      =   "DataEntradas"
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
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
            Begin VB.OptionButton OptAut 
               Caption         =   "Captura Automatica"
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
               Left            =   6000
               TabIndex        =   30
               Top             =   240
               Value           =   -1  'True
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.OptionButton OptMan 
               Caption         =   "Captura Manual"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   6840
               TabIndex        =   29
               Top             =   240
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox TxtObs 
               Appearance      =   0  'Flat
               DataField       =   "Observaciones"
               DataSource      =   "DataEntradas"
               Height          =   285
               Left            =   1560
               MaxLength       =   50
               TabIndex        =   28
               Top             =   2400
               Width           =   6855
            End
            Begin VB.CheckBox ChkProInt 
               Caption         =   "Produccion"
               DataField       =   "ProduccionInterna"
               DataSource      =   "DataEntradas"
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
               Left            =   1560
               TabIndex        =   23
               Top             =   600
               Width           =   1455
            End
            Begin VB.CheckBox ChkProLib 
               Caption         =   "Produccion Liberada"
               DataField       =   "ProduccionLiberada"
               DataSource      =   "DataEntradas"
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
               Left            =   1560
               TabIndex        =   24
               Top             =   960
               Width           =   2295
            End
            Begin VB.TextBox TxtLinea 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               DataField       =   "Linea"
               DataSource      =   "DataEntradas"
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
               DataField       =   "FechaEntrada"
               DataSource      =   "DataEntradas"
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
               Caption         =   "Fecha Entrada"
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
               TabIndex        =   44
               Top             =   240
               Width           =   1260
            End
            Begin VB.Label Label7 
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
               Left            =   2880
               TabIndex        =   43
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label6 
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
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   42
               Top             =   1320
               Width           =   975
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
               Left            =   2880
               TabIndex        =   41
               Top             =   1320
               Width           =   5535
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
               TabIndex        =   40
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
               TabIndex        =   39
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
               Left            =   8760
               TabIndex        =   38
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
               Left            =   8880
               TabIndex        =   37
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
               TabIndex        =   36
               Top             =   2400
               Width           =   1275
            End
            Begin VB.Label Label6 
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
               Height          =   255
               Index           =   11
               Left            =   120
               TabIndex        =   35
               Top             =   2040
               Width           =   615
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
               TabIndex        =   34
               Top             =   2040
               Width           =   5535
            End
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&AGREGAR"
            Height          =   700
            Left            =   120
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":77F0
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3240
            Width           =   1200
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "&GRABAR"
            Enabled         =   0   'False
            Height          =   700
            Left            =   2760
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":7D22
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   3240
            Width           =   1200
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&CANCELAR"
            Enabled         =   0   'False
            Height          =   700
            Left            =   4080
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":8254
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   3240
            Width           =   1200
         End
         Begin VB.CommandButton CmdBorrar 
            Caption         =   "&BORRAR"
            Height          =   700
            Left            =   5400
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":8786
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   3240
            Width           =   1200
         End
         Begin VB.CommandButton CmdSalida 
            Appearance      =   0  'Flat
            Height          =   700
            Left            =   10920
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":8CB8
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Salida"
            Top             =   3240
            Width           =   600
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "B&USCAR"
            Height          =   700
            Left            =   6720
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":AD2A
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   3240
            Width           =   1200
         End
         Begin VB.CommandButton CmdEditar 
            Caption         =   "&EDITAR"
            Height          =   700
            Left            =   1440
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":B25C
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3240
            Width           =   1200
         End
         Begin VB.CommandButton CmdImprimir 
            Caption         =   "&Imprimir Entrad"
            Height          =   700
            Left            =   8040
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":B78E
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   3240
            Width           =   1300
         End
         Begin VB.CommandButton CmdImprimirAmarillas 
            Caption         =   "Imprimir Amarillas"
            Height          =   700
            Left            =   9480
            Picture         =   "EntradasProductoTerminadoXEvadeva.frx":BCC0
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3240
            Width           =   1320
         End
      End
   End
   Begin VB.Data DataDetalleEntradas 
      Caption         =   "Detalle Entradas Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\erick\Amapro Metalenvases\metalenvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DetalleEntradasProductoTerminado"
      Top             =   8160
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Data DataEntradas 
      Caption         =   "Entradas De Producto Terminado"
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Erick\Amapro\METALENVASES.MDB"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   465
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EncabezadoEntradasProductoTerminado"
      Top             =   8040
      Width           =   11655
   End
End
Attribute VB_Name = "EntradasProductoTerminadoXEvadeva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As Double
Dim VDocumentoDetalle As Double
Dim VCantidad As Double
Dim VCodigoProducto As String
Dim VCantidadProducto As Double
Dim VBodega As String
Dim VBatch As Double
Dim VClasificacion As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim BBodega As Boolean
Dim BProducto As Boolean
Dim BBodegaDetalle As Boolean
Dim BProduccionInterna As Boolean
Dim BProduccionLiberada As Boolean
Dim BLineas As Boolean

Dim RBuscaProducto As Recordset
Dim RMaximo As Recordset
Dim RBuscaBodega As Recordset
Dim RBuscaDetalle As Recordset
Dim RBuscaEncabezado As Recordset
Dim RBuscaProduccion As Recordset
Dim RBuscaTarima As Recordset
Dim RCuentaTarimas As Recordset
Dim RBuscaLinea As Recordset
Dim RBuscaFichaOrden As Recordset


Dim VUltimaFichaTecnica As String
Dim VUltimaEnvases As Long
Dim VUltimaFecha As String
Dim VUltimaTarima As Long
Dim VUltimaLinea As String
Dim VUltimaCalidad As String
Dim VLinea As String
Dim VUltimaOrden As String



Sub Botones1()
    If Bandera = True Then
         FrameCompras.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdSalida.Enabled = False
         CmdImprimir.Enabled = False
         CmdImprimirAmarillas.Enabled = False
         DataEntradas.Visible = False
    Else
         FrameCompras.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdSalida.Enabled = True
         CmdImprimir.Enabled = True
         CmdImprimirAmarillas.Enabled = True
         DataEntradas.Visible = True
         
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


Private Sub ChkProInt_Click()
        If ChkProInt.Value = 1 Then
            ChkProLib.Value = 0
        End If
End Sub

Private Sub ChkProInt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub ChkProLib_Click()
        If ChkProLib.Value = 1 Then
            ChkProInt.Value = 0
        End If
End Sub

Private Sub ChkProLib_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub CmdAgregar2_Click()
On Error Resume Next
    'AGREGA DATOS
    DataDetalleEntradas.Recordset.AddNew
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
    Bandera2 = True
    Botones2
    DBGridDetalleEntradas.Enabled = False
    
    TxtDocDet.Text = VDocumento
    
    'ASIGNA LOS DATOS DEL ENCABEZADO
    TxtBod.Text = VBodega
    TxtBat.Text = VBatch
    'ASIGNA LOS ULTIMOS DATOS DIGITADOS
    TxtCodPro.Text = VUltimaFichaTecnica
    TxtCanPro.Text = VUltimaEnvases
    Txtlin.Text = VUltimaLinea
    MskFecPro.Text = VUltimaFecha
    TxtTar.Text = VUltimaTarima
    TxtCal.Text = VUltimaCalidad
    TxtOrd.Text = VUltimaOrden
    
   
    TxtOrd.SetFocus
    TxtDesPro.Text = ""
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            If GBorrar = True Then
                'NO HACE NADA PORQUE SI TIENE ACCESO
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
                        'BORRA EL ENCABEZADO DE EL PEDIDO
                        DataEntradas.Recordset.Delete
                        If Err <> 0 Then
                            MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                            Exit Sub
                        End If
                        DataEntradas.Recordset.MoveLast
                MousePointer = 0
            End If
            If DataEntradas.Recordset.EOF Then
                DataEntradas.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
            
            'ASIGANMOS A UNA VARIABLE EL DOCUMENTO DETALLE
            VDocumentoDetalle = TxtDocDet.Text
            VBodega = TxtBod.Text
            VCodigoProducto = TxtCodPro.Text
            VCantidad = TxtCanPro.Text
    
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                                        
                   'BORRA EL DETALLE DE LA ENTRADA
                    DataDetalleEntradas.Recordset.Delete
                    
                    If Err <> 0 Then
                       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                       Exit Sub
                    End If
                    'SELECCIONA TODOS LOS DETALLES DE LA ENTRADAS
                    DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasProductoTermina where documento = " & VDocumentoDetalle & " order By Linea, Batch, Tarima")
                    DataDetalleEntradas.Refresh
                    DBGridDetalleEntradas.Refresh
                MousePointer = 0
            End If
  
            If DataEntradas.Recordset.EOF Then
                DataEntradas.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
End Sub

Private Sub CmdBuscar_Click()
    mensaje = InputBox("Documento a Buscar")
    If mensaje = "" Then
    Else
        DataEntradas.Recordset.FindFirst ("Documento = " & mensaje)
    End If
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
    'CANCELA LOS CAMBIOS
    DataEntradas.Recordset.CancelUpdate
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error "
        Err.Clear
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    FrameDetalle.Visible = True
    DBGridDetalleEntradas.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    'CANCELA LOS DATOS CAMBIADOS Y GRABA LOS DATOS COMO ESTABAN
    DataDetalleEntradas.Recordset.CancelUpdate
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    DBGridDetalleEntradas.Enabled = True
    Bandera2 = False
    Botones2

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
    If GEditar = True Then
        'NO HACE NADA PORQUE SI TIENE ACCESO
    ElseIf TxtEstado.Text = "LIBERADO" Then
        'VERIFICA SI YA FUE LIBERADA LA ENTRADA
        MsgBox "Esta Documento No Se Puede EDITAR Porque Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
        Exit Sub
    End If
    
    'EDITA EL REGISTRO
    DataEntradas.Recordset.Edit
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
    DBGridDetalleEntradas.Visible = False
    
End Sub


Private Sub CmdEditar2_Click()
    On Error Resume Next
    'AGREGA DATOS
    DataDetalleEntradas.Recordset.Edit
    
    If Err <> 0 Then
       MsgBox Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
       Exit Sub
    End If
    
    Bandera2 = True
    Botones2
    DBGridDetalleEntradas.Enabled = False
        
    TxtCodPro.SetFocus
    

End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
    
    'GUARDA VARIABLES
    VCantidad = TxtCanPro.Text
    VCodigoProducto = TxtCodPro.Text
    
    VUltimaFichaTecnica = TxtCodPro.Text
    VUltimaEnvases = TxtCanPro.Text
    VUltimaLinea = Txtlin.Text
    VUltimaFecha = MskFecPro.Text
    VUltimaTarima = Val(TxtTar.Text) + 1
    VUltimaCalidad = TxtCal.Text
    VUltimaOrden = TxtOrd.Text
    VBodega = TxtBod.Text
        
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
       MsgBox "Fecha De Boleta Incorrecta", vbOKOnly + vbCritical, "Error"
       MskFecPro.SetFocus
       Exit Sub
    End If
    
        
            'REVISA SI LA TARIMA EXISTE EN LA PRODUCCION ELEGIDA
            'PRODUCCION INTERNA
            If BProduccionInterna = True Then
                Set RBuscaProduccion = Db.OpenRecordset("Select * From Produccion Where Fec_Prd = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTar.Text & " And Linea = '" & Txtlin.Text & "' And Esp_Tec = '" & TxtCodPro.Text & "'")
            'PRODUCCION LIBERADA
            ElseIf BProduccionLiberada = True Then
                Set RBuscaProduccion = Db.OpenRecordset("Select * From ProduccionLiberada Where Fec_Prd = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTar.Text & " And Linea = '" & Txtlin.Text & "' And Esp_Tec = '" & TxtCodPro.Text & "'")
            End If
            
            
                If BProduccionInterna = True Then
                    If RBuscaProduccion.RecordCount > 0 Then
                    Else
                        MsgBox "La Tarima No Existe En La Produccion Interna De Evadeva", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                ElseIf BProduccionLiberada = True Then
                    If RBuscaProduccion.RecordCount > 0 Then
                    Else
                        MsgBox "La Tarima No Existe En La Produccion Liberada De Evadeva", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                End If
            
            
    
    
    
    
    'ASIGNAMOS A LA CANTIDAD DE SALDO LA CANTIDAD QUE ESTA ENTRANDO
    DataDetalleEntradas.Recordset!Saldo = VCantidad
    
    'ASIGNA LA BARRA
    TxtBarra.Text = Format(MskFecPro.Text, "dd-mm-yyyy") & "-" & Txtlin.Text & "-" & TxtCodPro.Text & "-" & TxtTar.Text
                
    'GRABA DATOS
    DataDetalleEntradas.Recordset.Update
        
    If Err <> 0 And Err <> 3022 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    ElseIf Err = 3022 Then
        MsgBox "La Tarima Ya Existe", vbOKOnly + vbExclamation, "Informacion"
        Exit Sub
    End If
        
    Bandera2 = False
    Botones2
         
    'ACTUALIZA EL GRID DE DETALLE PARA QUE SOLO APARESCAN LOS DETALLES DE LA FACTURA QUE SE ESTA GRABANDO
    DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasProductoTermina where Documento = " & VDocumento & " Order by Linea, Batch, Tarima")
    DataDetalleEntradas.Refresh
    DBGridDetalleEntradas.Refresh
           
    DBGridDetalleEntradas.Enabled = True
    TxtDesPro.Text = ""
    CmdAgregar2.SetFocus
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
    
    DataEntradas.Recordset.AddNew
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
    
    'ASIGNA VALOR AL CHECK DE PRODUCCION INTERNA
    ChkProInt.Value = 1
    
    TxtCueTar.Text = ""
    
    'BUSCA EL DOCUMENTO MAXIMO Y LE AGREGA UNO MAS
    Set RMaximo = Db.OpenRecordset("Select max(Documento) from EncabezadoEntradasProductoTerm")
        If RMaximo.RecordCount > 0 Then
            If IsNull(RMaximo(0)) Then
                TxtDocIng.Text = "1"
            Else
                TxtDocIng.Text = Val(RMaximo(0)) + 1
            End If
        End If
        
    FrameDetalle.Visible = False
    DBGridDetalleEntradas.Visible = False

End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next

MousePointer = 11

    VDocumento = TxtDocIng.Text
    VBodega = TxtBodega.Text
    VBatch = TxtBatch.Text
    VLinea = TxtLinea.Text
    
    'ASIGNA VALORES A LAS VARIABLES PARA PODER CONTROLAR DE DONDE VIENEN LAS TARIMAS
    BProduccionInterna = ChkProInt.Value
    BProduccionLiberada = ChkProLib.Value
    
    
    'REVISA SI ELIGIO ALGUNA PRODUCCION
    If (BProduccionInterna = True Or BProduccionLiberada = True) Then
        OptAut.Value = True
        OptMan.Value = False
    Else
        OptMan.Value = True
        OptAut.Value = False
    End If
    
    'REVISA LA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
        MskFec.SetFocus
        Exit Sub
    End If
            
    'REVISA SI ES NUMERICO EL DOCUMENTO DE RECEPCION
    If Not IsNumeric(TxtDocIng.Text) Then
        MsgBox "Numero De Documento Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        TxtDocIng.SetFocus
        Exit Sub
    End If
    
    'REVISA EL BATCH
    If Not IsNumeric(TxtBatch.Text) Then
        MsgBox "Numero De Batch Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        TxtBatch.SetFocus
        Exit Sub
    End If
    
    If TxtLinea.Text = "" Then
        MsgBox "Linea No Puede Estar En Blanco", vbOKOnly + vbInformation, "Informacion"
        TxtLinea.SetFocus
        Exit Sub
    End If
    
    
               
    'GRABA DATOS
    DataEntradas.Recordset.Update
    
    If Err = 3022 Then
        MsgBox "Documento Ya Existe ", vbOKOnly + vbCritical, "Informacion"
        TxtDocIng.SetFocus
        Exit Sub
    ElseIf Err <> 0 And Err <> 3022 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    'SELECCIONA TODAS LAS TARIMAS DE PRODUCCION DE ACUERDO AL BATCH Y DE DONDE PROVIENEN LAS TARIMAS
    'QUE PUEDE SER POR PRODUCCION INTERNA, PRODUCCION LIBERADA O PRODUCCION EXTERNA
    If OptAut.Value = True Then
        If BProduccionInterna = True Then
            Set RBuscaProduccion = Db.OpenRecordset("Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From Produccion Where Batch = " & VBatch & " And Linea = '" & VLinea & "' And (Calidad = 'A' or Calidad = 'R' Or Calidad = 'I')")
        ElseIf BProduccionLiberada = True Then
            Set RBuscaProduccion = Db.OpenRecordset("Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From ProduccionLiberada Where Batch = " & VBatch & " And Linea = '" & VLinea & "'")
        Else
            Set RBuscaProduccion = Db.OpenRecordset("Select * From Lineas")
        End If
                    
            If RBuscaProduccion.RecordCount > 0 Then
                
                'INICIALIZA EL RECORDSET PARA AGREGAR DATOS
                Set RBuscaDetalle = Db.OpenRecordset("Select * From DetalleEntradasProductoTermina")
                    
                    'CREA UN CICLO CON LOS DATOS DE PRODUCCION DEL BATCH
                    Do Until RBuscaProduccion.EOF
                        
                        'BUSCAMOS SI EXISTE LA TARIMA
                        Set RBuscaTarima = Db.OpenRecordset("Select FichaTecnica From DetalleEntradasProductoTermina Where FichaTecnica = '" & RBuscaProduccion!Esp_Tec & "' And Tarima = " & RBuscaProduccion!Tarima & " And FechaProduccion = #" & Format(RBuscaProduccion!fec_prd, "mm/dd/yyyy") & "# And Linea = '" & RBuscaProduccion!Linea & "'")
                        
                            'SI ENCUENTRA LA TARIMA LA EDITA
                            If RBuscaTarima.RecordCount > 0 Then
                                MsgBox "Tarima " & RBuscaProduccion!Tarima & " Ya Fue Ingresada", vbOKOnly + vbInformation, "Revise Por Favor"
                            'AGREGA AL DETALLE DE LA ENTRADA DE PRODUCTO LO QUE SE CAPTURO EN PRODUCCION
                            Else
                                    RBuscaDetalle.AddNew
                                        RBuscaDetalle!Documento = VDocumento
                                        RBuscaDetalle!FichaTecnica = RBuscaProduccion!Esp_Tec
                                        RBuscaDetalle!Cantidad = RBuscaProduccion!Envases
                                        RBuscaDetalle!Tarima = RBuscaProduccion!Tarima
                                        RBuscaDetalle!FechaProduccion = RBuscaProduccion!fec_prd
                                        RBuscaDetalle!Linea = RBuscaProduccion!Linea
                                        RBuscaDetalle!Batch = VBatch
                                        RBuscaDetalle!Bodega = VBodega
                                        RBuscaDetalle!Saldo = RBuscaProduccion!Envases
                                        RBuscaDetalle!Salidas = 0
                                        RBuscaDetalle!Calidad = RBuscaProduccion!Calidad
                                        RBuscaDetalle!OrdenProduccion = RBuscaProduccion!Orden
                                        RBuscaDetalle!Barra = RBuscaProduccion!Barra
                                    RBuscaDetalle.Update
                            End If
                                    
                                    If Err <> 0 Then
                                 '       MsgBox Err.Number & Err.Description & "Ojo Tarima " & RBuscaProduccion!Tarima & " Ya Existe, No Se Grabara, Pero Revise las Tarimas Por Favor", vbOKOnly + vbInformation, "Informacion"
                                    End If
                        'SE MUEVE AL SIGUIENTE REGISTRO
                        RBuscaProduccion.MoveNext
                    Loop
            End If
    'SI ES MANUAL NO HACE NADA EL LA TIENE QUE AGREGAR
    Else
    
    End If
            
    'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
    DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasProductoTermina where Documento = " & VDocumento & " Order By Linea, Batch, Tarima")
    DataDetalleEntradas.Refresh
    DBGridDetalleEntradas.Refresh
            
    'MUEVE EL RECORDSET A EL DOCUMENTO ACTUAL PARA QUE SE ACTUALIZEN LOS CAMBIOS
    DataEntradas.Recordset.FindFirst ("Documento = " & VDocumento)
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Visible = True
    FrameDetalle.Enabled = True
    FrameEncabezado.Enabled = False
    'VISUALIZA EL GRID DE DETALEE
    DBGridDetalleEntradas.Visible = True
    
    'HABILITA LAS COLUMNAS PARA PODER MODIFICARLAS PERO SOLO LA UBICACION
    DBGridDetalleEntradas.AllowUpdate = True
    DBGridDetalleEntradas.AllowDelete = True
        
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = False
    BotonesVisiblesEncabezado
    
    DataEntradas.Visible = False
    
    TabEntradas.Tab = 1
       
    CmdAgregar2.SetFocus
    
MousePointer = 0
    
End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDocIng.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
        
                CrReportes.SelectionFormula = "{EncabezadoEntradasProductoTerm.Documento} = " & TxtDocIng.Text
                CrReportes.ReportFileName = App.Path & "\EntradasProductoTerminado.rpt"
                CrReportes.Action = 1
                
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                    Exit Sub
                End If
MousePointer = 0

End Sub

Private Sub CmdImprimirAmarillas_Click()
On Error Resume Next
MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDocIng.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
        
                CrReportes.SelectionFormula = "{EncabezadoEntradasProductoTerm.Documento} = " & TxtDocIng.Text
                CrReportes.ReportFileName = App.Path & "\BoletaAmarilla.rpt"
                CrReportes.Action = 1
                
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                    Exit Sub
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
    
    'DESHABILITA LAS COLUMNAS PARA PODER MODIFICARLAS PERO SOLO LA UBICACION
    DBGridDetalleEntradas.AllowUpdate = True
    DBGridDetalleEntradas.AllowDelete = True
    
    DataEntradas.Visible = True
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesVisiblesEncabezado
    
    TabEntradas.Tab = 0
  
End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub



Private Sub DataDetalleEntradas_Reposition()
        If IsNumeric(TxtDocDet.Text) Then
            'CUENTA CUANTAS TARIMAS TIENE EL DOCUMENTO
            Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From DetalleEntradasProductoTermina Where Documento = " & TxtDocDet.Text)
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


Private Sub DataEntradas_Error(DataErr As Integer, Response As Integer)
    On Error Resume Next
        If Err <> 0 Then
            'MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical, "Error"
        End If
End Sub

Private Sub DataEntradas_Reposition()
    If IsNumeric(TxtDocIng.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
        DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasProductoTermina where Documento = " & TxtDocIng.Text & " Order by Linea, Batch, Tarima")
        DataDetalleEntradas.Refresh
        DBGridDetalleEntradas.Refresh
    End If
End Sub


Private Sub DBGridBuscar_DblClick()
    'BODEGA
    If BBodega = True Then
        TxtBodega.Text = DBGridBuscar.Columns(0)
        TxtBodega.SetFocus
    'BODEGA DETALLE
    ElseIf BBodegaDetalle = True Then
        TxtBod.Text = DBGridBuscar.Columns(0)
        TxtBod.SetFocus
    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        TxtCodPro.Text = DBGridBuscar.Columns(0)
        TxtCodPro.SetFocus
    'LINEAS
    ElseIf BLineas = True Then
        Txtlin.Text = DBGridBuscar.Columns(0)
        Txtlin.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
    'BODEGA
    If BBodega = True Then
        TxtBodega.Text = DBGridBuscar.Columns(0)
        TxtBodega.SetFocus
    'BODEGA DETALLE
    ElseIf BBodegaDetalle = True Then
        TxtBod.Text = DBGridBuscar.Columns(0)
        TxtBod.SetFocus
    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        TxtCodPro.Text = DBGridBuscar.Columns(0)
        TxtCodPro.SetFocus
    'LINEAS
    ElseIf BLineas = True Then
        Txtlin.Text = DBGridBuscar.Columns(0)
        Txtlin.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End If

End Sub

Private Sub DBGridDetalleEntradas_BeforeUpdate(Cancel As Integer)
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        End If
    
End Sub

Private Sub Form_Activate()
    If IsNumeric(TxtDocIng.Text) Then
        'SELECCIONA TODOS LOS DETALLES DE EL PEDIDO
        DataDetalleEntradas.RecordSource = ("Select * from DetalleEntradasProductoTermina where Documento = " & TxtDocIng.Text & " Order by Linea, Tarima")
        DataDetalleEntradas.Refresh
        DBGridDetalleEntradas.Refresh
    End If
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
End Sub

Private Sub Form_Load()
    DataEntradas.ConnectionString = GTipoProveedor
    DataDetalleEntradas.ConnectionString = GTipoProveedor
    DataBuscar.ConnectionString = GTipoProveedor
    
    DataEntradas.Refresh
    DataDetalleEntradas.Refresh
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

Private Sub OptAut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub OptMan_KeyPress(KeyAscii As Integer)
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
   TxtBuscar.Visible = True
   OptDescripcion.Visible = True
   OptCodigo.Visible = True
   FrameTipos.Visible = True
   BBodega = False
   BProducto = False
   BBodegaDetalle = True
   BLineas = False
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
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        FrameTipos.Visible = True
        BBodega = False
        BProducto = False
        BBodegaDetalle = True
        BLineas = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Order by CodigoBodega")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
    End If
End Sub
Private Sub TxtBodega_Change()
    Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasProductoTerminado Where CodigoBodega = '" & TxtBodega.Text & "'")
    If RBuscaBodega.RecordCount > 0 Then
        LblBodega.Caption = RBuscaBodega!Descripcion
    Else
        LblBodega.Caption = ""
    End If
End Sub
Private Sub TxtBodega_DblClick()
   TxtBuscar.Visible = True
   OptDescripcion.Visible = True
   OptCodigo.Visible = True
   FrameTipos.Visible = True
   BBodega = True
   BProducto = False
   BBodegaDetalle = False
   BLineas = False
   FrameBuscar.Visible = True
   TxtBuscar.SetFocus
   DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Order by CodigoBodega")
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
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        FrameTipos.Visible = True
        BBodega = True
        BProducto = False
        BBodegaDetalle = False
        BLineas = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select * from BodegasProductoTerminado Order by CodigoBodega")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
    End If
End Sub
Private Sub Txtbuscar_Change()
    'BODEGA
    If (BBodega = True Or BBodegaDetalle = True) Then
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
                    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Esp_Tec Like '" & TxtBuscar.Text & "*' And Activa = -1 Order by Esp_Tec")
            Else
                    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Esp_Tec Like '*" & TxtBuscar.Text & "*' And Activa = -1 Order by Esp_Tec")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Descrip Like '" & TxtBuscar.Text & "*' And Activa = -1 Order by Esp_Tec")
            Else
                    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Descrip Like '*" & TxtBuscar.Text & "*' And Activa = -1 Order by Esp_Tec")
            End If
        End If
        'LINEAS
    ElseIf BLineas = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from Lineas Where Linea Like '" & TxtBuscar.Text & "*' Order by Linea")
            Else
                    DataBuscar.RecordSource = ("Select * from Lineas Where Linea Like '*" & TxtBuscar.Text & "*' Order by Linea")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If OptPalIni.Value = True Then
                    DataBuscar.RecordSource = ("Select * from Lineas Where Descrip Like '" & TxtBuscar.Text & "*' Order by Linea")
            Else
                    DataBuscar.RecordSource = ("Select * from Lineas Where Descrip Like '*" & TxtBuscar.Text & "*' Order by Linea")
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

Private Sub TxtCal_GotFocus()
        TxtCal.SelStart = 0
        TxtCal.SelLength = Len(TxtCal.Text)
End Sub

Private Sub TxtCal_KeyPress(KeyAscii As Integer)
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
    TxtBuscar.Visible = True
    OptDescripcion.Visible = True
    OptCodigo.Visible = True
    FrameTipos.Visible = True
    BBodega = False
    BProducto = True
    BBodegaDetalle = False
    BLineas = False
    FrameBuscar.Visible = True
    TxtBuscar.SetFocus
    DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1 Order by Esp_Tec")
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
       TxtBuscar.Visible = True
       OptDescripcion.Visible = True
       OptCodigo.Visible = True
       FrameTipos.Visible = True
       BBodega = False
       BProducto = True
       BBodegaDetalle = False
       BLineas = False
       FrameBuscar.Visible = True
       TxtBuscar.SetFocus
       DataBuscar.RecordSource = ("Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1 Order by Esp_Tec")
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
Private Sub TxtDocing_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtLin_Change()
        Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & Txtlin.Text & "'")
            If RBuscaLinea.RecordCount > 0 Then
                LblLin2.Caption = RBuscaLinea!Descrip
            Else
                LblLin2.Caption = ""
            End If

End Sub

Private Sub Txtlin_DblClick()
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        FrameTipos.Visible = True
        BBodega = False
        BProducto = False
        BBodegaDetalle = False
        BLineas = True
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select * from Lineas")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
End Sub

Private Sub TxtLin_GotFocus()
        Txtlin.SelStart = 0
        Txtlin.SelLength = Len(Txtlin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        FrameTipos.Visible = True
        BBodega = False
        BProducto = False
        BBodegaDetalle = False
        BLineas = True
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        DataBuscar.RecordSource = ("Select * from Lineas")
        DataBuscar.Refresh
        DBGridBuscar.Refresh
        DBGridBuscar.Columns(1).Width = "4000"
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

Private Sub TxtLinea_GotFocus()
        TxtLinea.SelStart = 0
        TxtLinea.SelLength = Len(TxtLinea.Text)
End Sub

Private Sub TxtLinea_KeyPress(KeyAscii As Integer)
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

Private Sub TxtOrd_GotFocus()
        TxtOrd.SelStart = 0
        TxtOrd.SelLength = Len(TxtOrd.Text)
End Sub

Private Sub TxtOrd_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtOrd_LostFocus()
        'ORDEN EN DETALLE DE PRODUCCION
                Set RBuscaFichaOrden = Db.OpenRecordset("Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtOrd.Text & "'")
                    If RBuscaFichaOrden.RecordCount > 0 Then
                        TxtCodPro.Text = RBuscaFichaOrden!FichaTecnica
                    Else
                        TxtCodPro.Text = ""
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
         CmdImprimir.Visible = True
         CmdImprimirAmarillas.Visible = True
         CmdSalida.Visible = True
    Else
         CmdAgregar.Visible = False
         CmdEditar.Visible = False
         CmdGrabar.Visible = False
         CmdCancelar.Visible = False
         CmdBorrar.Visible = False
         CmdBuscar.Visible = False
         CmdImprimir.Visible = False
         CmdImprimirAmarillas.Visible = False
         CmdSalida.Visible = False
    End If

End Sub
