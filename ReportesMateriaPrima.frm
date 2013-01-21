VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form ReportesMateriaPrima 
   BackColor       =   &H00008000&
   Caption         =   "Reportes De MATERIA PRIMA"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ReportesMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin Crystal.CrystalReport CrReportes 
      Left            =   10440
      Top             =   4080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      Connect         =   "Pwd=metal"
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowAllowDrillDown=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
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
      Height          =   6615
      Left            =   120
      TabIndex        =   83
      Top             =   120
      Visible         =   0   'False
      Width           =   11775
      Begin VB.Data DataBusqueda 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   324
         Left            =   2400
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2520
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
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
         Height          =   735
         Left            =   6360
         TabIndex        =   88
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   90
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   89
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   85
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   84
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   86
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   10680
         Picture         =   "ReportesMateriaPrima.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   99
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin MSDBGrid.DBGrid DBGridBusqueda 
         Bindings        =   "ReportesMateriaPrima.frx":24B4
         Height          =   5175
         Left            =   120
         OleObjectBlob   =   "ReportesMateriaPrima.frx":24CF
         TabIndex        =   87
         Top             =   1080
         Width           =   11415
      End
      Begin VB.Label LblBusqueda 
         Alignment       =   1  'Right Justify
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
         Left            =   120
         TabIndex        =   100
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   10200
      Picture         =   "ReportesMateriaPrima.frx":2EA9
      Style           =   1  'Graphical
      TabIndex        =   92
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   10200
      Picture         =   "ReportesMateriaPrima.frx":4F1B
      Style           =   1  'Graphical
      TabIndex        =   91
      Top             =   240
      Width           =   1575
   End
   Begin TabDlg.SSTab TabReportes 
      Height          =   6372
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9972
      _ExtentX        =   17595
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   8
      TabsPerRow      =   4
      TabHeight       =   1058
      TabCaption(0)   =   "Inventario"
      TabPicture(0)   =   "ReportesMateriaPrima.frx":6C15
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblInvOpc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblInv"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblInv2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblInvEtiTipMatPri"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblInvDesTipMatPri"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblInvDesCod"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TxtInvOpc"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TxtInv"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "FrameInvResDet"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrameInvOpc"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "FrameInvTipRep"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameInvTipBus"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "OptInv(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "OptInv(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "TxtInvOpc2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "OptInv(2)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "FrameInvPro"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "Entradas"
      TabPicture(1)   =   "ReportesMateriaPrima.frx":74EF
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "OptEntradas(6)"
      Tab(1).Control(1)=   "OptEntradas(5)"
      Tab(1).Control(2)=   "FrameEntradas"
      Tab(1).Control(3)=   "OptEntradas(4)"
      Tab(1).Control(4)=   "OptEntradas(3)"
      Tab(1).Control(5)=   "DtpEntFecFin"
      Tab(1).Control(6)=   "DtpEntFecIni"
      Tab(1).Control(7)=   "TxtEntradas"
      Tab(1).Control(8)=   "OptEntradas(2)"
      Tab(1).Control(9)=   "OptEntradas(1)"
      Tab(1).Control(10)=   "OptEntradas(0)"
      Tab(1).Control(11)=   "LblEntFecFin"
      Tab(1).Control(12)=   "LblEntFecIni"
      Tab(1).Control(13)=   "LblEntEti"
      Tab(1).Control(14)=   "LblEntDes"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Traslados"
      TabPicture(2)   =   "ReportesMateriaPrima.frx":7809
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameTraslados2"
      Tab(2).Control(1)=   "TxtTraTipDoc"
      Tab(2).Control(2)=   "FrameTraslados"
      Tab(2).Control(3)=   "TxtTraslados2"
      Tab(2).Control(4)=   "OptTraslados(5)"
      Tab(2).Control(5)=   "OptTraslados(4)"
      Tab(2).Control(6)=   "OptTraslados(9)"
      Tab(2).Control(7)=   "OptTraslados(8)"
      Tab(2).Control(8)=   "OptTraslados(0)"
      Tab(2).Control(9)=   "OptTraslados(1)"
      Tab(2).Control(10)=   "OptTraslados(2)"
      Tab(2).Control(11)=   "OptTraslados(3)"
      Tab(2).Control(12)=   "TxtTraslados"
      Tab(2).Control(13)=   "OptTraslados(6)"
      Tab(2).Control(14)=   "OptTraslados(7)"
      Tab(2).Control(15)=   "DtpTraFecFin"
      Tab(2).Control(16)=   "DtpTraFecIni"
      Tab(2).Control(17)=   "LblTraDesDoc"
      Tab(2).Control(18)=   "LblTraEtiDoc"
      Tab(2).Control(19)=   "LblTraslados2"
      Tab(2).Control(20)=   "LblTraBod2"
      Tab(2).Control(21)=   "LblTraBod"
      Tab(2).Control(22)=   "LblLabel(0)"
      Tab(2).Control(23)=   "LblLabel(1)"
      Tab(2).Control(24)=   "LblTraslados"
      Tab(2).ControlCount=   25
      TabCaption(3)   =   "Salidas"
      TabPicture(3)   =   "ReportesMateriaPrima.frx":9FBB
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "FrameDespachos"
      Tab(3).Control(1)=   "OptSalidas(4)"
      Tab(3).Control(2)=   "DtpSalFecFin"
      Tab(3).Control(3)=   "DtpSalFecIni"
      Tab(3).Control(4)=   "TxtSalidas"
      Tab(3).Control(5)=   "OptSalidas(3)"
      Tab(3).Control(6)=   "OptSalidas(2)"
      Tab(3).Control(7)=   "OptSalidas(1)"
      Tab(3).Control(8)=   "OptSalidas(0)"
      Tab(3).Control(9)=   "LblSalFecFin"
      Tab(3).Control(10)=   "LblSalFecIni"
      Tab(3).Control(11)=   "LblSalDes"
      Tab(3).Control(12)=   "LblSalEti"
      Tab(3).ControlCount=   13
      TabCaption(4)   =   "Cierre De Bulto"
      TabPicture(4)   =   "ReportesMateriaPrima.frx":D96D
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "LblCerBulEti"
      Tab(4).Control(1)=   "LblCerBulDes"
      Tab(4).Control(2)=   "LblCerBulFecIni"
      Tab(4).Control(3)=   "LblCerBulFecFin"
      Tab(4).Control(4)=   "OptCerrarBulto(0)"
      Tab(4).Control(5)=   "OptCerrarBulto(1)"
      Tab(4).Control(6)=   "OptCerrarBulto(2)"
      Tab(4).Control(7)=   "OptCerrarBulto(3)"
      Tab(4).Control(8)=   "OptCerrarBulto(4)"
      Tab(4).Control(9)=   "OptCerrarBulto(5)"
      Tab(4).Control(10)=   "TxtCerrarBulto"
      Tab(4).Control(11)=   "DtpCerBulFecIni"
      Tab(4).Control(12)=   "DtpCerBulFecFin"
      Tab(4).Control(13)=   "OptCerrarBulto(6)"
      Tab(4).Control(14)=   "FrameCierreBulto"
      Tab(4).ControlCount=   15
      TabCaption(5)   =   "Catalogo Articulos"
      TabPicture(5)   =   "ReportesMateriaPrima.frx":1011F
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "TxtCorrelativos"
      Tab(5).Control(1)=   "OptCorrelativos(1)"
      Tab(5).Control(2)=   "OptCorrelativos(0)"
      Tab(5).Control(3)=   "LblDesCor"
      Tab(5).Control(4)=   "LblEtiCor"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Desperdicio"
      TabPicture(6)   =   "ReportesMateriaPrima.frx":15911
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "OptDesperdicio(3)"
      Tab(6).Control(0).Enabled=   0   'False
      Tab(6).Control(1)=   "FrameDesperdicio"
      Tab(6).Control(1).Enabled=   0   'False
      Tab(6).Control(2)=   "Frame1"
      Tab(6).Control(2).Enabled=   0   'False
      Tab(6).Control(3)=   "TxtDesperdicio"
      Tab(6).Control(3).Enabled=   0   'False
      Tab(6).Control(4)=   "DTPDesFecFin"
      Tab(6).Control(4).Enabled=   0   'False
      Tab(6).Control(5)=   "DTPDesFecIni"
      Tab(6).Control(5).Enabled=   0   'False
      Tab(6).Control(6)=   "OptDesperdicio(2)"
      Tab(6).Control(6).Enabled=   0   'False
      Tab(6).Control(7)=   "OptDesperdicio(1)"
      Tab(6).Control(7).Enabled=   0   'False
      Tab(6).Control(8)=   "OptDesperdicio(0)"
      Tab(6).Control(8).Enabled=   0   'False
      Tab(6).Control(9)=   "LblDesDes"
      Tab(6).Control(9).Enabled=   0   'False
      Tab(6).Control(10)=   "LblDesEti"
      Tab(6).Control(10).Enabled=   0   'False
      Tab(6).Control(11)=   "LblDesFecFin"
      Tab(6).Control(11).Enabled=   0   'False
      Tab(6).Control(12)=   "LblDesFecIni"
      Tab(6).Control(12).Enabled=   0   'False
      Tab(6).ControlCount=   13
      TabCaption(7)   =   "Bodegas"
      TabPicture(7)   =   "ReportesMateriaPrima.frx":1BBAB
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "TxtBodegas"
      Tab(7).Control(1)=   "OptBodegas(2)"
      Tab(7).Control(2)=   "OptBodegas(1)"
      Tab(7).Control(3)=   "OptBodegas(0)"
      Tab(7).Control(4)=   "LblBodegas"
      Tab(7).ControlCount=   5
      Begin VB.OptionButton OptDesperdicio 
         Caption         =   "Fechas y Grupo"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   165
         Top             =   1920
         Width           =   1695
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Fechas Y Tipo De Materia Prima"
         Height          =   195
         Index           =   6
         Left            =   -74400
         TabIndex        =   162
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Frame FrameDesperdicio 
         Caption         =   "Tipo De Reporte"
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
         Left            =   -70680
         TabIndex        =   158
         Top             =   1560
         Width           =   3015
         Begin VB.OptionButton OptDetalleFecha 
            Caption         =   "Detalle x Fecha"
            Height          =   195
            Left            =   240
            TabIndex        =   166
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton OptDesCuaPro 
            Caption         =   "Cuadricula"
            Height          =   195
            Left            =   240
            TabIndex        =   163
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton OptResumen 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   240
            TabIndex        =   160
            Top             =   600
            Width           =   1095
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detalle x Catalogo De Productos"
            Height          =   195
            Left            =   240
            TabIndex        =   159
            Top             =   360
            Value           =   -1  'True
            Width           =   2655
         End
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Orden "
         Height          =   195
         Index           =   5
         Left            =   -74400
         TabIndex        =   153
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Frame FrameInvPro 
         Caption         =   "Tipo De Bodega"
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
         Height          =   732
         Left            =   3840
         TabIndex        =   150
         Top             =   3960
         Width           =   5652
         Begin VB.OptionButton OptInvPro 
            Caption         =   "Demas"
            Height          =   192
            Index           =   3
            Left            =   3720
            TabIndex        =   157
            Top             =   360
            Width           =   972
         End
         Begin VB.OptionButton OptInvPro 
            Caption         =   "No Conforme"
            Height          =   192
            Index           =   2
            Left            =   2280
            TabIndex        =   156
            Top             =   360
            Width           =   1452
         End
         Begin VB.OptionButton OptInvPro 
            Caption         =   "Proceso"
            Height          =   192
            Index           =   1
            Left            =   1200
            TabIndex        =   152
            Top             =   360
            Width           =   1092
         End
         Begin VB.OptionButton OptInvPro 
            Caption         =   "Todas"
            Height          =   192
            Index           =   0
            Left            =   240
            TabIndex        =   151
            Top             =   360
            Value           =   -1  'True
            Width           =   972
         End
      End
      Begin VB.TextBox TxtBodegas 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -70080
         TabIndex        =   147
         Top             =   3720
         Width           =   1452
      End
      Begin VB.OptionButton OptBodegas 
         Caption         =   "Grupo"
         Height          =   192
         Index           =   2
         Left            =   -74160
         TabIndex        =   146
         Top             =   2520
         Width           =   1332
      End
      Begin VB.OptionButton OptBodegas 
         Caption         =   "Descripcion"
         Height          =   192
         Index           =   1
         Left            =   -74160
         TabIndex        =   145
         Top             =   2160
         Width           =   1332
      End
      Begin VB.OptionButton OptBodegas 
         Caption         =   "Codigo"
         Height          =   192
         Index           =   0
         Left            =   -74160
         TabIndex        =   144
         Top             =   1800
         Width           =   1332
      End
      Begin VB.OptionButton OptInv 
         Caption         =   "Orden"
         Height          =   192
         Index           =   2
         Left            =   2280
         TabIndex        =   139
         Top             =   4320
         Width           =   732
      End
      Begin VB.Frame FrameTraslados2 
         Caption         =   "Tipo De Reporte"
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
         Height          =   855
         Left            =   -68760
         TabIndex        =   44
         Top             =   1440
         Width           =   3255
         Begin VB.OptionButton OptTraRes 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   1320
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OptTraDet 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox TxtTraTipDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71520
         TabIndex        =   54
         Top             =   5880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame FrameTraslados 
         Caption         =   "Tipo De Documento"
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
         Height          =   855
         Left            =   -68760
         TabIndex        =   47
         Top             =   2400
         Width           =   3255
         Begin VB.OptionButton OptTraOpc 
            Caption         =   "Un Tipo Documento"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   49
            Top             =   360
            Width           =   1815
         End
         Begin VB.OptionButton OptTraOpc 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame FrameDespachos 
         Caption         =   "Tipo De Reporte"
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
         Height          =   735
         Left            =   -69720
         TabIndex        =   60
         Top             =   1560
         Width           =   4215
         Begin VB.OptionButton OptDesResCli 
            Caption         =   "Resumen Cliente"
            Height          =   195
            Left            =   2520
            TabIndex        =   63
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptDesRes 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   1320
            TabIndex        =   62
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton OptDesDet 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   240
            TabIndex        =   61
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Opciones De Reporte"
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
         Height          =   1335
         Left            =   -67560
         TabIndex        =   131
         Top             =   1560
         Width           =   2175
         Begin VB.OptionButton OptDesPac 
            Caption         =   "Pacas"
            Height          =   195
            Left            =   120
            TabIndex        =   133
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton OptDesPro 
            Caption         =   "Procesos"
            Height          =   195
            Left            =   120
            TabIndex        =   132
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.Frame FrameEntradas 
         Caption         =   "Tipo De Reporte"
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
         Left            =   -67680
         TabIndex        =   26
         Top             =   1440
         Width           =   2055
         Begin VB.OptionButton OptEntResCua 
            Caption         =   "Resumen Cuadricula"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1440
            Width           =   1815
         End
         Begin VB.OptionButton OptEntResPro 
            Caption         =   "Resumen Proveedor"
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton OptEntResumen 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton OptEntDetalle 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.Frame FrameCierreBulto 
         Caption         =   "Tipo De Reporte"
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
         Height          =   735
         Left            =   -68280
         TabIndex        =   74
         Top             =   1560
         Width           =   2535
         Begin VB.OptionButton OptCerrarBulto 
            Caption         =   "Resumen"
            Height          =   195
            Index           =   8
            Left            =   1320
            TabIndex        =   76
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton OptCerrarBulto 
            Caption         =   "Detalle"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   75
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.OptionButton OptCerrarBulto 
         Caption         =   "Fechas Y Tipo Materia Prima"
         Height          =   195
         Index           =   6
         Left            =   -74400
         TabIndex        =   73
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox TxtTraslados2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71520
         TabIndex        =   53
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Bodega Entrada Y Tipo Materia Prima"
         Height          =   195
         Index           =   5
         Left            =   -74400
         TabIndex        =   39
         Top             =   3240
         Width           =   3855
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Bodega Salida Y Tipo Materia Prima"
         Height          =   195
         Index           =   4
         Left            =   -74400
         TabIndex        =   38
         Top             =   2880
         Width           =   3615
      End
      Begin VB.TextBox TxtDesperdicio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71640
         TabIndex        =   126
         Top             =   4920
         Visible         =   0   'False
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DTPDesFecFin 
         Height          =   255
         Left            =   -66840
         TabIndex        =   123
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37399
      End
      Begin MSComCtl2.DTPicker DTPDesFecIni 
         Height          =   255
         Left            =   -69360
         TabIndex        =   122
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37399
      End
      Begin VB.OptionButton OptDesperdicio 
         Caption         =   "Fechas y Ficha Tecnica"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   121
         Top             =   2640
         Width           =   2175
      End
      Begin VB.OptionButton OptDesperdicio 
         Caption         =   "Fechas y Proceso"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   120
         Top             =   2280
         Width           =   1695
      End
      Begin VB.OptionButton OptDesperdicio 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   119
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Salidas No Liberados"
         Height          =   195
         Index           =   4
         Left            =   -74400
         TabIndex        =   59
         Top             =   2880
         Width           =   2175
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Entradas No Liberadas"
         Height          =   195
         Index           =   4
         Left            =   -74400
         TabIndex        =   25
         Top             =   3360
         Width           =   2055
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Traslados No Liberados"
         Height          =   195
         Index           =   9
         Left            =   -74400
         TabIndex        =   43
         Top             =   4680
         Width           =   2655
      End
      Begin MSComCtl2.DTPicker DtpSalFecFin 
         Height          =   255
         Left            =   -69000
         TabIndex        =   65
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37389
      End
      Begin MSComCtl2.DTPicker DtpSalFecIni 
         Height          =   255
         Left            =   -71040
         TabIndex        =   64
         Top             =   3600
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37389
      End
      Begin VB.TextBox TxtSalidas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         TabIndex        =   66
         Top             =   4560
         Width           =   1935
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Materia Prima"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   58
         Top             =   2520
         Width           =   1815
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas Y Materia Prima"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   57
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas Y Cliente"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   56
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   55
         Top             =   1440
         Width           =   1815
      End
      Begin VB.TextBox TxtCorrelativos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71520
         TabIndex        =   82
         Top             =   4560
         Width           =   1455
      End
      Begin VB.OptionButton OptCorrelativos 
         Caption         =   "Tipo De Articulo"
         Height          =   195
         Index           =   1
         Left            =   -74280
         TabIndex        =   81
         Top             =   2040
         Width           =   2055
      End
      Begin VB.OptionButton OptCorrelativos 
         Caption         =   "Codigo Articulo"
         Height          =   195
         Index           =   0
         Left            =   -74280
         TabIndex        =   80
         Top             =   1680
         Width           =   2055
      End
      Begin VB.TextBox TxtInvOpc2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   18
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DtpCerBulFecFin 
         Height          =   255
         Left            =   -66960
         TabIndex        =   78
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37337
      End
      Begin MSComCtl2.DTPicker DtpCerBulFecIni 
         Height          =   255
         Left            =   -69000
         TabIndex        =   77
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37337
      End
      Begin VB.TextBox TxtCerrarBulto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         TabIndex        =   79
         Top             =   5040
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptCerrarBulto 
         Caption         =   "Numero Ingreso"
         Height          =   195
         Index           =   5
         Left            =   -74400
         TabIndex        =   72
         Top             =   3240
         Width           =   1695
      End
      Begin VB.OptionButton OptCerrarBulto 
         Caption         =   "Materia Prima"
         Height          =   195
         Index           =   4
         Left            =   -74400
         TabIndex        =   71
         Top             =   2880
         Width           =   1695
      End
      Begin VB.OptionButton OptCerrarBulto 
         Caption         =   "Fechas Codigo Materia Prima"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   70
         Top             =   2520
         Width           =   2535
      End
      Begin VB.OptionButton OptCerrarBulto 
         Caption         =   "Fechas Y Linea"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   69
         Top             =   2160
         Width           =   1695
      End
      Begin VB.OptionButton OptCerrarBulto 
         Caption         =   "Fechas Y Turno"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   68
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton OptCerrarBulto 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   67
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Materia Prima"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   24
         Top             =   3000
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DtpEntFecFin 
         Height          =   255
         Left            =   -68280
         TabIndex        =   32
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37336
      End
      Begin MSComCtl2.DTPicker DtpEntFecIni 
         Height          =   255
         Left            =   -70440
         TabIndex        =   31
         Top             =   3720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37336
      End
      Begin VB.TextBox TxtEntradas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72360
         TabIndex        =   33
         Top             =   4680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Fechas Y Materia Prima"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   23
         Top             =   2280
         Width           =   2655
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Fechas Y Proveedor"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   22
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   21
         Top             =   1560
         Width           =   1815
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Materia Prima"
         Height          =   195
         Index           =   8
         Left            =   -74400
         TabIndex        =   42
         Top             =   3960
         Width           =   2655
      End
      Begin VB.OptionButton OptInv 
         Caption         =   "Codigo"
         Height          =   192
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   4320
         Value           =   -1  'True
         Width           =   852
      End
      Begin VB.OptionButton OptInv 
         Caption         =   "Descripcion"
         Height          =   192
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   4320
         Width           =   1212
      End
      Begin VB.Frame FrameInvTipBus 
         Caption         =   "Tipo De Busqueda"
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
         Height          =   735
         Left            =   3840
         TabIndex        =   6
         Top             =   1320
         Width           =   5655
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Index           =   1
            Left            =   2040
            TabIndex        =   8
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Igual a"
            Height          =   195
            Index           =   2
            Left            =   3720
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame FrameInvTipRep 
         Caption         =   "Tipo De Existencia"
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
         Height          =   612
         Left            =   3840
         TabIndex        =   16
         Top             =   3240
         Width           =   5655
         Begin VB.OptionButton OptTipRep 
            Caption         =   "Todos"
            Height          =   195
            Index           =   3
            Left            =   4080
            TabIndex        =   138
            Top             =   360
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.OptionButton OptTipRep 
            Caption         =   "= 0"
            Height          =   195
            Index           =   2
            Left            =   3240
            TabIndex        =   137
            Top             =   360
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton OptTipRep 
            Caption         =   "<= 0"
            Height          =   195
            Index           =   1
            Left            =   1320
            TabIndex        =   136
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton OptTipRep 
            Caption         =   "> 0"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.Frame FrameInvOpc 
         Caption         =   "Opciones De Inventario"
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
         Height          =   972
         Left            =   3840
         TabIndex        =   10
         Top             =   2160
         Width           =   5655
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Bodega Y Pasillo"
            Height          =   195
            Index           =   5
            Left            =   3840
            TabIndex        =   15
            Top             =   600
            Width           =   1575
         End
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Grupo Bodega"
            Height          =   195
            Index           =   4
            Left            =   3840
            TabIndex        =   149
            Top             =   360
            Width           =   1452
         End
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Bodega Y Tipo M.P."
            Height          =   195
            Index           =   3
            Left            =   1680
            TabIndex        =   14
            Top             =   600
            Width           =   2052
         End
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1212
         End
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Bodega"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   12
            Top             =   600
            Width           =   1212
         End
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Tipo Materia Prima"
            Height          =   195
            Index           =   2
            Left            =   1680
            TabIndex        =   13
            Top             =   360
            Width           =   1812
         End
      End
      Begin VB.Frame FrameInvResDet 
         Caption         =   "Opciones De Reporte"
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
         Height          =   2895
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   3375
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen Cuadricula x Orden"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   164
            Top             =   1440
            Width           =   2655
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen Cuadricula x Codigo"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   161
            Top             =   1200
            Width           =   2655
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Detallado x Bodega y Tipo"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   155
            Top             =   2520
            Width           =   2292
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen x Bodega y Tipo"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   154
            Top             =   960
            Width           =   2415
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Detallado x Tipo"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   143
            Top             =   2280
            Width           =   1572
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen x Tipo"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   142
            Top             =   720
            Width           =   1572
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Detallado x Ficha y Bodega"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   141
            Top             =   2040
            Width           =   2292
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen x Ficha y Bodega "
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   140
            Top             =   480
            Width           =   2292
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen x Bodega y Orden"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   2412
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Detallado x Bodega y Orden"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   5
            Top             =   1800
            Width           =   2292
         End
      End
      Begin VB.TextBox TxtInv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   20
         Top             =   5760
         Width           =   1695
      End
      Begin VB.TextBox TxtInvOpc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4680
         TabIndex        =   19
         ToolTipText     =   "Signo '+' o Doble Click para Ayuda"
         Top             =   5400
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   34
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Numero Documento"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   35
         Top             =   1800
         Width           =   2295
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Bodega Salida y Codigo Materia Prima"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   36
         Top             =   2160
         Width           =   3855
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Bodega Entrada y Codigo Materia Prima"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   37
         Top             =   2520
         Width           =   4095
      End
      Begin VB.TextBox TxtTraslados 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71520
         TabIndex        =   52
         ToolTipText     =   "Signo '+' o Doble Click para Ayuda"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Codigo Materia Prima"
         Height          =   195
         Index           =   6
         Left            =   -74400
         TabIndex        =   40
         Top             =   3600
         Width           =   2655
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Orden"
         Height          =   195
         Index           =   7
         Left            =   -74400
         TabIndex        =   41
         Top             =   4320
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DtpTraFecFin 
         Height          =   255
         Left            =   -66960
         TabIndex        =   51
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37330
      End
      Begin MSComCtl2.DTPicker DtpTraFecIni 
         Height          =   255
         Left            =   -69120
         TabIndex        =   50
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62062595
         CurrentDate     =   37330
      End
      Begin VB.Label LblBodegas 
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
         Height          =   252
         Left            =   -71760
         TabIndex        =   148
         Top             =   3720
         Width           =   1572
      End
      Begin VB.Label LblTraDesDoc 
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
         Left            =   -69840
         TabIndex        =   135
         Top             =   5880
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label LblTraEtiDoc 
         AutoSize        =   -1  'True
         Caption         =   "Tipo De Documento"
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
         Left            =   -73320
         TabIndex        =   134
         Top             =   5880
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label LblTraslados2 
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   130
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label LblTraBod2 
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
         Left            =   -69840
         TabIndex        =   129
         Top             =   5520
         Width           =   4215
      End
      Begin VB.Label LblDesDes 
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
         Left            =   -69600
         TabIndex        =   128
         Top             =   4920
         Width           =   4215
      End
      Begin VB.Label LblDesEti 
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
         Height          =   255
         Left            =   -74040
         TabIndex        =   127
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Label LblDesFecFin 
         Alignment       =   1  'Right Justify
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
         Height          =   195
         Left            =   -67440
         TabIndex        =   125
         Top             =   3600
         Width           =   510
      End
      Begin VB.Label LblDesFecIni 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -70200
         TabIndex        =   124
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label LblSalFecFin 
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
         Height          =   195
         Left            =   -69720
         TabIndex        =   118
         Top             =   3600
         Width           =   510
      End
      Begin VB.Label LblSalFecIni 
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
         Height          =   195
         Left            =   -71760
         TabIndex        =   117
         Top             =   3600
         Width           =   555
      End
      Begin VB.Label LblSalDes 
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
         Left            =   -69720
         TabIndex        =   116
         Top             =   4560
         Width           =   4335
      End
      Begin VB.Label LblSalEti 
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
         Left            =   -74520
         TabIndex        =   115
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label LblDesCor 
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
         Left            =   -69960
         TabIndex        =   114
         Top             =   4560
         Width           =   4575
      End
      Begin VB.Label LblEtiCor 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo Articulo"
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
         Left            =   -74040
         TabIndex        =   113
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label LblInvDesCod 
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
         TabIndex        =   112
         Top             =   5760
         Width           =   3015
      End
      Begin VB.Label LblInvDesTipMatPri 
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
         TabIndex        =   111
         Top             =   5040
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label LblInvEtiTipMatPri 
         Alignment       =   1  'Right Justify
         Caption         =   "Tipo De Materia Prima"
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
         Left            =   1560
         TabIndex        =   110
         Top             =   5040
         Visible         =   0   'False
         Width           =   3015
      End
      Begin VB.Label LblCerBulFecFin 
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
         Height          =   195
         Left            =   -67560
         TabIndex        =   109
         Top             =   3960
         Width           =   510
      End
      Begin VB.Label LblCerBulFecIni 
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
         Height          =   195
         Left            =   -69600
         TabIndex        =   108
         Top             =   3960
         Width           =   555
      End
      Begin VB.Label LblCerBulDes 
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
         Left            =   -70080
         TabIndex        =   107
         Top             =   5040
         Width           =   4335
      End
      Begin VB.Label LblCerBulEti 
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
         Height          =   255
         Left            =   -74280
         TabIndex        =   106
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label LblEntFecFin 
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
         Height          =   195
         Left            =   -69000
         TabIndex        =   105
         Top             =   3720
         Width           =   510
      End
      Begin VB.Label LblEntFecIni 
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
         Height          =   195
         Left            =   -71040
         TabIndex        =   104
         Top             =   3720
         Width           =   555
      End
      Begin VB.Label LblEntEti 
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
         Height          =   255
         Left            =   -74640
         TabIndex        =   103
         Top             =   4680
         Width           =   2175
      End
      Begin VB.Label LblEntDes 
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
         Left            =   -70440
         TabIndex        =   102
         Top             =   4680
         Width           =   4935
      End
      Begin VB.Label LblInv2 
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
         TabIndex        =   101
         Top             =   5400
         Width           =   3015
      End
      Begin VB.Label LblInv 
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
         Left            =   1560
         TabIndex        =   98
         Top             =   5760
         Width           =   3015
      End
      Begin VB.Label LblInvOpc 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   97
         Top             =   5400
         Width           =   3015
      End
      Begin VB.Label LblTraBod 
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
         Left            =   -69840
         TabIndex        =   96
         Top             =   5160
         Width           =   4215
      End
      Begin VB.Label LblLabel 
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
         Height          =   195
         Index           =   0
         Left            =   -69840
         TabIndex        =   95
         Top             =   4800
         Width           =   555
      End
      Begin VB.Label LblLabel 
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
         Height          =   195
         Index           =   1
         Left            =   -67560
         TabIndex        =   94
         Top             =   4800
         Width           =   510
      End
      Begin VB.Label LblTraslados 
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
         Height          =   255
         Left            =   -74760
         TabIndex        =   93
         Top             =   5160
         Width           =   3135
      End
   End
End
Attribute VB_Name = "ReportesMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaBodega As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaTipoMateriaPrima As Recordset
Dim RBuscaProveedor As Recordset
Dim RBuscaCliente As Recordset
Dim RBuscaLinea As Recordset
Dim RBuscaProceso As Recordset
Dim RBuscaFichaTecnica As Recordset
Dim RBuscaDocumento As Recordset
Dim RBuscaBodegaGrupo As Recordset
Dim RBuscaTipo As Recordset

'VARIABLES PARA CARPETA DE INVENTARIO
Dim BInvMateriaPrima As Boolean
Dim BInvBodega As Boolean
Dim BInvTipoMateriaPrima As Boolean
Dim BInvTipoMateriaPrima2 As Boolean
Dim BInvBodegaGrupo As Boolean

'VARIABLES PARA CARPETA DE TRASLADOS
Dim BTraBodega As Boolean
Dim BTraMateriaPrima As Boolean
Dim BTraTipoMateriaPrima As Boolean
Dim BTraDocumentos As Boolean

'VARIABLES PARA CARPETA DE ENTRADAS
Dim BEntProveedor As Boolean
Dim BEntMateriaPrima As Boolean
Dim BEntTipoMateriaPrima As Boolean

'VARIABLES PARA CARPETA DE SALIDAS
Dim BSalCliente As Boolean
Dim BSalMateriaPrima As Boolean

'VARIABLES PARA CARPETA DE CERRAR BULTO
Dim BCerBulLinea As Boolean
Dim BCerBulMateriaPrima As Boolean
Dim BCerBulTipoDeMateriaPrima As Boolean

'VARIABLES PARA CARPETA DE DESPERDICIO
Dim BDesFichaTecnica As Boolean
Dim BDesProceso As Boolean



Dim VTexto As String

Dim VDia As String
Dim VMes As String
Dim VAño As String
Dim VDia2 As String
Dim VMes2 As String
Dim VAño2 As String


Private Sub CmdImprimir_Click()
On Error Resume Next
    MousePointer = 11
    
    'INVENTARIO MATERIA PRIMA
    If TabReportes.Tab = 0 Then
            'VA AL PROCEDIMIENTO DE INVENTARIO
            Inventario
    'ENTRADAS DE MATERIA PRIMA
    ElseIf TabReportes.Tab = 1 Then
            'VA AL PROCEDIMIENTO DE ENTRADAS
            Entradas
    'TRASLADOS DE MATERIA PRIMA
    ElseIf TabReportes.Tab = 2 Then
            'VA AL PROCEDIMIENTO DE TRASLADOS
            Traslados
    'SALIDAS DE MATERIA PRIMA
    ElseIf TabReportes.Tab = 3 Then
            'VA AL PROCEDIMIENTO DE SALIDAS
            Salidas
    'CERRAR BULTO
    ElseIf TabReportes.Tab = 4 Then
            CerrarBulto
    'CORRELATIVOS MAXIMOS
    ElseIf TabReportes.Tab = 5 Then
            CorrelativosMaximos
    'DESPERDICIO DE PROCESO
    ElseIf TabReportes.Tab = 6 Then
            Desperdicio
    'BODEGAS
    ElseIf TabReportes.Tab = 7 Then
            Bodegas
    
    End If
    
             'DESPLIEGA EL REPORTE
             CrReportes.Action = 1
             'CrReportes.PrintReport
             
             
             'CrReportes.DiscardSavedData = True
             MousePointer = 0
             If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Exit Sub
             End If
                
    
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
        'INVENTARIO
        If (BInvBodega = True Or BInvTipoMateriaPrima = True Or BInvBodegaGrupo = True) Then
                TxtInvOpc.Text = DBGridBusqueda.Columns(0).Text
                TxtInvOpc.SetFocus
        'TIPO DE MATERIA PRIMA 2
        ElseIf BInvTipoMateriaPrima2 = True Then
                TxtInvOpc2.Text = DBGridBusqueda.Columns(0).Text
                TxtInvOpc2.SetFocus
        'TIPO DE MATERIA PRIMA TRASLADOS
        ElseIf BTraTipoMateriaPrima = True Then
                TxtTraslados2.Text = DBGridBusqueda.Columns(0).Text
                TxtTraslados2.SetFocus
        'MATERIA PRIMA
        ElseIf BInvMateriaPrima = True Then
                TxtInv.Text = DBGridBusqueda.Columns(0).Text
                TxtInv.SetFocus
        'MATERIA PRIMA EN CARPETA DE ENTRADAS
        ElseIf BEntMateriaPrima = True Then
                TxtEntradas.Text = DBGridBusqueda.Columns(0).Text
                TxtEntradas.SetFocus
        'MATERIA PRIMA EN CARPETA DE DESPACHOS
        ElseIf BSalMateriaPrima = True Then
                TxtSalidas.Text = DBGridBusqueda.Columns(0).Text
                TxtSalidas.SetFocus
        'TRASLADOS
        ElseIf (BTraBodega = True Or BTraMateriaPrima = True) Then
                TxtTraslados.Text = DBGridBusqueda.Columns(0).Text
                TxtTraslados.SetFocus
        'TRASLADOS DOCUMENTO
        ElseIf BTraDocumentos = True Then
                TxtTraTipDoc.Text = DBGridBusqueda.Columns(0).Text
                TxtTraTipDoc.SetFocus
        'PROVEEDOR EN CARPETA DE ENTRADAS
        ElseIf BEntProveedor = True Then
                TxtEntradas.Text = DBGridBusqueda.Columns(1).Text
                TxtEntradas.SetFocus
        'CLIENTE EN CARPETA DE SALIDAS O DESPACHOS
        ElseIf BSalCliente = True Then
                TxtSalidas.Text = DBGridBusqueda.Columns(0).Text
                TxtSalidas.SetFocus
        'CERRAR BULTO
        ElseIf (BCerBulLinea = True Or BCerBulMateriaPrima = True Or BCerBulTipoDeMateriaPrima = True) Then
                TxtCerrarBulto.Text = DBGridBusqueda.Columns(0)
                TxtCerrarBulto.SetFocus
        'DESPERDICIO (PROCESOS)
        ElseIf BDesProceso = True Then
                TxtDesperdicio.Text = DBGridBusqueda.Columns(0)
                TxtDesperdicio.SetFocus
        'DESPERDICIO (FICHA TECNICA)
        ElseIf BDesFichaTecnica = True Then
                TxtDesperdicio.Text = DBGridBusqueda.Columns(0)
                TxtDesperdicio.SetFocus
        'TIPO MATERIA PRIMA EN CARPETA DE ENTRADAS
        ElseIf BEntTipoMateriaPrima = True Then
                TxtEntradas.Text = DBGridBusqueda.Columns(0).Text
                TxtEntradas.SetFocus
        
        End If
                FrameBusqueda.Visible = False
                
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        'SI PRECIONA LA TECLA DEL SIGNO '+'
        If KeyAscii = 43 Then
                'INVENTARIO
                If (BInvBodega = True Or BInvTipoMateriaPrima = True Or BInvBodegaGrupo = True) Then
                        TxtInvOpc.Text = DBGridBusqueda.Columns(0).Text
                        TxtInvOpc.SetFocus
                'TIPO DE MATERIA PRIMA 2
                ElseIf BInvTipoMateriaPrima2 = True Then
                        TxtInvOpc2.Text = DBGridBusqueda.Columns(0).Text
                        TxtInvOpc2.SetFocus
                'TIPO DE MATERIA PRIMA TRASLADOS
                ElseIf BTraTipoMateriaPrima = True Then
                        TxtTraslados2.Text = DBGridBusqueda.Columns(0).Text
                        TxtTraslados2.SetFocus
                'MATERIA PRIMA
                ElseIf BInvMateriaPrima = True Then
                        TxtInv.Text = DBGridBusqueda.Columns(0).Text
                        TxtInv.SetFocus
                'MATERIA PRIMA EN CARPETA DE ENTRADAS
                ElseIf BEntMateriaPrima = True Then
                        TxtEntradas.Text = DBGridBusqueda.Columns(0).Text
                        TxtEntradas.SetFocus
                'MATERIA PRIMA EN CARPETA DE DESPACHOS
                ElseIf BSalMateriaPrima = True Then
                        TxtSalidas.Text = DBGridBusqueda.Columns(0).Text
                        TxtSalidas.SetFocus
                'TRASLADOS
                ElseIf (BTraBodega = True Or BTraMateriaPrima = True) Then
                        TxtTraslados.Text = DBGridBusqueda.Columns(0).Text
                        TxtTraslados.SetFocus
                'TRASLADOS DOCUMENTO
                ElseIf BTraDocumentos = True Then
                        TxtTraTipDoc.Text = DBGridBusqueda.Columns(0).Text
                        TxtTraTipDoc.SetFocus
                'PROVEEDOR EN CARPETA DE ENTRADAS
                ElseIf BEntProveedor = True Then
                        TxtEntradas.Text = DBGridBusqueda.Columns(1).Text
                        TxtEntradas.SetFocus
                'CLIENTE EN CARPETA DE SALIDAS O DESPACHOS
                ElseIf BSalCliente = True Then
                        TxtSalidas.Text = DBGridBusqueda.Columns(0).Text
                        TxtSalidas.SetFocus
                'CERRAR BULTO
                ElseIf (BCerBulLinea = True Or BCerBulMateriaPrima = True Or BCerBulTipoDeMateriaPrima = True) Then
                        TxtCerrarBulto.Text = DBGridBusqueda.Columns(0)
                        TxtCerrarBulto.SetFocus
                'DESPERDICIO (PROCESOS)
                ElseIf BDesProceso = True Then
                        TxtDesperdicio.Text = DBGridBusqueda.Columns(0)
                        TxtDesperdicio.SetFocus
                'DESPERDICIO (FICHA TECNICA)
                ElseIf BDesFichaTecnica = True Then
                        TxtDesperdicio.Text = DBGridBusqueda.Columns(0)
                        TxtDesperdicio.SetFocus
                'TIPO MATERIA PRIMA EN CARPETA DE ENTRADAS
                ElseIf BEntTipoMateriaPrima = True Then
                        TxtEntradas.Text = DBGridBusqueda.Columns(0).Text
                        TxtEntradas.SetFocus
                
                End If
                        FrameBusqueda.Visible = False
        End If
End Sub


Private Sub Form_Load()

        
        'FECHAS DE TAB DE TRASLADOS
        DtpTraFecIni.Value = Date
        DtpTraFecFin.Value = Date
        
        'FECHAS DE TAB DE ENTRADAS
        DtpEntFecIni.Value = Date
        DtpEntFecFin.Value = Date
        
        'FECHAS DE TAB DE CERRAR BULTO
        DtpCerBulFecIni.Value = Date
        DtpCerBulFecFin.Value = Date
        
        'FECHAS DE TAB DE SALIDAS
        DtpSalFecIni.Value = Date
        DtpSalFecFin.Value = Date
        
        'FECHAS DE TAB DE DESPERDICIO
        DTPDesFecIni.Value = Date
        DTPDesFecFin.Value = Date
        
        DataBusqueda.ConnectionString = GTipoProveedor
        DataBusqueda.Refresh
End Sub

Private Sub OptBodegas_Click(Index As Integer)
        If Index = 0 Then
            LblBodegas.Caption = "Codigo"
        ElseIf Index = 1 Then
            LblBodegas.Caption = "Descripcion"
        ElseIf Index = 2 Then
            LblBodegas.Caption = "Grupo"
        End If
            TxtBodegas.SetFocus
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
        If OptBusqueda.Item(0).Value = True Then
                LblBusqueda.Caption = "Descripcion"
        ElseIf OptBusqueda.Item(1).Value = True Then
            LblBusqueda.Caption = "Codigo"
        End If
            Txtbusqueda.SetFocus
End Sub

Private Sub OptCerrarBulto_Click(Index As Integer)
        If Index = 0 Then
            LblCerBulEti.Caption = ""
        ElseIf Index = 1 Then
            LblCerBulEti.Caption = "Turno"
        ElseIf Index = 2 Then
            LblCerBulEti.Caption = "Linea"
        ElseIf Index = 3 Then
            LblCerBulEti.Caption = "Codigo Materia Prima"
        ElseIf Index = 4 Then
            LblCerBulEti.Caption = "Codigo Materia Prima"
        ElseIf Index = 5 Then
            LblCerBulEti.Caption = "Numero Ingreso"
        ElseIf Index = 6 Then
            LblCerBulEti.Caption = "Tipo De Materia Prima"
        End If
        
        If (Index = 0 Or Index = 1 Or Index = 2 Or Index = 3 Or Index = 6 Or Index = 7 Or Index = 8) Then
            LblCerBulFecIni.Visible = True
            LblCerBulFecFin.Visible = True
            DtpCerBulFecIni.Visible = True
            DtpCerBulFecFin.Visible = True
        Else
            LblCerBulFecIni.Visible = False
            LblCerBulFecFin.Visible = False
            DtpCerBulFecIni.Visible = False
            DtpCerBulFecFin.Visible = False
        End If
        
        If Index = 0 Then
            TxtCerrarBulto.Visible = False
        Else
            TxtCerrarBulto.Visible = True
            TxtCerrarBulto.SetFocus
        End If
        
End Sub

Private Sub OptCorrelativos_Click(Index As Integer)
        If OptCorrelativos.Item(0).Value = True Then
            LblEtiCor.Caption = "Codigo Articulo"
        ElseIf OptCorrelativos.Item(1).Value = True Then
            LblEtiCor.Caption = "Tipo Articulo"
        End If
            TxtCorrelativos.SetFocus
        
End Sub

Private Sub OptDesPac_Click()
        OptDesperdicio.Item(0).Visible = True
        OptDesperdicio.Item(1).Visible = False
        OptDesperdicio.Item(2).Visible = False
        OptDesperdicio.Item(0).Value = True
        FrameDesperdicio.Visible = False
End Sub

Private Sub OptDesperdicio_Click(Index As Integer)
        If Index = 0 Then
            LblDesEti.Caption = ""
            TxtDesperdicio.Visible = False
        ElseIf Index = 1 Then
            LblDesEti.Caption = "Codigo De Proceso"
            TxtDesperdicio.Visible = True
            TxtDesperdicio.SetFocus
        ElseIf Index = 2 Then
            LblDesEti.Caption = "Codigo Ficha Tecnica"
            TxtDesperdicio.Visible = True
            TxtDesperdicio.SetFocus
        ElseIf Index = 3 Then
            LblDesEti.Caption = "Codigo De Grupo"
            TxtDesperdicio.Visible = True
            TxtDesperdicio.SetFocus

        End If
End Sub

Private Sub OptDesPro_Click()
        OptDesperdicio.Item(0).Visible = True
        OptDesperdicio.Item(1).Visible = True
        OptDesperdicio.Item(2).Visible = True
        OptDesperdicio.Item(0).Value = True
        FrameDesperdicio.Visible = True
End Sub

Private Sub OptEntradas_Click(Index As Integer)
        If Index = 0 Then
                DtpEntFecIni.Visible = True
                DtpEntFecFin.Visible = True
                LblEntFecIni.Visible = True
                LblEntFecFin.Visible = True
                TxtEntradas.Visible = False
                LblEntEti.Caption = ""
        ElseIf Index = 1 Then
                DtpEntFecIni.Visible = True
                DtpEntFecFin.Visible = True
                LblEntFecIni.Visible = True
                LblEntFecFin.Visible = True
                TxtEntradas.Visible = False
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Codigo Proveedor"
        ElseIf Index = 2 Then
                DtpEntFecIni.Visible = True
                DtpEntFecFin.Visible = True
                LblEntFecIni.Visible = True
                LblEntFecFin.Visible = True
                TxtEntradas.Visible = False
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Codigo Materia Prima"
        ElseIf Index = 3 Then
                DtpEntFecIni.Visible = False
                DtpEntFecFin.Visible = False
                LblEntFecIni.Visible = False
                LblEntFecFin.Visible = False
                TxtEntradas.Visible = False
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Codigo Materia Prima"
        ElseIf Index = 4 Then
                DtpEntFecIni.Visible = False
                DtpEntFecFin.Visible = False
                LblEntFecIni.Visible = False
                LblEntFecFin.Visible = False
                TxtEntradas.Visible = False
                TxtEntradas.Visible = False
                LblEntEti.Caption = ""
        ElseIf Index = 5 Then
                DtpEntFecIni.Visible = False
                DtpEntFecFin.Visible = False
                LblEntFecIni.Visible = False
                LblEntFecFin.Visible = False
                TxtEntradas.Visible = False
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Orden Produccion"
        ElseIf Index = 6 Then
                DtpEntFecIni.Visible = True
                DtpEntFecFin.Visible = True
                LblEntFecIni.Visible = True
                LblEntFecFin.Visible = True
                TxtEntradas.Visible = False
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Tipo Materia Prima"
        End If
End Sub

Private Sub OptInv_Click(Index As Integer)
    If Index = 0 Then
        LblInv.Caption = "Codigo "
    ElseIf Index = 1 Then
        LblInv.Caption = "Descripcion"
    ElseIf Index = 2 Then
        LblInv.Caption = "Orden"
    End If
        TxtInv.SetFocus
End Sub

Private Sub OptInvOpc_Click(Index As Integer)
    If Index = 0 Then
        LblInvOpc.Caption = ""
        TxtInvOpc.Visible = False
        LblInvDesTipMatPri.Visible = False
        TxtInvOpc2.Visible = False
        LblInvEtiTipMatPri.Visible = False
    ElseIf Index = 1 Then
        LblInvOpc.Caption = "Codigo Bodega Disponible"
        TxtInvOpc.Visible = True
        TxtInvOpc.SetFocus
        LblInvDesTipMatPri.Visible = False
        TxtInvOpc2.Visible = False
        LblInvEtiTipMatPri.Visible = False
    ElseIf Index = 2 Then
        LblInvOpc.Caption = "Tipo Materia Prima"
        TxtInvOpc.Visible = True
        TxtInvOpc.SetFocus
        LblInvDesTipMatPri.Visible = False
        TxtInvOpc2.Visible = False
        LblInvEtiTipMatPri.Visible = False
    ElseIf Index = 3 Then
        LblInvOpc.Caption = "Codigo Bodega Disponible"
        TxtInvOpc.Visible = True
        LblInvDesTipMatPri.Visible = True
        TxtInvOpc2.Visible = True
        LblInvEtiTipMatPri.Visible = True
        LblInvEtiTipMatPri.Caption = "Tipo MateriaPrima"
        TxtInvOpc2.SetFocus
    ElseIf Index = 4 Then
        LblInvOpc.Caption = "Grupo Bodega"
        TxtInvOpc.Visible = True
        TxtInvOpc.SetFocus
        LblInvDesTipMatPri.Visible = False
        TxtInvOpc2.Visible = False
        LblInvEtiTipMatPri.Visible = False
    ElseIf Index = 5 Then
        LblInvOpc.Caption = "Codigo Bodega Disponible"
        TxtInvOpc.Visible = True
        LblInvDesTipMatPri.Visible = True
        TxtInvOpc2.Visible = True
        LblInvEtiTipMatPri.Visible = True
        LblInvEtiTipMatPri.Caption = "Pasillo"
        TxtInvOpc2.SetFocus

    
    End If
    
End Sub


Private Sub OptSalidas_Click(Index As Integer)
        If Index = 0 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = False
                LblSalEti.Caption = ""
        ElseIf Index = 1 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = False
                TxtSalidas.Visible = True
                TxtSalidas.SetFocus
                LblSalEti.Caption = "Codigo Cliente"
        ElseIf Index = 2 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = False
                TxtSalidas.Visible = True
                TxtSalidas.SetFocus
                LblSalEti.Caption = "Codigo Materia Prima"
        ElseIf Index = 3 Then
                DtpSalFecIni.Visible = False
                DtpSalFecFin.Visible = False
                LblSalFecIni.Visible = False
                LblSalFecFin.Visible = False
                TxtSalidas.Visible = False
                TxtSalidas.Visible = True
                TxtSalidas.SetFocus
                LblSalEti.Caption = "Codigo Materia Prima"
        ElseIf Index = 4 Then
                DtpSalFecIni.Visible = False
                DtpSalFecFin.Visible = False
                LblSalFecIni.Visible = False
                LblSalFecFin.Visible = False
                TxtSalidas.Visible = False
                TxtSalidas.Visible = False
                LblSalEti.Caption = ""
        End If

End Sub

Private Sub OptTraOpc_Click(Index As Integer)
        'SI LA OPCION ES DE TODOS
        If Index = 0 Then
            TxtTraTipDoc.Visible = False
            LblTraDesDoc.Visible = False
            LblTraEtiDoc.Visible = False
        'SI LA OPCION ES POR UN TIPO DE DOCUMENTO
        Else
            TxtTraTipDoc.Visible = True
            LblTraDesDoc.Visible = True
            LblTraEtiDoc.Visible = True
        End If
End Sub

Private Sub OptTraslados_Click(Index As Integer)
        'FECHAS
        If OptTraslados.Item(0).Value = True Then
            TxtTraslados.Visible = False
            TxtTraslados2.Visible = False
            LblTraslados.Caption = ""
            Lbllabel.Item(0).Visible = True
            Lbllabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'DOCUMENTO
        ElseIf OptTraslados.Item(1).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = False
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Numero De Documento"
            Lbllabel.Item(0).Visible = False
            Lbllabel.Item(1).Visible = False
            DtpTraFecIni.Visible = False
            DtpTraFecFin.Visible = False
        'FECHAS Y BODEGA DE SALIDA Y CODIGO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(2).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = True
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Materia Prima"
            LblTraslados2.Caption = "Codigo Bodega Salida"
            Lbllabel.Item(0).Visible = True
            Lbllabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'FECHAS Y BODEGA DE ENTRADA Y CODIGO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(3).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = True
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Materia Prima"
            LblTraslados2.Caption = "Codigo Bodega Entrada"
            Lbllabel.Item(0).Visible = True
            Lbllabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'FECHAS Y BODEGA DE SALIDA Y TIPO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(4).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = True
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Bodega Salida"
            LblTraslados2.Caption = "Codigo Tipo De Materia Prima"
            Lbllabel.Item(0).Visible = True
            Lbllabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'FECHAS Y BODEGA DE ENTRADA Y TIPO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(5).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = True
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Bodega Entrada"
            LblTraslados2.Caption = "Codigo Tipo De Materia Prima"
            Lbllabel.Item(0).Visible = True
            Lbllabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'FECHAS Y CODIGO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(6).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = False
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Materia Prima"
            Lbllabel.Item(0).Visible = True
            Lbllabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'NUMERO DE INGRESO
        ElseIf OptTraslados.Item(7).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = False
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Numero De Documento"
            Lbllabel.Item(0).Visible = False
            Lbllabel.Item(1).Visible = False
            DtpTraFecIni.Visible = False
            DtpTraFecFin.Visible = False
        'MATERIA PRIMA
        ElseIf OptTraslados.Item(8).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = False
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Materia Prima"
            Lbllabel.Item(0).Visible = False
            Lbllabel.Item(1).Visible = False
            DtpTraFecIni.Visible = False
            DtpTraFecFin.Visible = False
        'NO LIBERADO
        ElseIf OptTraslados.Item(9).Value = True Then
            TxtTraslados.Visible = False
            TxtTraslados2.Visible = False
            LblTraslados.Caption = ""
            Lbllabel.Item(0).Visible = False
            Lbllabel.Item(1).Visible = False
            DtpTraFecIni.Visible = False
            DtpTraFecFin.Visible = False
        End If
End Sub

Private Sub TabReportes_Click(PreviousTab As Integer)
            If TabReportes.Tab = 0 Then
                    OptInv.Item(0).Value = True
            ElseIf TabReportes.Tab = 1 Then
                    OptEntradas.Item(0).Value = True
            ElseIf TabReportes.Tab = 2 Then
                    OptTraslados.Item(0).Value = True
            ElseIf TabReportes.Tab = 3 Then
                    OptSalidas.Item(0).Value = True
            ElseIf TabReportes.Tab = 4 Then
                    OptCerrarBulto.Item(0).Value = True
            ElseIf TabReportes.Tab = 5 Then
                    OptCorrelativos.Item(0).Value = True
            ElseIf TabReportes.Tab = 6 Then
                    OptDesperdicio.Item(0).Value = True
            ElseIf TabReportes.Tab = 7 Then
                    OptBodegas.Item(0).Value = True
            End If
End Sub


Private Sub TxtBodegas_GotFocus()
        TxtBodegas.SelStart = 0
        TxtBodegas.SelLength = Len(TxtBodegas.Text)
End Sub


Private Sub Txtbusqueda_Change()
        'BODEGAS EN CARPETA DE INVENTARIO, TRASLADOS
        If (BInvBodega = True Or BTraBodega = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima Where CodigoBodega Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima Where CodigoBodega Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
        'BODEGAS GRUPOS EN CARPETA DE INVENTARIO
        ElseIf (BInvBodegaGrupo = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos Where Codigo Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos Where Codigo Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
            
            
            
        'TIPOS DE MATERIA PRIMA EN CARPETA DE INVENTARIO
        ElseIf (BInvTipoMateriaPrima = True Or BInvTipoMateriaPrima2 = True Or BTraTipoMateriaPrima = True Or BCerBulTipoDeMateriaPrima = True Or BEntTipoMateriaPrima = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima Where CodigoTipo Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima Where CodigoTipo Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
            
            
        'PROVEEDORES
        ElseIf (BEntProveedor = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoProveedor, Proveedor From Proveedores Where CodigoProveedor Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoProveedor, Proveedor From Proveedores Where CodigoProveedor Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoProveedor, Proveedor From Proveedores Where Proveedor Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoProveedor, Proveedor From Proveedores Where Proveedor Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
        
        
        'CLIENTES
        ElseIf (BSalCliente = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoCliente, Descripcion From Clientes Where CodigoCliente Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoCliente, Descripcion From Clientes Where CodigoCliente Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoCliente, Descripcion From Clientes Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoCliente, Descripcion From Clientes Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
      
            
        'LINEA
        ElseIf BCerBulLinea = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descripcion From Lineas Where Linea Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descripcion From Lineas Where Linea Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descripcion From Lineas Where Descrip Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select Linea, Descripcion From Lineas Where Descrip Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
            
            
        'MATERIA PRIMA
        ElseIf (BInvMateriaPrima = True Or BTraMateriaPrima = True Or BEntMateriaPrima = True Or BSalMateriaPrima = True Or BCerBulMateriaPrima = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
            
            
        'DESPERDICIO
        ElseIf BDesProceso = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoProceso, Descripcion From ProcesosMateriaPrima Where CodigoProceso Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoProceso, Descripcion From ProcesosMateriaPrima Where CodigoProceso Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoProceso, Descripcion From ProcesosMateriaPrima Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoProceso, Descripcion From ProcesosMateriaPrima Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
            
            
        'FICHA TECNICA
        ElseIf BDesFichaTecnica = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
        'DOCUMENTOS
        ElseIf BTraDocumentos = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoDocumento, Descripcion From Documentos Where CodigoDocumento Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoDocumento, Descripcion From Documentos Where CodigoDocumento Like '*" & Txtbusqueda.Text & "*'"
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                'PALABRA INICIAL
                If OptBusqueda.Item(2).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoDocumento, Descripcion From Documentos Where Descripcion Like '" & Txtbusqueda.Text & "*'"
                'CUALQUIER PALABRA
                ElseIf OptBusqueda.Item(3).Value = True Then
                    DataBusqueda.RecordSource = "Select CodigoDocumento, Descripcion From Documentos Where Descripcion Like '*" & Txtbusqueda.Text & "*'"
                End If
            End If
        
        
        End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "4000"
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

Private Sub TxtCerrarBulto_Change()
        'BUSCA POR LINEA
        If OptCerrarBulto.Item(2).Value = True Then
                Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtCerrarBulto.Text & "'")
                    If RBuscaLinea.RecordCount > 0 Then
                        LblCerBulDes.Caption = RBuscaLinea!Descrip
                    Else
                        LblCerBulDes.Caption = ""
                    End If
        'BUSCA POR MATERIA PRIMA
        ElseIf (OptCerrarBulto.Item(3).Value = True Or OptCerrarBulto.Item(4).Value = True) Then
                Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCerrarBulto.Text & "'")
                        If RBuscaMateriaPrima.RecordCount > 0 Then
                            LblCerBulDes.Caption = RBuscaMateriaPrima!Descripcion
                        Else
                            LblCerBulDes.Caption = ""
                        End If
                        
        'BUSCA TIPO DE MATERIA PRIMA
        ElseIf OptCerrarBulto.Item(6).Value = True Then
            Set RBuscaTipoMateriaPrima = Db.OpenRecordset("Select Descripcion From TiposDeMateriaPrima Where CodigoTipo = '" & TxtCerrarBulto.Text & "'")
                If RBuscaTipoMateriaPrima.RecordCount > 0 Then
                    LblCerBulDes.Caption = RBuscaTipoMateriaPrima!Descripcion
                Else
                    LblCerBulDes.Caption = ""
                End If

        Else
                LblCerBulDes.Caption = ""
        
        End If
End Sub

Private Sub TxtCerrarBulto_DblClick()
        'OPCION POR LINEA
        If OptCerrarBulto.Item(2).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = True
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesProceso = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        'OPCION POR MATERIA PRIMA
        ElseIf (OptCerrarBulto.Item(3).Value = True Or OptCerrarBulto.Item(4).Value = True) Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = True
                    BDesProceso = False
                    BDesProceso = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        'TIPO DE MATERIA PRIMA
        ElseIf OptCerrarBulto.Item(6).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesProceso = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = True
                    
        End If
            
        
        If (OptCerrarBulto.Item(2).Value = True Or OptCerrarBulto.Item(3).Value = True Or OptCerrarBulto.Item(4).Value = True Or OptCerrarBulto.Item(6).Value = True) Then
                    'OPCION DE LINEAS EN CARPETA DE CERRAR BULTO
                    If BCerBulLinea = True Then
                            DataBusqueda.RecordSource = "Select * From Lineas"
                    'OPCION DE MATERIA PRIMA EN CARPETA DE CERRAR BULTO
                    ElseIf BCerBulMateriaPrima = True Then
                            DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
                    ElseIf BCerBulTipoDeMateriaPrima = True Then
                    'OPCION DE TIPO DE MATERIA PRIMA EN CERRAR BULTO
                            DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                    End If
                            DataBusqueda.Refresh
                            DBGridBusqueda.Refresh
                            DBGridBusqueda.Columns(1).Width = "3000"
                            Columnas
                            FrameBusqueda.Visible = True
                            Txtbusqueda.SetFocus
        End If

End Sub

Private Sub TxtCerrarBulto_GotFocus()
        TxtCerrarBulto.SelStart = 0
        TxtCerrarBulto.SelLength = Len(TxtCerrarBulto.Text)
End Sub

Private Sub TxtCerrarBulto_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                'OPCION POR LINEA
                If OptCerrarBulto.Item(2).Value = True Then
                            BInvBodega = False
                            BInvTipoMateriaPrima = False
                            BInvMateriaPrima = False
                            BInvBodegaGrupo = False
                            BTraBodega = False
                            BTraMateriaPrima = False
                            BTraDocumentos = False
                            BEntProveedor = False
                            BEntMateriaPrima = False
                            BEntTipoMateriaPrima = False
                            BCerBulLinea = True
                            BCerBulMateriaPrima = False
                            BDesProceso = False
                            BDesProceso = False
                            BTraTipoMateriaPrima = False
                            BCerBulTipoDeMateriaPrima = False
                            
                'OPCION POR MATERIA PRIMA
                ElseIf (OptCerrarBulto.Item(3).Value = True Or OptCerrarBulto.Item(4).Value = True) Then
                            BInvBodega = False
                            BInvTipoMateriaPrima = False
                            BInvMateriaPrima = False
                            BInvBodegaGrupo = False
                            BTraBodega = False
                            BTraMateriaPrima = False
                            BTraDocumentos = False
                            BEntProveedor = False
                            BEntMateriaPrima = False
                            BEntTipoMateriaPrima = False
                            BCerBulLinea = False
                            BCerBulMateriaPrima = True
                            BDesProceso = False
                            BDesProceso = False
                            BTraTipoMateriaPrima = False
                            BCerBulTipoDeMateriaPrima = False
                            
                'TIPO DE MATERIA PRIMA
                ElseIf OptCerrarBulto.Item(6).Value = True Then
                            BInvBodega = False
                            BInvTipoMateriaPrima = False
                            BInvMateriaPrima = False
                            BInvBodegaGrupo = False
                            BTraBodega = False
                            BTraMateriaPrima = False
                            BTraDocumentos = False
                            BEntProveedor = False
                            BEntMateriaPrima = False
                            BEntTipoMateriaPrima = False
                            BSalMateriaPrima = False
                            BSalCliente = False
                            BCerBulLinea = False
                            BCerBulMateriaPrima = False
                            BDesProceso = False
                            BDesProceso = False
                            BTraTipoMateriaPrima = False
                            BCerBulTipoDeMateriaPrima = True
                            
                        End If
                
                If (OptCerrarBulto.Item(2).Value = True Or OptCerrarBulto.Item(3).Value = True Or OptCerrarBulto.Item(4).Value = True Or OptCerrarBulto.Item(6).Value = True) Then
                            'OPCION DE LINEAS EN CARPETA DE CERRAR BULTO
                            If BCerBulLinea = True Then
                                    DataBusqueda.RecordSource = "Select * From Lineas"
                            'OPCION DE MATERIA PRIMA EN CARPETA DE CERRAR BULTO
                            ElseIf BCerBulMateriaPrima = True Then
                                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
                             ElseIf BCerBulTipoDeMateriaPrima = True Then
                            'OPCION DE TIPO DE MATERIA PRIMA EN CERRAR BULTO
                                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                            End If
                                    DataBusqueda.Refresh
                                    DBGridBusqueda.Refresh
                                    DBGridBusqueda.Columns(1).Width = "3000"
                                    Columnas
                                    FrameBusqueda.Visible = True
                                    Txtbusqueda.SetFocus
                End If
        End If
End Sub

Private Sub TxtCorrelativos_Change()
        If OptCorrelativos.Item(0).Value = True Then
            'BUSCA CODIGO DE MATERIA PRIMA
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCorrelativos.Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblDesCor.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblDesCor.Caption = ""
                End If
        Else
            'BUSCA TIPO DE MATERIA PRIMA
            Set RBuscaTipoMateriaPrima = Db.OpenRecordset("Select Descripcion From TiposDeMateriaPrima Where CodigoTipo = '" & TxtCorrelativos.Text & "'")
                If RBuscaTipoMateriaPrima.RecordCount > 0 Then
                    LblDesCor.Caption = RBuscaTipoMateriaPrima!Descripcion
                Else
                    LblDesCor.Caption = ""
                End If

        End If
End Sub

Private Sub TxtDesperdicio_Change()
        'PROCESO
        If OptDesperdicio.Item(1).Value = True Then
                Set RBuscaProceso = Db.OpenRecordset("Select Descripcion From ProcesosMateriaPrima Where CodigoProceso = '" & TxtDesperdicio.Text & "'")
                    If RBuscaProceso.RecordCount > 0 Then
                        LblDesDes.Caption = RBuscaProceso!Descripcion
                    Else
                        LblDesDes.Caption = ""
                    End If
        'FICHA TECNICA
        ElseIf OptDesperdicio.Item(2).Value = True Then
                Set RBuscaFichaTecnica = Db.OpenRecordset("Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtDesperdicio.Text & "'")
                    If RBuscaFichaTecnica.RecordCount > 0 Then
                        LblDesDes.Caption = RBuscaFichaTecnica!Descrip
                    Else
                        LblDesDes.Caption = ""
                    End If
        End If
End Sub

Private Sub TxtDesperdicio_DblClick()
        'OPCION POR PROCESO
        If OptDesperdicio.Item(1).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = True
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
        'FICHA TECNICA
        ElseIf OptDesperdicio.Item(2).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = True
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
        End If
    
        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
        If BDesProceso = True Then
                DataBusqueda.RecordSource = "Select * From ProcesosMateriaPrima"
        'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
        ElseIf BDesFichaTecnica = True Then
                DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica"
        End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus

End Sub

Private Sub TxtDesperdicio_GotFocus()
        TxtDesperdicio.SelStart = 0
        TxtDesperdicio.SelLength = Len(TxtDesperdicio.Text)
End Sub

Private Sub TxtDesperdicio_KeyPress(KeyAscii As Integer)

If KeyAscii = 43 Then
        'OPCION POR PROCESO
        If OptDesperdicio.Item(1).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = True
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
        'FICHA TECNICA
        ElseIf OptDesperdicio.Item(2).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = True
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
        End If
    
        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
        If BDesProceso = True Then
                DataBusqueda.RecordSource = "Select * From ProcesosMateriaPrima"
        'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
        ElseIf BDesFichaTecnica = True Then
                DataBusqueda.RecordSource = "Select Esp_Tec, Descrip From FichaTecnica"
        End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
End If
End Sub

Private Sub TxtEntradas_Change()
        'BUSCA PROVEEDOR
        If OptEntradas.Item(1).Value = True Then
            Set RBuscaProveedor = Db.OpenRecordset("Select Proveedor From Proveedores Where CodigoProveedor = '" & TxtEntradas.Text & "'")
                If RBuscaProveedor.RecordCount > 0 Then
                    LblEntDes.Caption = RBuscaProveedor!Proveedor
                Else
                    LblEntDes.Caption = ""
                End If
        'BUSCA CODIGO DE MATERIA PRIMA
        ElseIf (OptEntradas.Item(2).Value = True Or OptEntradas.Item(3).Value = True) Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtEntradas.Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblEntDes.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblEntDes.Caption = ""
                End If
        'TIPO DE MATERIA PRIMA
        ElseIf OptEntradas.Item(6).Value = True Then
            Set RBuscaTipo = Db.OpenRecordset("Select Descripcion From TiposDeMateriaPrima Where CodigoTipo = '" & TxtEntradas.Text & "'")
                If RBuscaTipo.RecordCount > 0 Then
                    LblEntDes.Caption = RBuscaTipo!Descripcion
                Else
                    LblEntDes.Caption = ""
                End If
        Else
                    LblEntDes.Caption = ""
        End If
End Sub

Private Sub TxtEntradas_DblClick()
        'OPCION POR PROVEEDOR
        If OptEntradas.Item(1).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = True
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        'OPCION POR MATERIA PRIMA
        ElseIf (OptEntradas.Item(2).Value = True Or OptEntradas.Item(3).Value = True) Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = True
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
        'OPCION POR TIPO DE MATERIA PRIMA
        ElseIf OptEntradas.Item(6).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = True
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
        
                    
        End If
    
        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
        If BEntProveedor = True Then
                DataBusqueda.RecordSource = "Select * From Proveedores"
        'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
        ElseIf BEntMateriaPrima = True Then
                DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
        ElseIf BEntTipoMateriaPrima = True Then
                DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
        End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                If BEntProveedor = True Then
                    DBGridBusqueda.Columns(2).Width = "3000"
                ElseIf (BEntMateriaPrima = True Or BEntTipoMateriaPrima = True) Then
                    DBGridBusqueda.Columns(1).Width = "3000"
                End If
                Columnas
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus

End Sub

Private Sub TxtEntradas_GotFocus()
        TxtEntradas.SelStart = 0
        TxtEntradas.SelLength = Len(TxtEntradas.Text)
        
End Sub

Private Sub TxtEntradas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
                'OPCION POR PROVEEDOR
                If OptEntradas.Item(1).Value = True Then
                            BInvBodega = False
                            BInvTipoMateriaPrima = False
                            BInvMateriaPrima = False
                            BInvBodegaGrupo = False
                            BTraBodega = False
                            BTraMateriaPrima = False
                            BTraDocumentos = False
                            BEntProveedor = True
                            BEntMateriaPrima = False
                            BEntTipoMateriaPrima = False
                            BSalMateriaPrima = False
                            BSalCliente = False
                            BCerBulLinea = False
                            BCerBulMateriaPrima = False
                            BDesProceso = False
                            BDesFichaTecnica = False
                            BTraTipoMateriaPrima = False
                            BCerBulTipoDeMateriaPrima = False
                        
                'OPCION POR MATERIA PRIMA
                ElseIf (OptEntradas.Item(2).Value = True Or OptEntradas.Item(3).Value = True) Then
                            BInvBodega = False
                            BInvTipoMateriaPrima = False
                            BInvMateriaPrima = False
                            BTraBodega = False
                            BTraMateriaPrima = False
                            BTraDocumentos = False
                            BEntProveedor = False
                            BEntMateriaPrima = True
                            BEntTipoMateriaPrima = False
                            BSalMateriaPrima = False
                            BSalCliente = False
                            BCerBulLinea = False
                            BCerBulMateriaPrima = False
                            BDesProceso = False
                            BDesFichaTecnica = False
                            BTraTipoMateriaPrima = False
                            BCerBulTipoDeMateriaPrima = False
                        
                End If
            
                'OPCION DE BODEGA EN CARPETA DE TRASLADOS
                If BEntProveedor = True Then
                        DataBusqueda.RecordSource = "Select * From Proveedores"
                'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
                ElseIf BEntMateriaPrima = True Then
                        DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
                ElseIf BEntTipoMateriaPrima = True Then
                        DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                End If
                        DataBusqueda.Refresh
                        DBGridBusqueda.Refresh
                        If BEntProveedor = True Then
                            DBGridBusqueda.Columns(2).Width = "3000"
                        ElseIf (BEntMateriaPrima = True Or BEntTipoMateriaPrima = True) Then
                            DBGridBusqueda.Columns(1).Width = "3000"
                        End If
                        Columnas
                        FrameBusqueda.Visible = True
                        Txtbusqueda.SetFocus
        End If
End Sub

Private Sub TxtInv_Change()
        'BUSCA CODIGO DE MATERIA PRIMA
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtInv.Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblInvDesCod.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblInvDesCod.Caption = ""
                End If
End Sub

Private Sub TxtInv_DblClick()
        'OPCION DE MATERIA PRIMA EN CARPETA DE INVENTARIO
                BInvBodega = False
                BInvTipoMateriaPrima = False
                BInvMateriaPrima = True
                BInvTipoMateriaPrima2 = False
                BInvBodegaGrupo = False
                BTraBodega = False
                BTraMateriaPrima = False
                BTraDocumentos = False
                BSalMateriaPrima = False
                BSalCliente = False
                BEntMateriaPrima = False
                BEntTipoMateriaPrima = False
                BEntProveedor = False
                BCerBulLinea = False
                BCerBulMateriaPrima = False
                BDesProceso = False
                BDesFichaTecnica = False
                BTraTipoMateriaPrima = False
                BCerBulTipoDeMateriaPrima = False
                
                DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                Columnas
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
End Sub

Private Sub TxtInv_KeyPress(KeyAscii As Integer)
        'SI PRECIONA LA TECLA DE ENTER
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        'SI PRECIONA LA TECLA DEL SIGNO '+'
        If KeyAscii = 43 Then
                'OPCION DE MATERIA PRIMA EN CARPETA DE INVENTARIO
                BInvBodega = False
                BInvTipoMateriaPrima = False
                BInvMateriaPrima = True
                BInvTipoMateriaPrima2 = False
                BInvBodegaGrupo = False
                BTraBodega = False
                BTraMateriaPrima = False
                BTraDocumentos = False
                BSalMateriaPrima = False
                BEntMateriaPrima = False
                BEntTipoMateriaPrima = False
                BSalCliente = False
                BEntProveedor = False
                BCerBulLinea = False
                BCerBulMateriaPrima = False
                BDesProceso = False
                BDesFichaTecnica = False
                BTraTipoMateriaPrima = False
                BCerBulTipoDeMateriaPrima = False
                                
                DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                Columnas
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
        
        End If

End Sub

Private Sub TxtInvOpc_Change()
        'BUSCA LA BODEGA
        If (OptInvOpc.Item(1).Value = True Or OptInvOpc.Item(3).Value = True) Then
            Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtInvOpc.Text & "'")
                If RBuscaBodega.RecordCount > 0 Then
                    LblInv2.Caption = RBuscaBodega!Descripcion
                Else
                    LblInv2.Caption = ""
                End If
        'BUSCA TIPO DE MATERIA PRIMA
        ElseIf OptInvOpc.Item(2).Value = True Then
            Set RBuscaTipoMateriaPrima = Db.OpenRecordset("Select Descripcion From TiposDeMateriaPrima Where CodigoTipo = '" & TxtInvOpc.Text & "'")
                If RBuscaTipoMateriaPrima.RecordCount > 0 Then
                    LblInv2.Caption = RBuscaTipoMateriaPrima!Descripcion
                Else
                    LblInv2.Caption = ""
                End If
        'BODEGAS GRUPO
        ElseIf OptInvOpc.Item(4).Value = True Then
            Set RBuscaBodegaGrupo = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrimaGrupos Where Codigo = '" & TxtInvOpc.Text & "'")
                If RBuscaBodegaGrupo.RecordCount > 0 Then
                    LblInv2.Caption = RBuscaBodegaGrupo!Descripcion
                Else
                    LblInv2.Caption = ""
                End If
        Else
                    LblInv2.Caption = ""
        End If
        
End Sub

Private Sub TxtInvOpc_DblClick()
        'OPCION POR BODEGA
        If (OptInvOpc.Item(1).Value = True Or OptInvOpc.Item(3).Value = True) Then
                    BInvBodega = True
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    

        'OPCION POR TIPO DE MATERIA PRIMA
        ElseIf OptInvOpc.Item(2).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = True
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
        'OPCION POR BODEGAS GRUPO
        ElseIf OptInvOpc.Item(4).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = True
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        End If
    
        'OPCION DE BODEGA EN CARPETA DE INVENTARIO
        If BInvBodega = True Then
                DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima"
        'OPCION DE TIPO DE MATERIA PRIMA EN CARPETA DE INVENTARIO
        ElseIf BInvTipoMateriaPrima = True Then
                DataBusqueda.RecordSource = "Select CodigoTipo, Descripcion From TiposDeMateriaPrima"
        'OPCION DE GRUPO DE BODEGAS
        ElseIf BInvBodegaGrupo = True Then
                DataBusqueda.RecordSource = "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos"
        End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "3000"
                Columnas
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
        
End Sub

Private Sub TxtInvOpc_KeyPress(KeyAscii As Integer)
            
    'SI PRECIONA LA TECLA ENTER
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
        
    'SI PRECIONA LA TECLA DEL SIGNO '+'
    If KeyAscii = 43 Then
        'OPCION POR BODEGA
        If (OptInvOpc.Item(1).Value = True Or OptInvOpc.Item(3).Value = True) Then
                    BInvBodega = True
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BCerBulTipoDeMateriaPrima = False
                    
        'OPCION POR TIPO DE MATERIA PRIMA
        ElseIf OptInvOpc.Item(2).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = True
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BCerBulTipoDeMateriaPrima = False
        'OPCION POR BODEGAS GRUPO
        ElseIf OptInvOpc.Item(4).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = True
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
        
                    
        End If
    
        'OPCION DE BODEGA EN CARPETA DE INVENTARIO
        If BInvBodega = True Then
                DataBusqueda.RecordSource = "Select CodigoBodega, Descripcion From BodegasMateriaPrima"
        'OPCION DE TIPO DE MATERIA PRIMA EN CARPETA DE INVENTARIO
        ElseIf BInvTipoMateriaPrima = True Then
                DataBusqueda.RecordSource = "Select CodigoTipo, Descripcion From TiposDeMateriaPrima"
        'OPCION DE GRUPO DE BODEGAS
        ElseIf BInvBodegaGrupo = True Then
                DataBusqueda.RecordSource = "Select Codigo, Descripcion From BodegasMateriaPrimaGrupos"
        End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "3000"
                Columnas
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
    End If

End Sub

Private Sub TxtInvOpc2_Change()
        'BUSCA TIPO DE MATERIA PRIMA
            Set RBuscaTipoMateriaPrima = Db.OpenRecordset("Select Descripcion From TiposDeMateriaPrima Where CodigoTipo = '" & TxtInvOpc2.Text & "'")
                If RBuscaTipoMateriaPrima.RecordCount > 0 Then
                    LblInvDesTipMatPri.Caption = RBuscaTipoMateriaPrima!Descripcion
                Else
                    LblInvDesTipMatPri.Caption = ""
                End If
End Sub

Private Sub TxtInvOpc2_DblClick()
      'OPCION POR TIPO DE MATERIA PRIMA
        If OptInvOpc.Item(3).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvTipoMateriaPrima2 = True
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        End If
                DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "3000"
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
      
End Sub

Private Sub TxtInvOpc2_KeyPress(KeyAscii As Integer)
        
        
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            'OPCION POR TIPO DE MATERIA PRIMA
            If OptInvOpc.Item(3).Value = True Then
                        BInvBodega = False
                        BInvTipoMateriaPrima = False
                        BInvTipoMateriaPrima2 = True
                        BInvMateriaPrima = False
                        BInvBodegaGrupo = False
                        BTraBodega = False
                        BTraMateriaPrima = False
                        BTraDocumentos = False
                        BEntProveedor = False
                        BEntMateriaPrima = False
                        BEntTipoMateriaPrima = False
                        BSalMateriaPrima = False
                        BSalCliente = False
                        BCerBulLinea = False
                        BCerBulMateriaPrima = False
                        BDesProceso = False
                        BDesFichaTecnica = False
                        BTraTipoMateriaPrima = False
                        BCerBulTipoDeMateriaPrima = False
            End If
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
        End If
End Sub



Private Sub TxtSalidas_Change()
        'BUSCA cliente
        If OptSalidas.Item(1).Value = True Then
            Set RBuscaCliente = Db.OpenRecordset("Select Descripcion From Clientes Where CodigoCliente = '" & TxtSalidas.Text & "'")
                If RBuscaCliente.RecordCount > 0 Then
                    LblSalDes.Caption = RBuscaCliente!Descripcion
                Else
                    LblSalDes.Caption = ""
                End If
        'BUSCA CODIGO DE MATERIA PRIMA
        ElseIf (OptSalidas.Item(2).Value = True Or OptSalidas.Item(3).Value = True) Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtSalidas.Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblSalDes.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblSalDes.Caption = ""
                End If
        Else
                    LblSalDes.Caption = ""
        End If

End Sub

Private Sub TxtSalidas_DblClick()
        'OPCION POR CLIENTE
        If OptSalidas.Item(1).Value = True Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = True
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesProceso = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        'OPCION POR MATERIA PRIMA
        ElseIf (OptSalidas.Item(2).Value = True Or OptSalidas.Item(3).Value = True) Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = True
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesProceso = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        End If
    
        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
        If BSalCliente = True Then
                DataBusqueda.RecordSource = "Select * From Clientes"
        'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
        ElseIf BSalMateriaPrima = True Then
                DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
        End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                If BEntProveedor = True Then
                    DBGridBusqueda.Columns(2).Width = "3000"
                ElseIf BEntMateriaPrima = True Then
                    DBGridBusqueda.Columns(1).Width = "3000"
                End If
                Columnas
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus

End Sub

Private Sub TxtSalidas_GotFocus()
        TxtSalidas.SelStart = 0
        TxtSalidas.SelLength = Len(TxtSalidas.Text)
End Sub

Private Sub TxtSalidas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            'OPCION POR CLIENTE
                If OptSalidas.Item(1).Value = True Then
                            BInvBodega = False
                            BInvTipoMateriaPrima = False
                            BInvMateriaPrima = False
                            BInvBodegaGrupo = False
                            BTraBodega = False
                            BTraMateriaPrima = False
                            BTraDocumentos = False
                            BEntProveedor = False
                            BEntMateriaPrima = False
                            BEntTipoMateriaPrima = False
                            BSalMateriaPrima = False
                            BSalCliente = True
                            BCerBulLinea = False
                            BCerBulMateriaPrima = False
                            BDesProceso = False
                            BDesProceso = False
                            BTraTipoMateriaPrima = False
                            BCerBulTipoDeMateriaPrima = False
                            
                'OPCION POR MATERIA PRIMA
                ElseIf (OptSalidas.Item(2).Value = True Or OptSalidas.Item(3).Value = True) Then
                            BInvBodega = False
                            BInvTipoMateriaPrima = False
                            BInvMateriaPrima = False
                            BInvBodegaGrupo = False
                            BTraBodega = False
                            BTraMateriaPrima = False
                            BTraDocumentos = False
                            BEntProveedor = False
                            BEntMateriaPrima = False
                            BEntTipoMateriaPrima = False
                            BSalMateriaPrima = True
                            BSalCliente = False
                            BCerBulLinea = False
                            BCerBulMateriaPrima = False
                            BDesProceso = False
                            BDesProceso = False
                            BTraTipoMateriaPrima = False
                            BCerBulTipoDeMateriaPrima = False
                            
                End If
            
                'OPCION DE BODEGA EN CARPETA DE TRASLADOS
                If BSalCliente = True Then
                        DataBusqueda.RecordSource = "Select * From Clientes"
                'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
                ElseIf BSalMateriaPrima = True Then
                        DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
                End If
                        DataBusqueda.Refresh
                        DBGridBusqueda.Refresh
                        If BEntProveedor = True Then
                            DBGridBusqueda.Columns(2).Width = "3000"
                        ElseIf BEntMateriaPrima = True Then
                            DBGridBusqueda.Columns(1).Width = "3000"
                        End If
                        Columnas
                        FrameBusqueda.Visible = True
                        Txtbusqueda.SetFocus
        End If
        
End Sub

Private Sub TxtTraslados_Change()
        'BUSCA LA BODEGA SI LA OPCION DE BODEGA DE SALIDA O ENTRADA SON VERDADERAS
        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
            Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtTraslados.Text & "'")
                If RBuscaBodega.RecordCount > 0 Then
                    LblTraBod.Caption = RBuscaBodega!Descripcion
                Else
                    LblTraBod.Caption = ""
                End If
        'BUSCA CODIGO DE MATERIA PRIMA
        ElseIf (OptTraslados.Item(4).Value = True Or OptTraslados.Item(6).Value = True Or OptTraslados.Item(2).Value = True Or OptTraslados.Item(3).Value = True) Then
            Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTraslados.Text & "'")
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblTraBod.Caption = RBuscaMateriaPrima!Descripcion
                Else
                    LblTraBod.Caption = ""
                End If
        Else
                    LblTraBod.Caption = ""
        End If
End Sub


Sub Columnas()
        DBGridBusqueda.Columns(1).Width = "4000"
End Sub

Private Sub TxtTraslados_DblClick()
        'OPCION POR BODEGA
        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = True
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        'OPCION POR MATERIA PRIMA
        ElseIf (OptTraslados.Item(2).Value = True Or OptTraslados.Item(3).Value = True Or OptTraslados.Item(6).Value = True Or OptTraslados.Item(8).Value = True) Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = True
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
                    
        End If
    
        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
        If BTraBodega = True Then
                DataBusqueda.RecordSource = "Select * From BodegasMateriaPrima"
        'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
        ElseIf BTraMateriaPrima = True Then
                DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
        End If
                DataBusqueda.Refresh
                DBGridBusqueda.Refresh
                DBGridBusqueda.Columns(1).Width = "3000"
                Columnas
                FrameBusqueda.Visible = True
                Txtbusqueda.SetFocus
End Sub

Private Sub TxtTraslados_GotFocus()
        TxtTraslados.SelStart = 0
        TxtTraslados.SelLength = Len(TxtTraslados.Text)
End Sub

Private Sub TxtTraslados_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            'OPCION POR BODEGA
            If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
                        BInvBodega = False
                        BInvTipoMateriaPrima = False
                        BInvMateriaPrima = False
                        BInvBodegaGrupo = False
                        BTraBodega = True
                        BTraMateriaPrima = False
                        BTraDocumentos = False
                        BEntProveedor = False
                        BEntMateriaPrima = False
                        BEntTipoMateriaPrima = False
                        BSalMateriaPrima = False
                        BSalCliente = False
                        BCerBulLinea = False
                        BCerBulMateriaPrima = False
                        BDesProceso = False
                        BDesProceso = False
                        BTraTipoMateriaPrima = False
                        BCerBulTipoDeMateriaPrima = False
                        
            'OPCION POR MATERIA PRIMA
            ElseIf (OptTraslados.Item(2).Value = True Or OptTraslados.Item(3).Value = True Or OptTraslados.Item(6).Value = True Or OptTraslados.Item(8).Value = True) Then
                        BInvBodega = False
                        BInvTipoMateriaPrima = False
                        BInvMateriaPrima = False
                        BInvBodegaGrupo = False
                        BTraBodega = False
                        BTraMateriaPrima = True
                        BTraDocumentos = False
                        BEntProveedor = False
                        BEntMateriaPrima = False
                        BEntTipoMateriaPrima = False
                        BSalMateriaPrima = False
                        BSalCliente = False
                        BCerBulLinea = False
                        BCerBulMateriaPrima = False
                        BDesProceso = False
                        BDesProceso = False
                        BTraTipoMateriaPrima = False
                        BCerBulTipoDeMateriaPrima = False
                        
            End If
        
            'OPCION DE BODEGA EN CARPETA DE TRASLADOS
            If BTraBodega = True Then
                    DataBusqueda.RecordSource = "Select * From BodegasMateriaPrima"
            'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
            ElseIf BTraMateriaPrima = True Then
                    DataBusqueda.RecordSource = "Select * From CorrelativosMateriaPrima"
            End If
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "3000"
                    Columnas
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
        End If
End Sub

Sub Inventario()
            'PALABRA INICIAL
            If OptInvTipBus.Item(0).Value = True Then
                VTexto = " Like '" & TxtInv.Text & "*'"
            'CUALQUIER PALABRA
            ElseIf OptInvTipBus.Item(1).Value = True Then
                VTexto = " Like '*" & TxtInv.Text & "*'"
            'IGUAL
            ElseIf OptInvTipBus.Item(2).Value = True Then
                VTexto = " = '" & TxtInv.Text & "'"
            End If

            'CODIGO MATERIA PRIMA
            If OptInv.Item(0).Value = True Then
                        'TODOS
                        If OptInvOpc.Item(0).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Total De Todas Las Bodegas Y Tipos De Materia Prima'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto
                            End If
                        'BodegaDisponibilidad
                        ElseIf OptInvOpc.Item(1).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "'"
                            End If
                        'Tipo De Materia Prima
                        ElseIf OptInvOpc.Item(2).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Tipo De Materia Prima " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc.Text & "'"
                            End If
                        'Bodega Disponibilidad y Tipo De Materia Prima
                        ElseIf OptInvOpc.Item(3).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Tipo De Materia Prima " & TxtInvOpc2.Text & " " & LblInvDesTipMatPri.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc2.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc2.Text & "'"
                            End If
                        'GRUPO DE BODEGAS
                        ElseIf OptInvOpc.Item(4).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Grupo De Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {BodegasMateriaPrimaGrupos.Codigo} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {BodegasMateriaPrimaGrupos.Codigo} = '" & TxtInvOpc.Text & "'"
                            End If
                        'Bodega Disponibilidad y PASILLO
                        ElseIf OptInvOpc.Item(5).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Pasillo " & TxtInvOpc2.Text & " " & LblInvDesTipMatPri.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {DetalleEntradasMateriaPrima.Pasillo} = '" & TxtInvOpc2.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {DetalleEntradasMateriaPrima.Pasillo} = '" & TxtInvOpc2.Text & "'"
                            End If
                        
                        
                        End If
                
            'DESCRIPCION DE MATERIA PRIMA
            ElseIf OptInv.Item(1).Value = True Then
                    CrReportes.Formulas(0) = "Texto = 'Total Por Todas Bodegas Y Tipos De Materia Prima'"
                        'TODOS
                        If OptInvOpc.Item(0).Value = True Then
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto
                            End If
                        'BodegaDisponibilidad
                        ElseIf OptInvOpc.Item(1).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "'"
                            End If
                        'TIPO DE MATERIA PRIMA
                        ElseIf OptInvOpc.Item(2).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Tipo De Materia Prima " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc.Text & "'"
                            End If
                        'Bodega Disponibilidad y Tipo De Materia Prima
                        ElseIf OptInvOpc.Item(3).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Tipo De Materia Prima " & TxtInvOpc2.Text & " " & LblInvDesTipMatPri.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc2.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc2.Text & "'"
                            End If
                        'GRUPO DE BODEGAS
                        ElseIf OptInvOpc.Item(4).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Grupo De Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {BodegasMateriaPrimaGrupos.Codigo} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {BodegasMateriaPrimaGrupos.Codigo} = '" & TxtInvOpc.Text & "'"
                            End If
                        'Bodega Disponibilidad y PASILLO
                        ElseIf OptInvOpc.Item(5).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Pasillo " & TxtInvOpc2.Text & " " & LblInvDesTipMatPri.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {DetalleEntradasMateriaPrima.Pasillo} = '" & TxtInvOpc2.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.Descripcion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {DetalleEntradasMateriaPrima.Pasillo} = '" & TxtInvOpc2.Text & "'"
                            End If
                        
                        
                        End If
            'ORDEN
            ElseIf OptInv.Item(2).Value = True Then
                        'TODOS
                        If OptInvOpc.Item(0).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Total De Todas Las Bodegas Y Tipos De Materia Prima'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto
                            End If
                        'BodegaDisponibilidad
                        ElseIf OptInvOpc.Item(1).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "'"
                            End If
                        'Tipo De Materia Prima
                        ElseIf OptInvOpc.Item(2).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Tipo De Materia Prima " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc.Text & "'"
                            End If
                        'Bodega Disponibilidad y Tipo De Materia Prima
                        ElseIf OptInvOpc.Item(3).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Tipo De Materia Prima " & TxtInvOpc2.Text & " " & LblInvDesTipMatPri.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc2.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {CorrelativosMateriaPrima.TipoDeMateriaPrima} = '" & TxtInvOpc2.Text & "'"
                            End If
                        'GRUPOS DE BODEGAS
                        ElseIf OptInvOpc.Item(4).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Grupo De Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {BodegasMateriaPrimaGrupos.Codigo} = '" & TxtInvOpc.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {BodegasMateriaPrimaGrupos.Codigo} = '" & TxtInvOpc.Text & "'"
                            End If
                        'Bodega Disponibilidad y PASILLO
                        ElseIf OptInvOpc.Item(5).Value = True Then
                            CrReportes.Formulas(0) = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Pasillo " & TxtInvOpc2.Text & " " & LblInvDesTipMatPri.Caption & "'"
                            'REPORTE DE RESUMEN
                            If OptInvResDet.Item(0).Value = True Then
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {DetalleEntradasMateriaPrima.Pasillo} = '" & TxtInvOpc2.Text & "'"
                            'REPORTE DE DETALLE
                            Else
                                CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion}" & VTexto & " AND {DetalleEntradasMateriaPrima.BodegaDisponibilidad} = '" & TxtInvOpc.Text & "' AND {DetalleEntradasMateriaPrima.Pasillo} = '" & TxtInvOpc2.Text & "'"
                            End If
                        
                        End If
                
            End If 'FIN DE IF DE CODIGO Y DESCRIPCION
            
                                'MAYOR
                                If OptTipRep.Item(0).Value = True Then
                                        CrReportes.SelectionFormula = CrReportes.SelectionFormula & " And {DetalleEntradasMateriaPrima.SaldoDisponibilidad} > 0"
                                'MENOR
                                ElseIf OptTipRep.Item(1).Value = True Then
                                        'CrReportes.SelectionFormula = CrReportes.SelectionFormula & " And {DetalleEntradasMateriaPrima.SaldoDisponibilidad} < 0"
                                            'TODOS
                                            If OptInvOpc.Item(0).Value = True Then
                                                CrReportes.Formulas(0) = "Texto = 'Materias Primas Sin Existencia'"
                                            'BodegaDisponibilidad
                                            ElseIf OptInvOpc.Item(1).Value = True Then
                                                CrReportes.Formulas(0) = "Texto = 'Materias Primas Sin Existecnia'"
                                            ElseIf OptInvOpc.Item(2).Value = True Then
                                                CrReportes.Formulas(0) = "Texto = 'Materias Primas Sin Existecnia'"
                                            'Bodega Disponibilidad y Tipo De Materia Prima
                                            ElseIf OptInvOpc.Item(3).Value = True Then
                                                CrReportes.Formulas(0) = "Texto = 'Materias Primas Sin Existecnia'"
                                            End If
                                'IGUAL
                                ElseIf OptTipRep.Item(2).Value = True Then
                                        CrReportes.SelectionFormula = CrReportes.SelectionFormula & " And {DetalleEntradasMateriaPrima.SaldoDisponibilidad} = 0"
                                'TODOS
                                ElseIf OptTipRep.Item(3).Value = True Then
                                        'CrReportes.SelectionFormula = CrReportes.SelectionFormula & " And {DetalleEntradasMateriaPrima.SaldoDisponibilidad} = 0"
                                End If
                                
                                
                                'TIPO DE BODEGA _____________________________________________
                                'TODAS LAS BODEGAS
                                If OptInvPro.Item(0).Value = True Then
                                'BODEGAS DE PROCESO
                                ElseIf OptInvPro.Item(1).Value = True Then
                                    CrReportes.SelectionFormula = CrReportes.SelectionFormula & " And {BodegasMateriaPrima.EsBodegaDeProceso} = true"
                                'BODEGAS DE NO CONFORME
                                ElseIf OptInvPro.Item(2).Value = True Then
                                    CrReportes.SelectionFormula = CrReportes.SelectionFormula & " And {BodegasMateriaPrima.EsBodegaDeNoConforme} = true"
                                'BODEGAS QUE NO ESTEN EN PROCESO NI EN NO CONFORME
                                ElseIf OptInvPro.Item(3).Value = True Then
                                    CrReportes.SelectionFormula = CrReportes.SelectionFormula & " And {BodegasMateriaPrima.EsBodegaDeProceso} = false And {BodegasMateriaPrima.EsBodegaDeNoConforme} = false"
                                End If
                                
                                
                                
                                
                                    'RESUMEN BODEGA Y ORDEN
                                    If OptInvResDet.Item(0).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResBodega.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResNegativo.rpt"
                                        End If
                                    'DETALLADO POR BODEGA Y ORDEN
                                    ElseIf OptInvResDet.Item(1).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiDetBodega.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiDetNegativo.rpt"
                                        End If
                                    'RESUMEN X FICHA BODEGA Y ORDEN
                                    ElseIf OptInvResDet.Item(2).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResFicha.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResOrdenNegativo.rpt"
                                        End If
                                    'DETALLADO X FICHA BODEGA Y ORDEN
                                    ElseIf OptInvResDet.Item(3).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiDetFicha.rpt"
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiDetOrdenNegativo.rpt"
                                        End If
                                    'RESUMEN TIPO
                                    ElseIf OptInvResDet.Item(4).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResTipo.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResNegativoGeneral.rpt"
                                        End If
                                    'DETALLADO TIPO
                                    ElseIf OptInvResDet.Item(5).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiDetGeneral.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiDetNegativoGeneral.rpt"
                                        End If
                                    'RESUMEN BODEGA (SIMPLE BODEGA Y TIPO)
                                    ElseIf OptInvResDet.Item(6).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResBodegaTipo.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResNegativoBodegaTipo.rpt"
                                        End If
                                    'DETALLADO BODEGA (SIMPLE BODEGA Y TIPO)
                                    ElseIf OptInvResDet.Item(7).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiDetBodegaTipo.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiDetNegativoBodegaTipo.rpt"
                                        End If
                                    'CUADRICULA
                                    ElseIf OptInvResDet.Item(8).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResCuadricula.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResCuadricula.rpt"
                                        End If
                                     'CUADRICULA X ORDEN
                                    ElseIf OptInvResDet.Item(9).Value = True Then
                                        'MAYOR
                                        If OptTipRep.Item(0).Value = True Then
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResCuadriculaOrden.rpt"
                                        'MENOR
                                        Else
                                            CrReportes.ReportFileName = App.Path & "\RepInvMatPriExiResCuadriculaOrden.rpt"
                                        End If
                                    

                                    End If
End Sub

Sub Entradas()
            VDia = Day(DtpEntFecIni.Value)
            VMes = Month(DtpEntFecIni.Value)
            VAño = Year(DtpEntFecIni.Value)
            VDia2 = Day(DtpEntFecFin.Value)
            VMes2 = Month(DtpEntFecFin.Value)
            VAño2 = Year(DtpEntFecFin.Value)
            'FECHAS
            If OptEntradas.Item(0).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpEntFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpEntFecFin.Value, "dd/mm/yyyy") & "'"
                 CrReportes.SelectionFormula = "{EncabezadoEntradasMateriaPrima.FechaEntrada} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            'FECHAS Y PROVEEDOR
            ElseIf OptEntradas.Item(1).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpEntFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpEntFecFin.Value, "dd/mm/yyyy") & " Por Proveedor " & TxtEntradas.Text & " " & LblEntDes.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoEntradasMateriaPrima.FechaEntrada} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoEntradasMateriaPrima.Proveedor} Like '" & TxtEntradas.Text & "*'"
            'FECHAS Y MATERIA PRIMA
            ElseIf OptEntradas.Item(2).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpEntFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpEntFecFin.Value, "dd/mm/yyyy") & " Por Materia Prima " & TxtEntradas.Text & " " & LblEntDes.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoEntradasMateriaPrima.FechaEntrada} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleEntradasMateriaPrima.Codigo} Like '" & TxtEntradas.Text & "*'"
            'MATERIA PRIMA
            ElseIf OptEntradas.Item(3).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Por Materia Prima " & TxtEntradas.Text & " " & LblEntDes.Caption & "'"
                 CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.Codigo} Like '" & TxtEntradas.Text & "*'"
            'ENTRADAS NO LIBERADAS
            ElseIf OptEntradas.Item(4).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Entradas No Liberadas'"
                 CrReportes.SelectionFormula = "{EncabezadoEntradasMateriaPrima.Estado} = 'NO LIBERADA'"
            'ORDEN
            ElseIf OptEntradas.Item(5).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Por Orden De Produccion " & TxtEntradas.Text & " " & LblEntDes.Caption & "'"
                 CrReportes.SelectionFormula = "{DetalleEntradasMateriaPrima.OrdenProduccion} Like '" & TxtEntradas.Text & "*'"
            'FECHAS Y TIPO DE MATERIA PRIMA
            ElseIf OptEntradas.Item(6).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpEntFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpEntFecFin.Value, "dd/mm/yyyy") & " Tipo Materia Prima " & TxtEntradas.Text & " " & LblEntDes.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoEntradasMateriaPrima.FechaEntrada} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleEntradasMateriaPrima.Codigo} = {CorrelativosMateriaPrima.CodigoMateriaPrima} And {CorrelativosMateriaPrima.TipoDeMateriaPrima} = {TiposDeMateriaPrima.CodigoTipo} And {TiposDeMateriaPrima.CodigoTipo} Like '" & TxtEntradas.Text & "*'"
            
            
            End If
                 'TIPO DE REPORTE
                 'DETALLE
                 If OptEntDetalle.Value = True Then
                        CrReportes.ReportFileName = App.Path & "\ReporteEntradasMateriaPrima.rpt"
                 'RESUMEN
                 ElseIf OptEntResumen.Value = True Then
                        CrReportes.ReportFileName = App.Path & "\ReporteEntradasMateriaPrimaResumen.rpt"
                 'RESUMEN PROVEEDOR
                 ElseIf OptEntResPro.Value = True Then
                        CrReportes.ReportFileName = App.Path & "\ReporteEntradasMateriaPrimaResumenProveedor.rpt"
                 'RESUMEN CUADRICULA
                 ElseIf OptEntResCua.Value = True Then
                        CrReportes.ReportFileName = App.Path & "\ReporteEntradasMateriaPrimaResumenCuadricula.rpt"
                 End If
End Sub

Sub Traslados()
            VDia = Day(DtpTraFecIni.Value)
            VMes = Month(DtpTraFecIni.Value)
            VAño = Year(DtpTraFecIni.Value)
            VDia2 = Day(DtpTraFecFin.Value)
            VMes2 = Month(DtpTraFecFin.Value)
            VAño2 = Year(DtpTraFecFin.Value)
            'FECHAS
            If OptTraslados.Item(0).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & "'"
                 CrReportes.SelectionFormula = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                 'CrReportes.SelectionFormula = "{EncabezadoDevolucionesMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            'NUMERO DOCUMENTO
            ElseIf OptTraslados.Item(1).Value = True Then
                    CrReportes.Formulas(0) = "Texto = ' Numero De Documento " & TxtTraslados.Text & "'"
                    CrReportes.SelectionFormula = "{EncabezadoTrasladosMateriaPrim.NumeroDocumento} = " & TxtTraslados.Text
            'FECHAS Y BODEGA SALIDA Y CODIGO DE MATERIA PRIMA
            ElseIf OptTraslados.Item(2).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Bodega Salida " & TxtTraslados.Text & " " & LblTraBod.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleTrasladosMateriaPrimaP.CodigoSalida} Like '" & TxtTraslados.Text & "*' And {EncabezadoTrasladosMateriaPrim.BodegaSalida} Like '" & TxtTraslados2.Text & "*'"
            'FECHAS BODEGA ENTRADA Y CODIGO DE MATERIA PRIMA
            ElseIf OptTraslados.Item(3).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Bodega Entrada " & TxtTraslados.Text & " " & LblTraBod.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleTrasladosMateriaPrimaP.CodigoSalida} Like '" & TxtTraslados.Text & "*' And {DetalleTrasladosMateriaPrimaP.BodegaEntrada} Like '" & TxtTraslados2.Text & "*'"
            'FECHAS Y BODEGA SALIDA Y TIPO DE MATERIA PRIMA
            ElseIf OptTraslados.Item(4).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Bodega Salida " & TxtTraslados.Text & " " & LblTraBod.Caption & " Materia Prima " & LblTraBod2.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoTrasladosMateriaPrim.BodegaSalida} Like '" & TxtTraslados.Text & "*' And {CorrelativosMateriaPrima.TipoDeMateriaPrima} Like '" & TxtTraslados2.Text & "*'"
            'FECHAS BODEGA ENTRADA Y TIPO DE MATERIA PRIMA
            ElseIf OptTraslados.Item(5).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Bodega Entrada " & TxtTraslados.Text & " " & LblTraBod.Caption & " Materia Prima " & LblTraBod2.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleTrasladosMateriaPrimaP.BodegaEntrada} Like '" & TxtTraslados.Text & "*' And {CorrelativosMateriaPrima.TipoDeMateriaPrima} Like '" & TxtTraslados2.Text & "*'"
            'FECHAS CODIGO MATERIA PRIMA
            ElseIf OptTraslados.Item(6).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Codigo " & TxtTraslados.Text & " " & LblTraBod.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoTrasladosMateriaPrim.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleTrasladosMateriaPrimaP.CodigoSalida} Like '" & TxtTraslados.Text & "*'"
            'ORDEN
            ElseIf OptTraslados.Item(7).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Orden " & TxtTraslados.Text & "'"
                 CrReportes.SelectionFormula = "{DetalleTrasladosMateriaPrimaP.Orden} = '" & TxtTraslados.Text & "'"
            'MATERIA PRIMA
            ElseIf OptTraslados.Item(8).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Codigo Materia Prima " & TxtTraslados.Text & " " & LblTraBod.Caption & "'"
                 CrReportes.SelectionFormula = "{DetalleTrasladosMateriaPrimaP.CodigoSalida} Like '" & TxtTraslados.Text & "*'"
            'NO LIBERADO
            ElseIf OptTraslados.Item(9).Value = True Then
                 CrReportes.Formulas(0) = "Texto = 'Traslados No Liberados'"
                 CrReportes.SelectionFormula = "{EncabezadoTrasladosMateriaPrim.Estado} = 'NO LIBERADO'"
            End If
            
            'OPCION DE TODOS LOS TIPOS DE DOCUMENTO
            If OptTraOpc.Item(0).Value = True Then
                'POR UN TIPO DE DOCUMENTO
            Else
                CrReportes.SelectionFormula = CrReportes.SelectionFormula & " And {EncabezadoTrasladosMateriaPrim.TipoDeDocumento} = '" & TxtTraTipDoc.Text & "'"
            End If
        
            
            If OptTraDet.Value = True Then
                CrReportes.ReportFileName = App.Path & "\ReporteTrasladosMateriaPrima.rpt"
            Else
                CrReportes.ReportFileName = App.Path & "\ReporteTrasladosMateriaPrimaResumen.rpt"
            End If
            'CrReportes.ReportFileName = App.Path & "\ReporteDevolucionesMateriaPrima.rpt"
          
End Sub

Sub Salidas()
            VDia = Day(DtpSalFecIni.Value)
            VMes = Month(DtpSalFecIni.Value)
            VAño = Year(DtpSalFecIni.Value)
            VDia2 = Day(DtpSalFecFin.Value)
            VMes2 = Month(DtpSalFecFin.Value)
            VAño2 = Year(DtpSalFecFin.Value)
            'FECHAS
            If OptSalidas.Item(0).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpSalFecFin.Value, "dd/mm/yyyy") & "'"
                 CrReportes.SelectionFormula = "{EncabezadoEgresosMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            'FECHAS Y CLIENTE
            ElseIf OptSalidas.Item(1).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpEntFecFin.Value, "dd/mm/yyyy") & " Por Cliente " & TxtSalidas.Text & " " & LblSalDes.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoEgresosMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoEgresosMateriaPrima.Cliente} Like '" & TxtSalidas.Text & "*'"
            'FECHAS Y MATERIA PRIMA
            ElseIf OptSalidas.Item(2).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpSalFecFin.Value, "dd/mm/yyyy") & " Por Materia Prima " & TxtSalidas.Text & " " & LblSalDes.Caption & "'"
                 CrReportes.SelectionFormula = "{EncabezadoEgresosMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleEgresosMateriaPrima.Codigo} Like '" & TxtSalidas.Text & "*'"
            'MATERIA PRIMA
            ElseIf OptSalidas.Item(3).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Por Materia Prima " & TxtSalidas.Text & " " & LblSalDes.Caption & "'"
                 CrReportes.SelectionFormula = "{DetalleEgresosMateriaPrima.Codigo} Like '" & TxtSalidas.Text & "*'"
            'NO LIBERADO
            ElseIf OptSalidas.Item(4).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Salidas No Liberadas '"
                 CrReportes.SelectionFormula = "{EncabezadoEgresosMateriaPrima.Estado} = 'NO LIBERADO'"
            End If
            
                If OptDesDet.Value = True Then
                    CrReportes.ReportFileName = App.Path & "\ReporteSalidasMateriaPrima.rpt"
                ElseIf OptDesRes.Value = True Then
                    CrReportes.ReportFileName = App.Path & "\ReporteSalidasMateriaPrimaResumen.rpt"
                ElseIf OptDesResCli.Value = True Then
                    CrReportes.ReportFileName = App.Path & "\ReporteSalidasMateriaPrimaResumenCliente.rpt"
                End If

End Sub

Sub CerrarBulto()
            VDia = Day(DtpCerBulFecIni.Value)
            VMes = Month(DtpCerBulFecIni.Value)
            VAño = Year(DtpCerBulFecIni.Value)
            VDia2 = Day(DtpCerBulFecFin.Value)
            VMes2 = Month(DtpCerBulFecFin.Value)
            VAño2 = Year(DtpCerBulFecFin.Value)
            'FECHAS
            If OptCerrarBulto.Item(0).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpCerBulFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCerBulFecFin.Value, "dd/mm/yyyy") & "'"
                 CrReportes.SelectionFormula = "{NumerosIngresosProcesados.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            'FECHAS Y TURNO
            ElseIf OptCerrarBulto.Item(1).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpCerBulFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCerBulFecFin.Value, "dd/mm/yyyy") & " Del Turno " & TxtCerrarBulto.Text & "'"
                 CrReportes.SelectionFormula = "{NumerosIngresosProcesados.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {NumerosIngresosProcesados.Turno} Like '" & TxtCerrarBulto.Text & "*'"
            'FECHAS Y LINEA
            ElseIf OptCerrarBulto.Item(2).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpCerBulFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCerBulFecFin.Value, "dd/mm/yyyy") & " De la Linea " & TxtCerrarBulto.Text & " " & LblCerBulDes.Caption & "'"
                 CrReportes.SelectionFormula = "{NumerosIngresosProcesados.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {NumerosIngresosProcesados.Linea} Like '" & TxtCerrarBulto.Text & "*'"
            'FECHAS MATERIA PRIMA
            ElseIf OptCerrarBulto.Item(3).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpCerBulFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCerBulFecFin.Value, "dd/mm/yyyy") & " De Materia Prima " & TxtCerrarBulto.Text & " " & LblCerBulDes.Caption & "'"
                 CrReportes.SelectionFormula = "{NumerosIngresosProcesados.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {NumerosIngresosProcesados.CodigoMateriaPrima} Like '" & TxtCerrarBulto.Text & "*'"
            'MATERIA PRIMA
            ElseIf OptCerrarBulto.Item(4).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' De Materia Prima " & TxtCerrarBulto.Text & " " & LblCerBulDes.Caption & "'"
                 CrReportes.SelectionFormula = "{NumerosIngresosProcesados.CodigoMateriaPrima} Like '" & TxtCerrarBulto.Text & "*'"
            'NUMERO INGRESO
            ElseIf OptCerrarBulto.Item(5).Value = True Then
                 If Not IsNumeric(TxtCerrarBulto.Text) Then
                        MsgBox "Numero De Ingreso Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                 End If
                 CrReportes.Formulas(0) = "Texto = ' De Numero Ingreso " & TxtCerrarBulto.Text & " " & LblCerBulDes.Caption & "'"
                 CrReportes.SelectionFormula = "{NumerosIngresosProcesados.NumeroIngreso} = " & TxtCerrarBulto.Text
            'FECHAS Y TIPO DE MATERIA PRIMA
            ElseIf OptCerrarBulto.Item(6).Value = True Then
                 CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DtpCerBulFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCerBulFecFin.Value, "dd/mm/yyyy") & " De Materia Prima " & TxtCerrarBulto.Text & " " & LblCerBulDes.Caption & "'"
                 CrReportes.SelectionFormula = "{NumerosIngresosProcesados.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {CorrelativosMateriaPrima.TipoDeMateriaPrima} Like '" & TxtCerrarBulto.Text & "*'"
                 
            End If
            
            'TIPO DE REPORTE
            If OptCerrarBulto.Item(7).Value = True Then
                 CrReportes.ReportFileName = App.Path & "\ReporteCerrarBultoMateriaPrima.rpt"
            ElseIf OptCerrarBulto.Item(8).Value = True Then
                 CrReportes.ReportFileName = App.Path & "\ReporteCerrarBultoMateriaPrimaResumen.rpt"
            End If
End Sub

Sub CorrelativosMaximos()
            If OptCorrelativos.Item(0).Value = True Then
                 CrReportes.SelectionFormula = "{CorrelativosMateriaPrima.CodigoMateriaPrima} Like '" & TxtCorrelativos.Text & "*'"
            ElseIf OptCorrelativos.Item(1).Value = True Then
                CrReportes.SelectionFormula = "{CorrelativosMateriaPrima.TipodeMateriaPrima} Like '" & TxtCorrelativos.Text & "*'"
            End If
                 CrReportes.ReportFileName = App.Path & "\CorrelativosMateriaPrima.rpt"

End Sub

Private Sub TxtTraslados2_Change()
    If (OptTraslados.Item(2).Value = True Or OptTraslados.Item(3).Value = True) Then
        Set RBuscaBodega = Db.OpenRecordset("Select Descripcion From BodegasMateriaPrima Where CodigoBodega = '" & TxtTraslados2.Text & "'")
                        If RBuscaBodega.RecordCount > 0 Then
                            LblTraBod2.Caption = RBuscaBodega!Descripcion
                        Else
                            LblTraBod2.Caption = ""
                        End If
    
    Else
            'BUSCA TIPO DE MATERIA PRIMA
            Set RBuscaTipoMateriaPrima = Db.OpenRecordset("Select Descripcion From TiposDeMateriaPrima Where CodigoTipo = '" & TxtTraslados2.Text & "'")
                If RBuscaTipoMateriaPrima.RecordCount > 0 Then
                    LblTraBod2.Caption = RBuscaTipoMateriaPrima!Descripcion
                Else
                    LblTraBod2.Caption = ""
                End If
    End If
End Sub

Private Sub TxtTraslados2_DblClick()
  'OPCION POR TIPO DE MATERIA PRIMA
        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = True
                    BCerBulTipoDeMateriaPrima = False
                                        
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
        End If
        
End Sub

Private Sub TxtTraslados2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 43 Then
        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = True
                    BCerBulTipoDeMateriaPrima = False
                    
                    DataBusqueda.RecordSource = "Select * From TiposDeMateriaPrima"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
        End If
    End If

End Sub

Public Sub Desperdicio()
            VDia = Day(DTPDesFecIni.Value)
            VMes = Month(DTPDesFecIni.Value)
            VAño = Year(DTPDesFecIni.Value)
            VDia2 = Day(DTPDesFecFin.Value)
            VMes2 = Month(DTPDesFecFin.Value)
            VAño2 = Year(DTPDesFecFin.Value)
            
            
            If OptDesPro.Value = True Then
                    'FECHAS
                    If OptDesperdicio.Item(0).Value = True Then
                         CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy") & "'"
                         CrReportes.SelectionFormula = "{CapturaDesperdicioMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    'FECHAS Y PROCESO
                    ElseIf OptDesperdicio.Item(1).Value = True Then
                         CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy") & " Del Proceso " & TxtDesperdicio.Text & " " & LblDesDes.Caption & "'"
                         CrReportes.SelectionFormula = "{CapturaDesperdicioMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {CapturaDesperdicioMateriaPrima.CodigoProceso} like '" & TxtDesperdicio.Text & "*'"
                    'FECHAS Y FICHA TECNICA
                    ElseIf OptDesperdicio.Item(2).Value = True Then
                         CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy") & " De Ficha Tecnica " & TxtDesperdicio.Text & " " & LblDesDes.Caption & "'"
                         CrReportes.SelectionFormula = "{CapturaDesperdicioMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {CapturaDesperdicioMateriaPrima.FichaTecnica} Like '" & TxtDesperdicio.Text & "*'"
                    'FECHAS Y Grupo
                    ElseIf OptDesperdicio.Item(3).Value = True Then
                         CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy") & " Del Grupo " & TxtDesperdicio.Text & " " & LblDesDes.Caption & "'"
                         CrReportes.SelectionFormula = "{CapturaDesperdicioMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {CapturaDesperdicioMateriaPrima.CodigoProceso} = {ProcesosMateriaPrima.CodigoProceso}  And {ProcesosMateriaPrima.Grupo} like '" & TxtDesperdicio.Text & "*'"
                    End If
            Else
               'FECHAS
                    If OptDesperdicio.Item(0).Value = True Then
                         CrReportes.Formulas(0) = "Texto = ' Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy") & "'"
                         CrReportes.SelectionFormula = "{DesperdicioPacas.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    End If
            End If
                
                'ELIGE REPORTE DE ACUERDO A LA OPCION
                If OptDesPro.Value = True Then
                    If OptDetalle.Value = True Then
                        CrReportes.ReportFileName = App.Path & "\ReporteDesperdicioProcesoMateriaPrimaDetalle.rpt"
                    ElseIf OptResumen.Value = True Then
                        CrReportes.ReportFileName = App.Path & "\ReporteDesperdicioProcesoMateriaPrimaResumen.rpt"
                    'DETALLE FECHA
                    ElseIf OptDetalleFecha.Value = True Then
                        CrReportes.ReportFileName = App.Path & "\ReporteDesperdicioProcesoMateriaPrimaDetallexFecha.rpt"
                    'CUADRICULA
                    ElseIf OptDesCuaPro.Value = True Then
                        CrReportes.ReportFileName = App.Path & "\DesperdicioProcesoMateriaPrimaCuadricula.rpt"
                    End If
                Else
                    CrReportes.ReportFileName = App.Path & "\DesperdicioPacas.rpt"
                End If
    
End Sub

Private Sub TxtTraTipDoc_Change()
        Set RBuscaDocumento = Db.OpenRecordset("Select Descripcion From Documentos Where CodigoDocumento = '" & TxtTraTipDoc.Text & "'")
            If RBuscaDocumento.RecordCount > 0 Then
                LblTraDesDoc.Caption = RBuscaDocumento!Descripcion
            Else
                LblTraDesDoc.Caption = ""
            End If
End Sub

Private Sub TxtTraTipDoc_DblClick()
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = True
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
    
                    DataBusqueda.RecordSource = "Select * From Documentos"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
End Sub

Private Sub TxtTraTipDoc_KeyPress(KeyAscii As Integer)
                    
                If KeyAscii = 13 Then
                    SendKeys "{tab}"
                End If
                    
                If KeyAscii = 43 Then
                    BInvBodega = False
                    BInvTipoMateriaPrima = False
                    BInvMateriaPrima = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraMateriaPrima = False
                    BTraDocumentos = True
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalMateriaPrima = False
                    BSalCliente = False
                    BCerBulLinea = False
                    BCerBulMateriaPrima = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipoMateriaPrima = False
                    BCerBulTipoDeMateriaPrima = False
    
                    DataBusqueda.RecordSource = "Select * From Documentos"
                    DataBusqueda.Refresh
                    DBGridBusqueda.Refresh
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    Txtbusqueda.SetFocus
                End If
                
                
End Sub

Public Sub Bodegas()
            If OptBodegas.Item(0).Value = True Then
                 'CrReportes.SelectionFormula = "{BodegasMateriaPrima.CodigoBodega} Like '" & TxtBodegas.Text & "*'"
                 CrReportes.SelectionFormula = "{BodegasMateriaPrima.CodigoBodega} Like '" & TxtBodegas.Text & "*'"
            ElseIf OptBodegas.Item(1).Value = True Then
                CrReportes.SelectionFormula = "{BodegasMateriaPrima.Descripcion} Like '" & TxtBodegas.Text & "*'"
            ElseIf OptBodegas.Item(2).Value = True Then
                CrReportes.SelectionFormula = "{BodegasMateriaPrima.Grupo} Like '" & TxtBodegas.Text & "*'"
            End If
                 'CrReportes.ReportFileName = App.Path & "\BodegasMateriaPrima.rpt"
                 
                 CrReportes.ReportFileName = App.Path & "\BodegasMateriaPrima.rpt"

End Sub
