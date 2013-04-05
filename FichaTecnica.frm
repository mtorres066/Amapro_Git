VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FichaTecnica 
   BackColor       =   &H000000FF&
   Caption         =   "Fichas Tecnicas De Producto Terminado"
   ClientHeight    =   8595
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "FichaTecnica.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8535
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   62
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   59
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   60
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   10815
      End
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   7335
         Left            =   120
         TabIndex        =   61
         Top             =   1080
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12938
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   11040
         Picture         =   "FichaTecnica.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "FichaTecnica.frx":237C
      Picture         =   "FichaTecnica.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   71
      ToolTipText     =   "Primer Registro"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "FichaTecnica.frx":2CF0
      Picture         =   "FichaTecnica.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   70
      ToolTipText     =   "Registro Anterior"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   11040
      MouseIcon       =   "FichaTecnica.frx":3664
      Picture         =   "FichaTecnica.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Siguiente Registro"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   11400
      MouseIcon       =   "FichaTecnica.frx":3FD8
      Picture         =   "FichaTecnica.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   68
      ToolTipText     =   "Ultimo Registro"
      Top             =   7920
      Width           =   375
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   700
      Left            =   960
      MouseIcon       =   "FichaTecnica.frx":494C
      Picture         =   "FichaTecnica.frx":4D8E
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   7800
      Width           =   1600
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   700
      Left            =   2640
      MouseIcon       =   "FichaTecnica.frx":510B
      Picture         =   "FichaTecnica.frx":554D
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7800
      Width           =   1600
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   700
      Left            =   4320
      MouseIcon       =   "FichaTecnica.frx":5924
      Picture         =   "FichaTecnica.frx":5D66
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7800
      Width           =   1600
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   700
      Left            =   6000
      MouseIcon       =   "FichaTecnica.frx":62C2
      Picture         =   "FichaTecnica.frx":6704
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7800
      Width           =   1600
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   700
      Left            =   7680
      MouseIcon       =   "FichaTecnica.frx":6C3B
      Picture         =   "FichaTecnica.frx":707D
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7800
      Width           =   1600
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   700
      Left            =   9360
      MouseIcon       =   "FichaTecnica.frx":7645
      Picture         =   "FichaTecnica.frx":7A87
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7800
      Width           =   1600
   End
   Begin TabDlg.SSTab TabFichasTecnicas 
      Height          =   7695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   -2147483645
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "FichaTecnica.frx":7FA2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameFichaTecnica"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TxtDatos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "FichaTecnica.frx":82BC
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridFichaTecnica"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Seleccion O Busquedad De Datos"
      TabPicture(2)   =   "FichaTecnica.frx":870E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameOpciones"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DbGridFichaTecnica 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   67
         Top             =   720
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   12091
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   26
         BeginProperty Column00 
            DataField       =   "Esp_Tec"
            Caption         =   "Codigo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descrip"
            Caption         =   "Descripcion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Tipo"
            Caption         =   "Tipo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "TipoVenta"
            Caption         =   "TipoVenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Envases"
            Caption         =   "Uni.x Tar."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Nombre_Comercial"
            Caption         =   "Nom.Com"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Imp_Defe"
            Caption         =   "Imp.Def."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Imp_Cali"
            Caption         =   "Imp.Cal."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column08 
            DataField       =   "Atributos"
            Caption         =   "Atributos"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "Variables"
            Caption         =   "Catalogo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Origen"
            Caption         =   "Origen"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "UnidadMedida"
            Caption         =   "Uni.Med."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column12 
            DataField       =   "MaterialEmpaque"
            Caption         =   "Mat.Emp."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "Activa"
            Caption         =   "Activa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column14 
            DataField       =   "PesoxUnidad"
            Caption         =   "PesoxUnidad"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column15 
            DataField       =   "PesoxUnidadConTapa"
            Caption         =   "Peso C/Tapa"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column16 
            DataField       =   "UnidadesxLamina"
            Caption         =   "Uni.x.Lam"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column17 
            DataField       =   "UnidadesxCaja"
            Caption         =   "Uni.x.Caja"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column18 
            DataField       =   "TipoInventario"
            Caption         =   "Tipo Inventario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column19 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column20 
            DataField       =   "CodigoCliente"
            Caption         =   "Codigo De Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column21 
            DataField       =   "ProductoDelCliente"
            Caption         =   "Producto Del Cliente"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column22 
            DataField       =   "SegundoNombre"
            Caption         =   "Seg.Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column23 
            DataField       =   "TercerNombre"
            Caption         =   "Tercer Nombre"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column24 
            DataField       =   "TipoVenta"
            Caption         =   "TipoVenta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column25 
            DataField       =   "NumeroGrafica"
            Caption         =   "NumeroGrafica"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3614.74
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column14 
               Alignment       =   1
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column16 
               Alignment       =   1
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column17 
               Alignment       =   1
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
            EndProperty
            BeginProperty Column20 
            EndProperty
            BeginProperty Column21 
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column22 
            EndProperty
            BeginProperty Column23 
            EndProperty
            BeginProperty Column24 
            EndProperty
            BeginProperty Column25 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   3480
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   49
         Top             =   2520
         Width           =   8175
      End
      Begin VB.Frame FrameOpciones 
         Caption         =   "Opciones De Busqueda"
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
         TabIndex        =   33
         Top             =   720
         Width           =   11655
         Begin VB.TextBox TxtBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8880
            TabIndex        =   63
            ToolTipText     =   " "
            Top             =   3600
            Width           =   2565
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Buscar"
            Height          =   855
            Left            =   8880
            Picture         =   "FichaTecnica.frx":8B60
            Style           =   1  'Graphical
            TabIndex        =   64
            Top             =   4080
            Width           =   2535
         End
         Begin VB.CommandButton CmdActualizar 
            Caption         =   "Actualizar"
            Height          =   855
            Left            =   8880
            Picture         =   "FichaTecnica.frx":8FA2
            Style           =   1  'Graphical
            TabIndex        =   65
            Top             =   5040
            Width           =   2535
         End
         Begin VB.OptionButton OptTipFicTec 
            Caption         =   "Tipo Ficha Tecnica"
            Height          =   1095
            Left            =   4440
            Picture         =   "FichaTecnica.frx":92AC
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton OptCodigoVariable 
            Caption         =   "Codigo De Catalogo"
            Height          =   1095
            Left            =   2400
            Picture         =   "FichaTecnica.frx":AFA6
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   480
            Width           =   1935
         End
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Codigo Ficha Tecnica"
            Height          =   1095
            Left            =   360
            Picture         =   "FichaTecnica.frx":B3F0
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   480
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.Label LblFicha 
            Alignment       =   1  'Right Justify
            Caption         =   "Codigo De Ficha Tecnica"
            Height          =   255
            Left            =   6000
            TabIndex        =   66
            Top             =   3600
            Width           =   2775
         End
      End
      Begin VB.Frame FrameFichaTecnica 
         Caption         =   "Datos Generales Ficha Tecnica"
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
         Height          =   6855
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   11655
         Begin VB.ComboBox CboNumGra 
            BackColor       =   &H008080FF&
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
            ItemData        =   "FichaTecnica.frx":1167A
            Left            =   6960
            List            =   "FichaTecnica.frx":1168D
            TabIndex        =   26
            Text            =   "0"
            Top             =   5760
            Width           =   1695
         End
         Begin VB.TextBox TxtTipVen 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   5
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   1440
            Width           =   1695
         End
         Begin VB.TextBox TxtPesUniTap 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1560
            TabIndex        =   8
            Top             =   2520
            Width           =   1692
         End
         Begin VB.TextBox TxtDes3 
            Appearance      =   0  'Flat
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
            MaxLength       =   100
            TabIndex        =   25
            Top             =   6480
            Width           =   9975
         End
         Begin VB.TextBox TxtDes2 
            Appearance      =   0  'Flat
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
            MaxLength       =   100
            TabIndex        =   24
            Top             =   6120
            Width           =   9975
         End
         Begin VB.TextBox TxtTap 
            Appearance      =   0  'Flat
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
            Left            =   9840
            MaxLength       =   15
            TabIndex        =   20
            Top             =   4680
            Width           =   1695
         End
         Begin VB.TextBox TxtFon 
            Appearance      =   0  'Flat
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
            Left            =   6960
            MaxLength       =   15
            TabIndex        =   19
            Top             =   4680
            Width           =   1695
         End
         Begin VB.TextBox TxtEsp 
            Appearance      =   0  'Flat
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
            Left            =   4080
            MaxLength       =   15
            TabIndex        =   18
            Top             =   4680
            Width           =   1695
         End
         Begin VB.ComboBox CboTip 
            BackColor       =   &H0080FFFF&
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
            ItemData        =   "FichaTecnica.frx":116A0
            Left            =   1560
            List            =   "FichaTecnica.frx":116AA
            TabIndex        =   23
            Text            =   "PRODUCTO TERMINADO"
            Top             =   5760
            Width           =   2895
         End
         Begin VB.TextBox TxtUniCaj 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            TabIndex        =   11
            Top             =   3600
            Width           =   1695
         End
         Begin VB.TextBox TxtUniLam 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1560
            TabIndex        =   9
            Top             =   2880
            Width           =   1692
         End
         Begin VB.TextBox TxtPesUni 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1560
            TabIndex        =   7
            Top             =   2160
            Width           =   1692
         End
         Begin VB.TextBox TxtUniMed 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   12
            Top             =   3960
            Width           =   1695
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Activa"
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
            Left            =   3360
            TabIndex        =   14
            Top             =   4320
            Width           =   975
         End
         Begin VB.TextBox TxtMatEmp 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox TxtUsuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   9960
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox TxtGru 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   4
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.ComboBox CboOrigen 
            BackColor       =   &H008080FF&
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
            ItemData        =   "FichaTecnica.frx":116D1
            Left            =   1560
            List            =   "FichaTecnica.frx":116DB
            TabIndex        =   17
            Text            =   "INTERNO"
            Top             =   4680
            Width           =   1695
         End
         Begin VB.TextBox TxtFicTec 
            Appearance      =   0  'Flat
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
            MaxLength       =   15
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox TxtDes 
            Appearance      =   0  'Flat
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
            MaxLength       =   100
            TabIndex        =   3
            Top             =   720
            Width           =   9975
         End
         Begin VB.TextBox TxtEnv 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
            TabIndex        =   10
            Top             =   3240
            Width           =   1695
         End
         Begin VB.TextBox TxtNomCom 
            Appearance      =   0  'Flat
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
            MaxLength       =   12
            TabIndex        =   13
            Top             =   4320
            Width           =   1695
         End
         Begin VB.CheckBox chkImpCal 
            Caption         =   "Imprime Calidad"
            Height          =   192
            Left            =   6240
            TabIndex        =   15
            Top             =   4320
            Width           =   1575
         End
         Begin VB.CheckBox ChkImpDef 
            Caption         =   "Imprime Defectos"
            Height          =   192
            Left            =   8040
            TabIndex        =   16
            Top             =   4320
            Width           =   1575
         End
         Begin VB.TextBox TxtVar 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            MaxLength       =   10
            TabIndex        =   21
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   5040
            Width           =   1695
         End
         Begin VB.TextBox TxtAtr 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            MaxLength       =   10
            TabIndex        =   22
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   5400
            Width           =   1695
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Numero Grafica"
            Height          =   195
            Index           =   14
            Left            =   5400
            TabIndex        =   80
            Top             =   5760
            Width           =   1110
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Ventas"
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
            Left            =   240
            TabIndex        =   79
            Top             =   1440
            Width           =   1035
         End
         Begin VB.Label LblTipVen 
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
            TabIndex        =   78
            Top             =   1440
            Width           =   8175
         End
         Begin VB.Label lblLabels 
            Caption         =   "Peso x Unidad En Kilos Tapa"
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
            Height          =   435
            Index           =   12
            Left            =   240
            TabIndex        =   77
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Nomb.Cliente"
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
            Left            =   240
            TabIndex        =   76
            Top             =   6480
            Width           =   1140
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Desp. Larga"
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
            Left            =   240
            TabIndex        =   75
            Top             =   6120
            Width           =   1050
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Fondo"
            Height          =   195
            Index           =   9
            Left            =   5880
            TabIndex        =   74
            Top             =   4680
            Width           =   990
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Tapa"
            Height          =   195
            Index           =   8
            Left            =   8760
            TabIndex        =   73
            Top             =   4680
            Width           =   915
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Espesor"
            Height          =   195
            Index           =   7
            Left            =   3360
            TabIndex        =   72
            Top             =   4680
            Width           =   570
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Unid. x Caja"
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
            Index           =   6
            Left            =   240
            TabIndex        =   58
            Top             =   3600
            Width           =   1050
         End
         Begin VB.Label lblLabels 
            Caption         =   "Peso x Unidad En Kilos"
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
            Height          =   435
            Index           =   5
            Left            =   240
            TabIndex        =   57
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Unid. x Lamina"
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
            Index           =   4
            Left            =   240
            TabIndex        =   56
            Top             =   2880
            Width           =   1230
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Material Empaque"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   55
            Top             =   1800
            Width           =   1275
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Unidad Medida"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   54
            Top             =   3960
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   1
            Left            =   9360
            TabIndex        =   53
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Lbltip 
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
            TabIndex        =   51
            Top             =   1080
            Width           =   8175
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Fic.Tecn."
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
            TabIndex        =   50
            Top             =   1080
            Width           =   1260
         End
         Begin VB.Label LblVariable 
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
            TabIndex        =   48
            Top             =   5040
            Width           =   8175
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Origen"
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
            Left            =   240
            TabIndex        =   47
            Top             =   4680
            Width           =   570
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo De Producto"
            Height          =   195
            Left            =   240
            TabIndex        =   46
            Top             =   5880
            Width           =   1260
         End
         Begin VB.Label lblLabels 
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
            Index           =   18
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   975
         End
         Begin VB.Label lblLabels 
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
            Index           =   19
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Unid. x Tarima"
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
            Index           =   28
            Left            =   240
            TabIndex        =   43
            Top             =   3240
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Comercial"
            Height          =   195
            Index           =   29
            Left            =   240
            TabIndex        =   42
            Top             =   4320
            Width           =   1290
         End
         Begin VB.Label lblLabels 
            Caption         =   "Catalogo De Rutinas"
            Height          =   435
            Index           =   30
            Left            =   240
            TabIndex        =   41
            Top             =   4920
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Atributo P/Defectos"
            Height          =   435
            Index           =   31
            Left            =   240
            TabIndex        =   40
            Top             =   5400
            Width           =   1170
         End
         Begin VB.Label LblAtributos 
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
            TabIndex        =   39
            Top             =   5400
            Width           =   8175
         End
      End
   End
End
Attribute VB_Name = "FichaTecnica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VVariable As Boolean
Dim VAtributo As Boolean
Dim VTipo As Boolean
Dim VTipoVenta As Boolean

Dim mensaje As String
Dim buscar As String
Dim VCodigoViejo As String
Dim VCodigoNuevo As String
Dim VDescripcion As String
Dim VPeso As Single
Dim VUnidadesxLamina As Integer

Dim RVariable As New ADODB.Recordset
Dim RAtributo As New ADODB.Recordset
Dim RBuscaMateriasPrimas As New ADODB.Recordset
Dim RBuscaCatalogo As New ADODB.Recordset
Dim RBuscaTipo As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RFichaTecnica As New ADODB.Recordset

Dim VTexto As String
Dim BEditar As Boolean




Sub botones()
    If Bandera = True Then
         FrameFichaTecnica.Enabled = True
         
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtFicTec.SetFocus
         
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         DbGridFichaTecnica.Visible = False
         FrameOpciones.Visible = False
                  
    Else
         FrameFichaTecnica.Enabled = False
         
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True

         DbGridFichaTecnica.Visible = True
         FrameOpciones.Visible = True
         
    End If
End Sub

Private Sub CboOrigen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub CboTip_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub chkImpCal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Private Sub ChkImpDef_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub



Private Sub CmdActualizar_Click()
MousePointer = 11
            Set RFichaTecnica = New ADODB.Recordset
            Call Abrir_Recordset(RFichaTecnica, "Select * from FichaTecnica")
            Set DbGridFichaTecnica.DataSource = RFichaTecnica
MousePointer = 0
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
        TabFichasTecnicas.Tab = 0
        Bandera = True
        botones
        Limpia_Campos
        TxtFicTec.Enabled = True
        TxtFicTec.SetFocus
        CboOrigen.Text = "INTERNO"
        CboTip.Text = "PRODUCTO TERMINADO"
        TxtUsuario.Text = GUsuario
        CboNumGra.Text = "0"
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            
                mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RFichaTecnica.Delete
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147467259 Then
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                        
                        'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RFichaTecnica.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RFichaTecnica.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If

            
End Sub


Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RFichaTecnica.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RFichaTecnica.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RFichaTecnica.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RFichaTecnica.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RFichaTecnica.BOF Then
        RFichaTecnica.MoveFirst
    ElseIf RFichaTecnica.EOF Then
        RFichaTecnica.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
        TxtDatos.Text = ""
            Set RBuscaMateriasPrimas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where FMP.CodigoMateriaPrima = C.Esp_Tec And FMP.Esp_Tec = '" & TxtFicTec.Text & "'")
            Else
                    Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where UPPER(FMP.CodigoMateriaPrima) = UPPER(C.Esp_Tec) And UPPER(FMP.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            If RBuscaMateriasPrimas.RecordCount > 0 Then
                    Do Until RBuscaMateriasPrimas.EOF
                        TxtDatos.Text = TxtDatos.Text & Left(RBuscaMateriasPrimas(0) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(1) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(2) & Space(15), 15) & " " & RBuscaMateriasPrimas(3) & vbCrLf
                        RBuscaMateriasPrimas.MoveNext
                    Loop
            End If

    
MousePointer = 0


End Sub

Private Sub CmdBuscar_Click()
MousePointer = 11
    Set RFichaTecnica = New ADODB.Recordset
    If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RFichaTecnica, "Select * from FichaTecnica where Esp_Tec Like '" & TxtBuscar.Text & "%'")
            Else 'ORACLE
                Call Abrir_Recordset(RFichaTecnica, "Select * from FichaTecnica where UPPER(Esp_Tec) Like '" & UCase(TxtBuscar.Text) & "%'")
            End If
    ElseIf OptCodigoVariable.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RFichaTecnica, "Select * from FichaTecnica where Variables Like '" & TxtBuscar.Text & "%'")
            Else 'ORACLE
                Call Abrir_Recordset(RFichaTecnica, "Select * from FichaTecnica where UPPER(Variables) Like '" & UCase(TxtBuscar.Text) & "%'")
            End If
    ElseIf OptTipFicTec.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RFichaTecnica, "Select * from FichaTecnica where Tipo Like '" & TxtBuscar.Text & "%'")
            Else 'ORACLE
                Call Abrir_Recordset(RFichaTecnica, "Select * from FichaTecnica where UPPER(Tipo) Like '" & UCase(TxtBuscar.Text) & "%'")
            End If
    End If
            
            Set DbGridFichaTecnica.DataSource = RFichaTecnica
            TabFichasTecnicas.Tab = 1
MousePointer = 0

End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
        Bandera = False
        botones
        Llena_Campos
        TxtFicTec.Enabled = True
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
        TabFichasTecnicas.Tab = 0
        Bandera = True
        botones
        TxtFicTec.Enabled = False
        TxtDes.SetFocus
        TxtUsuario.Text = GUsuario
        BEditar = True
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   'VALIDA EL CODIGO
   If TxtFicTec.Text = "" Then
        MsgBox "Codigo Ficha Tecnica No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtFicTec.SetFocus
        Exit Sub
   End If
   
   'VALIDA LA DESCRIPCION
   If TxtDes.Text = "" Then
        MsgBox "Descripcion De Ficha Tecnica No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtDes.SetFocus
        Exit Sub
   End If
   
   
   If TxtMatEmp.Text = "" Then
        MsgBox "Material De Empaque No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtMatEmp.SetFocus
        Exit Sub
   End If
   
   If TxtUniMed.Text = "" Then
        MsgBox "Unidad De Medida No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtUniMed.SetFocus
        Exit Sub
   End If
   
   'VERIFICA EL ORIGEN DE LA FICHA TECNICA
   If CboOrigen.Text <> "INTERNO" And CboOrigen.Text <> "EXTERNO" Then
        MsgBox "Origen De Ficha Tecnica Incorrecto", vbOKOnly + vbInformation, "Informacion"
        CboOrigen.SetFocus
        Exit Sub
   End If
   
   'VERIFICA EL ORIGEN DE LA FICHA TECNICA
   If CboTip.Text <> "PRODUCTO TERMINADO" And CboTip.Text <> "MATERIA PRIMA" Then
        MsgBox "Tipo De inventario Incorrecto", vbOKOnly + vbInformation, "Informacion"
        CboTip.SetFocus
        Exit Sub
   End If
   
   'REVISA SI ES NUMERICO EL NUMERO DE ENVASES
   If Not IsNumeric(TxtEnv.Text) Then
        MsgBox "Unidades x Tarima Incorrecta", vbOKOnly + vbInformation, "Informacion"
        TxtEnv.SetFocus
        Exit Sub
   End If
   
   If Not IsNumeric(TxtPesUni.Text) Then
        MsgBox "Peso x Unidad Incorrecta", vbOKOnly + vbInformation, "Informacion"
        TxtPesUni.SetFocus
        Exit Sub
   End If
   
   If Not IsNumeric(TxtUniLam.Text) Then
        MsgBox "Unidades x Lamina Incorrecta", vbOKOnly + vbInformation, "Informacion"
        TxtUniLam.SetFocus
        Exit Sub
   End If
   
   If Not IsNumeric(TxtUniCaj.Text) Then
        MsgBox "Unidades x Caja Incorrecta", vbOKOnly + vbInformation, "Informacion"
        TxtUniCaj.SetFocus
        Exit Sub
   End If
   
   Set RBuscaTipo = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicatiposVentas Where CodigoTipo = '" & TxtGru.Text & "'")
        Else 'oracle
            Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicatiposVentas Where UPPER(CodigoTipo) = '" & UCase(TxtGru.Text) & "'")
        End If
        If RBuscaTipo.RecordCount > 0 Then
            
        Else
            MsgBox "Tipo Ficha Tecnica No Existe", vbOKOnly + vbInformation, "Informacion"
            TxtGru.SetFocus
            Exit Sub
        End If
   
       'TIPOS DE VENTAS
       Set RBuscaTipo = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicatiposVentas Where Codigo = '" & TxtTipVen.Text & "'")
        Else 'oracle
            Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicatiposVentas Where UPPER(Codigo) = '" & UCase(TxtTipVen.Text) & "'")
        End If
        If RBuscaTipo.RecordCount > 0 Then
            
        Else
            MsgBox "Tipo Venta No Existe", vbOKOnly + vbInformation, "Informacion"
            TxtTipVen.SetFocus
            Exit Sub
        End If
        
        'REVISA CATALOGO DE RUTINAS
        Set RVariable = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RVariable, "Select * From VariablesDescripcion Where CodigoVariable = '" & TxtVar.Text & "'")
            Else
                Call Abrir_Recordset(RVariable, "Select * From VariablesDescripcion Where UPPER(CodigoVariable) = '" & UCase(TxtVar.Text) & "'")
            End If
            If RVariable.RecordCount > 0 Then
                
            Else
                MsgBox "Catalogo De Rutinas No Existe", vbOKOnly + vbInformation, "Informacion"
                TxtVar.SetFocus
                Exit Sub
            End If
            
        'ATRIBUTOS
        Set RAtributo = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RAtributo, "Select * From Atributos Where Codigo = '" & TxtAtr.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RAtributo, "Select * From Atributos Where UPPER(Codigo) = '" & UCase(TxtAtr.Text) & "'")
            End If
            If RAtributo.RecordCount > 0 Then
                
            Else
                MsgBox "Atributos De Defectos No Existe", vbOKOnly + vbInformation, "Informacion"
                TxtAtr.SetFocus
                Exit Sub
            End If
   
   VCodigoNuevo = TxtFicTec.Text
   VDescripcion = TxtDes.Text
   VPeso = TxtPesUni.Text
   VUnidadesxLamina = TxtUniLam.Text
      
   'AGREGAR
                    If BEditar = False Then
                            VTexto = "'" & TxtFicTec.Text & "', '" 'FICHA TECNICA
                            VTexto = VTexto & TxtDes.Text & "', '" 'DESCRIPCION
                            VTexto = VTexto & TxtGru.Text & "', " 'TIPO FICHA TECNICA
                            VTexto = VTexto & TxtEnv.Text & ", '" 'ENVASES
                            VTexto = VTexto & TxtNomCom.Text & "', " 'NOMBRE COMERCIAL
                            If ChkImpDef.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'IMPRIMIE DEFECTOS
                            Else
                                VTexto = VTexto & "0" & ", " 'IMPRIME DEFECTOS
                            End If
                            If chkImpCal.Value = "1" Then
                                VTexto = VTexto & "-1" & ", '" 'IMPRIME CALIDAD
                            Else
                                VTexto = VTexto & "0" & ", '" 'IMPRIME CALIDAD
                            End If
                            VTexto = VTexto & TxtAtr.Text & "', '" 'ATRIBUTOS
                            VTexto = VTexto & TxtVar.Text & "', '" 'VARIABLES
                            VTexto = VTexto & CboOrigen.Text & "', '" 'ORIGEN
                            VTexto = VTexto & TxtUsuario.Text & "', '" 'USUARIO
                            VTexto = VTexto & TxtUniMed.Text & "', '" 'UNIDAD DE MEDIDA
                            VTexto = VTexto & TxtMatEmp.Text & "', " 'MATERIAL DE EMPAQUE
                            If Check1.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'ACTIVA
                            Else
                                VTexto = VTexto & "0" & ", " 'ACTIVA
                            End If
                            VTexto = VTexto & TxtPesUni.Text & ", " 'PESO X UNIDAD
                            VTexto = VTexto & TxtUniLam.Text & ", " 'UNIDADES X LAMINA
                            VTexto = VTexto & TxtUniCaj.Text & ", '" 'UNIDADES X CAJA
                            VTexto = VTexto & CboTip.Text & "', " 'TIPO DE INVENTARIO
                            VTexto = VTexto & TxtEsp.Text & ", '', '', '" 'CODIGO CLIENTE, PRODUCTO DEL CLIENTE
                            VTexto = VTexto & TxtFon.Text & "', '" ' FONDO
                            VTexto = VTexto & TxtTap.Text & "', '" ' TAPA
                            VTexto = VTexto & TxtDes2.Text & "', '" ' SEGUNDO NOMBRE
                            VTexto = VTexto & TxtDes3.Text & "', " ' TERCER NOMBRE
                            VTexto = VTexto & TxtPesUniTap.Text & ", '"  ' PESO X UNIDAD EN KILOS CON TAPA
                            VTexto = VTexto & TxtTipVen.Text & "', '" 'TIPO VENTA
                            VTexto = VTexto & CboNumGra.Text & "'" 'NUMERO GRAFICA
                            
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into FichaTecnica Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            
                            VTexto = "Descrip = '" & TxtDes.Text & "', " 'DESCRIPCION
                            VTexto = VTexto & "Tipo = '" & TxtGru.Text & "', " 'TIPO FICHA TECNICA
                            VTexto = VTexto & "Envases = " & TxtEnv.Text & ", " 'ENVASES
                            VTexto = VTexto & "Nombre_Comercial = '" & TxtNomCom.Text & "', " 'NOMBRE COMERCIAL
                            If ChkImpDef.Value = "1" Then
                                VTexto = VTexto & "Imp_Defe = -1, " 'IMPRIME DEFECTOS
                            Else
                                VTexto = VTexto & "Imp_Defe = 0, " 'IMPRIME DEFECTOS
                            End If
                            If chkImpCal.Value = "1" Then
                                VTexto = VTexto & "Imp_Cali = -1, " 'IMPRIME CALIDAD
                            Else
                                VTexto = VTexto & "Imp_Cali = 0, " 'IMPRIME CALIDAD
                            End If
                            VTexto = VTexto & "Atributos = '" & TxtAtr.Text & "', " 'ATRIBUTOS
                            VTexto = VTexto & "Variables = '" & TxtVar.Text & "', " 'VARIABLES
                            VTexto = VTexto & "Origen = '" & CboOrigen.Text & "', " 'ORIGEN
                            VTexto = VTexto & "Usuario = '" & TxtUsuario.Text & "', " 'USUARIO
                            VTexto = VTexto & "UnidadMedida = '" & TxtUniMed.Text & "', " 'UNIDAD DE MEDIDA
                            VTexto = VTexto & "MaterialEmpaque = '" & TxtMatEmp.Text & "', " 'MATERIAL DE EMPAQUE
                            If Check1.Value = "1" Then
                                VTexto = VTexto & "Activa = -1, " 'ACTIVA
                            Else
                                VTexto = VTexto & "Activa = 0, " 'ACTIVA
                            End If
                            VTexto = VTexto & "PesoxUnidad = " & TxtPesUni.Text & ", " 'PESO X UNIDAD
                            VTexto = VTexto & "UnidadesxLamina = " & TxtUniLam.Text & ", " 'UNIDADES X LAMINA
                            VTexto = VTexto & "UnidadesxCaja = " & TxtUniCaj.Text & ", " 'UNIDADES X CAJA
                            VTexto = VTexto & "TipoInventario = '" & CboTip.Text & "', " 'TIPO DE INVENTARIO
                            VTexto = VTexto & "Espesor = " & TxtEsp.Text & ", " 'TIPO DE INVENTARIO
                            VTexto = VTexto & "Fondo = '" & TxtFon.Text & "', " 'TIPO DE INVENTARIO
                            VTexto = VTexto & "Tapa = '" & TxtTap.Text & "', " 'TIPO DE INVENTARIO
                            VTexto = VTexto & "SegundoNombre = '" & TxtDes2.Text & "', " 'SEGUNDO NOMBRE
                            VTexto = VTexto & "TercerNombre = '" & TxtDes3.Text & "', " 'TERCER NOMBRE
                            VTexto = VTexto & "PesoxUnidadConTapa = " & TxtPesUniTap.Text & ", " 'PESO X UNIDAD EN KILOS Y TAPA"
                            VTexto = VTexto & "TipoVenta = '" & TxtTipVen.Text & "', " 'TIPO VENTA
                            VTexto = VTexto & "NumeroGrafica = '" & CboNumGra.Text & "'" 'NUMERO GRAFICA
                            
                            VTexto = VTexto & " Where Esp_Tec = '" & TxtFicTec.Text & "'" 'CODIGO
                            
                            Conexion.Execute "UPDATE FichaTecnica SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                      'SI ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo FichaTecnica Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtFicTec.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                       
        
        
      Bandera = False
      botones
      CmdAgregar.SetFocus
      
      TxtFicTec.Enabled = True
      
      RFichaTecnica.Requery
      RFichaTecnica.MoveLast
      Llena_Campos
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub




Private Sub DBGridBusqueda_DblClick()
    'PARA SELECCIONAR LA VARIABLE
    If VVariable = True Then
        TxtVar.Text = DbGridBusqueda.Columns(0)
        TxtVar.SetFocus
        FrameConsultas.Visible = False
    'PARA SELECCIONAR LA ATRIBUTO
    ElseIf VAtributo = True Then
        TxtAtr.Text = DbGridBusqueda.Columns(0)
        TxtAtr.SetFocus
        FrameConsultas.Visible = False
    'PARA SELECCIONAR EL TIPO DE FICHA TECNICA
    ElseIf VTipo = True Then
        TxtGru.Text = DbGridBusqueda.Columns(0)
        TxtGru.SetFocus
        FrameConsultas.Visible = False
    ElseIf VTipoVenta = True Then
        TxtTipVen.Text = DbGridBusqueda.Columns(0)
        TxtTipVen.SetFocus
        FrameConsultas.Visible = False
    End If
    
        
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)

If KeyAscii = 43 Then
    'PARA SELECCIONAR LA VARIABLE
    If VVariable = True Then
        TxtVar.Text = DbGridBusqueda.Columns(0)
        TxtVar.SetFocus
        FrameConsultas.Visible = False
    'PARA SELECCIONAR LA ATRIBUTO
    ElseIf VAtributo = True Then
        TxtAtr.Text = DbGridBusqueda.Columns(0)
        TxtAtr.SetFocus
        FrameConsultas.Visible = False
    'PARA SELECCIONAR EL TIPO DE FICHA TECNICA
    ElseIf VTipo = True Then
        TxtGru.Text = DbGridBusqueda.Columns(0)
        TxtGru.SetFocus
        FrameConsultas.Visible = False
    ElseIf VTipoVenta = True Then
        TxtTipVen.Text = DbGridBusqueda.Columns(0)
        TxtTipVen.SetFocus
        FrameConsultas.Visible = False
    End If
End If
End Sub

Private Sub DbGridFichaTecnica_HeadClick(ByVal ColIndex As Integer)
            RFichaTecnica.Sort = RFichaTecnica.Fields(ColIndex).Name
End Sub


Private Sub Form_Activate()
    TxtDatos.Text = ""
    Set RBuscaMateriasPrimas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where FMP.CodigoMateriaPrima = C.Esp_Tec And FMP.Esp_Tec = '" & TxtFicTec.Text & "'")
            Else
                    Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where UPPER(FMP.CodigoMateriaPrima) = UPPER(C.Esp_Tec) And UPPER(FMP.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
                    If RBuscaMateriasPrimas.RecordCount > 0 Then
                            Do Until RBuscaMateriasPrimas.EOF
                                TxtDatos.Text = TxtDatos.Text & Left(RBuscaMateriasPrimas(0) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(1) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(2) & Space(15), 15) & " " & RBuscaMateriasPrimas(3) & vbCrLf
                                RBuscaMateriasPrimas.MoveNext
                            Loop
                    End If
    
End Sub

Private Sub Form_Load()

        Set RFichaTecnica = New ADODB.Recordset
        Call Abrir_Recordset(RFichaTecnica, "Select * From FichaTecnica")
        Set DbGridFichaTecnica.DataSource = RFichaTecnica
        Llena_Campos

    
    'PARA HABILITAR EL GRID SOLO A USUARIOS AVANZADOS
    If GEditar = True Then
        DbGridFichaTecnica.AllowAddNew = True
        DbGridFichaTecnica.AllowUpdate = True
    End If
    
End Sub



Private Sub OptCodigo_Click()
    LblFicha.Caption = "Codigo Ficha Tecnica"
    TxtBuscar.SetFocus
End Sub

Private Sub OptCodigoVariable_Click()
    LblFicha.Caption = "Codigo Catalogo"
    TxtBuscar.SetFocus
End Sub


Private Sub OptTipFicTec_Click()
    LblFicha.Caption = "Tipo Ficha Tecnica"
    TxtBuscar.SetFocus
End Sub

Private Sub TabFichasTecnicas_Click(PreviousTab As Integer)

    If TabFichasTecnicas.Tab = 0 Then
        CmdBorrar.Enabled = True
            If CmdGrabar.Enabled = False Then
                Llena_Campos
            End If
    Else
        CmdBorrar.Enabled = False
    End If
End Sub



Private Sub TxtAtr_Change()
        Set RAtributo = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RAtributo, "Select * From Atributos Where Codigo = '" & TxtAtr.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RAtributo, "Select * From Atributos Where UPPER(Codigo) = '" & UCase(TxtAtr.Text) & "'")
            End If
            If RAtributo.RecordCount > 0 Then
                LblAtributos.Caption = "Men. " & RAtributo(1) & "  May. " & RAtributo(2) & "  Cri. " & RAtributo(3) & "  Mue. " & RAtributo(4)
            Else
                LblAtributos.Caption = ""
            End If
End Sub

Private Sub TxtAtr_DblClick()
            VVariable = False
            VAtributo = True
            VTipo = False
            VTipoVenta = False
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select * from Atributos")
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "4000"
            FrameConsultas.Visible = True
            TxtBusqueda.SetFocus
End Sub

Private Sub TxtAtr_GotFocus()
            TxtAtr.SelStart = 0
            TxtAtr.SelLength = Len(TxtAtr.Text)
End Sub

Private Sub TxtAtr_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
               SendKeys "{tab}"
            End If
                
                
            If KeyAscii = 43 Then
                VVariable = False
                VAtributo = True
                VTipo = False
                VTipoVenta = False
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select * from Atributos")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
                FrameConsultas.Visible = True
                TxtBusqueda.SetFocus
            End If

End Sub

Private Sub TxtAtr_Validate(Cancel As Boolean)
                Set RAtributo = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RAtributo, "Select * From Atributos Where Codigo = '" & TxtAtr.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RAtributo, "Select * From Atributos Where UPPER(Codigo) = '" & UCase(TxtAtr.Text) & "'")
                    End If
                    If RAtributo.RecordCount > 0 Then
                        LblAtributos.Caption = "Menores " & RAtributo!Menores & " Mayores " & RAtributo!Mayores & " Criticos " & RAtributo!Criticos & " Muestra " & RAtributo!Muestra
                    Else
                        LblAtributos.Caption = ""
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



Private Sub TxtBusqueda_Change()
        Set RBusqueda = New ADODB.Recordset
        
        If VVariable = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From VariablesDescripcion Where DescripcionVariable Like '%" & TxtBusqueda.Text & "%'")
                Else 'oracle
                    Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From VariablesDescripcion Where UPPER(DescripcionVariable) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From VariablesDescripcion Where CodigoVariable Like '%" & TxtBusqueda.Text & "%'")
                Else 'oracle
                    Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From VariablesDescripcion Where UPPER(CodigoVariable) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
        
        ElseIf VAtributo = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Menores, Mayores, Criticos From Atributos Where Codigo Like '%" & TxtBusqueda.Text & "%'")
                Else 'oracle
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Menores, Mayores, Criticos From Atributos Where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Menores, Mayores, Criticos From Atributos Where Codigo Like '%" & TxtBusqueda.Text & "%'")
                Else 'oracle
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Menores, Mayores, Criticos From Atributos Where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
        ElseIf VTipo = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicatipos Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                Else 'oracle
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicatipos Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicatipos Where CodigoTipo Like '%" & TxtBusqueda.Text & "%'")
                Else 'oracle
                    Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicatipos Where UPPER(CodigoTipo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
        ElseIf VTipoVenta = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From FichaTecnicatiposVentas Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                Else 'oracle
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From FichaTecnicatiposVentas Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From FichaTecnicatiposVentas Where Codigo Like '%" & TxtBusqueda.Text & "%'")
                Else 'oracle
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From FichaTecnicatiposVentas Where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
        
        End If

        Set DbGridBusqueda.DataSource = RBusqueda
        DbGridBusqueda.Columns(1).Width = "4000"
        
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

Private Sub TxtDes_GotFocus()
    TxtDes.SelStart = 0
    TxtDes.SelLength = Len(TxtDes.Text)
End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub


Private Sub TxtDes2_GotFocus()
        TxtDes2.SelStart = 0
        TxtDes2.SelLength = Len(TxtDes2.Text)
End Sub

Private Sub TxtDes2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtDes3_GotFocus()
        TxtDes3.SelStart = 0
        TxtDes3.SelLength = Len(TxtDes3.Text)
End Sub

Private Sub TxtDes3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
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


Private Sub TxtEsp_GotFocus()
        TxtEsp.SelStart = 0
        TxtEsp.SelLength = Len(TxtEsp.Text)
End Sub

Private Sub TxtEsp_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtFicTec_Change()
    TxtDatos.Text = ""
    Set RBuscaMateriasPrimas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where FMP.CodigoMateriaPrima = C.Esp_Tec And FMP.Esp_Tec = '" & TxtFicTec.Text & "'")
            Else
                    Call Abrir_Recordset(RBuscaMateriasPrimas, "Select FMP.Consumo, C.UnidadMedida, FMP.CodigoMateriaPrima, C.Descrip from FichaTecnicaConMateriaPrima FMP, FichaTecnica C where UPPER(FMP.CodigoMateriaPrima) = UPPER(C.Esp_Tec) And UPPER(FMP.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
    If RBuscaMateriasPrimas.RecordCount > 0 Then
            Do Until RBuscaMateriasPrimas.EOF
                TxtDatos.Text = TxtDatos.Text & Left(RBuscaMateriasPrimas(0) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(1) & Space(10), 10) & " " & Left(RBuscaMateriasPrimas(2) & Space(15), 15) & " " & RBuscaMateriasPrimas(3) & vbCrLf
                RBuscaMateriasPrimas.MoveNext
            Loop
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

Private Sub TxtFon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TxtGru_Change()
    Set RBuscaTipo = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicatipos Where CodigoTipo = '" & TxtGru.Text & "'")
        Else 'oracle
            Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicatipos Where UPPER(CodigoTipo) = '" & UCase(TxtGru.Text) & "'")
        End If
        If RBuscaTipo.RecordCount > 0 Then
            Lbltip.Caption = RBuscaTipo!Descripcion
        Else
            Lbltip.Caption = ""
        End If
End Sub

Private Sub TxtGru_DblClick()
        VVariable = False
        VAtributo = False
        VTipo = True
        VTipoVenta = False
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select * from FichaTecnicaTipos")
        Set DbGridBusqueda.DataSource = RBusqueda
        DbGridBusqueda.Columns(1).Width = "4000"
        FrameConsultas.Visible = True
        TxtBusqueda.SetFocus
End Sub

Private Sub TxtGru_GotFocus()
    TxtGru.SelStart = 0
    TxtGru.SelLength = Len(TxtGru.Text)
End Sub

Private Sub TxtGru_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
        
    If KeyAscii = 43 Then
            VVariable = False
            VAtributo = False
            VTipo = True
            VTipoVenta = False
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select * from FichaTecnicaTipos")
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "4000"
            FrameConsultas.Visible = True
            TxtBusqueda.SetFocus
    End If

    
End Sub

Private Sub TxtMatEmp_GotFocus()
        TxtMatEmp.SelStart = 0
        TxtMatEmp.SelLength = Len(TxtMatEmp.Text)
End Sub

Private Sub TxtMatEmp_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtNomCom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TxtPesUni_GotFocus()
        TxtPesUni.SelStart = 0
        TxtPesUni.SelLength = Len(TxtPesUni.Text)
End Sub

Private Sub TxtPesUni_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub


Private Sub TxtPesUniTap_GotFocus()
    TxtPesUniTap.SelStart = 0
    TxtPesUniTap.SelLength = Len(TxtPesUniTap.Text)
End Sub

Private Sub TxtPesUniTap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtTap_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TxtTipVen_Change()
    Set RBuscaTipo = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicatiposVentas Where Codigo = '" & TxtTipVen.Text & "'")
        Else 'oracle
            Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicatiposVentas Where UPPER(Codigo) = '" & UCase(TxtTipVen.Text) & "'")
        End If
        If RBuscaTipo.RecordCount > 0 Then
            LblTipVen.Caption = RBuscaTipo!Descripcion
        Else
            LblTipVen.Caption = ""
        End If
End Sub

Private Sub TxtTipVen_DblClick()
        VVariable = False
        VAtributo = False
        VTipo = False
        VTipoVenta = True
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select * from FichaTecnicaTiposVentas")
        Set DbGridBusqueda.DataSource = RBusqueda
        DbGridBusqueda.Columns(1).Width = "4000"
        FrameConsultas.Visible = True
        TxtBusqueda.SetFocus
End Sub

Private Sub TxtTipVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
        
    If KeyAscii = 43 Then
            VVariable = False
            VAtributo = False
            VTipo = True
            VTipoVenta = False
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select * from FichaTecnicaTiposVentas")
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "4000"
            FrameConsultas.Visible = True
            TxtBusqueda.SetFocus
    End If
End Sub

Private Sub TxtUniCaj_GotFocus()
        TxtUniCaj.SelStart = 0
        TxtUniCaj.SelLength = Len(TxtUniCaj.Text)
End Sub

Private Sub TxtUniCaj_KeyPress(KeyAscii As Integer)
         If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtUniLam_GotFocus()
        TxtUniLam.SelStart = 0
        TxtUniLam.SelLength = Len(TxtUniLam.Text)
End Sub

Private Sub TxtUniLam_KeyPress(KeyAscii As Integer)
         If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtUniMed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TxtVar_Change()
        Set RVariable = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RVariable, "Select * From VariablesDescripcion Where CodigoVariable = '" & TxtVar.Text & "'")
            Else
                Call Abrir_Recordset(RVariable, "Select * From VariablesDescripcion Where UPPER(CodigoVariable) = '" & UCase(TxtVar.Text) & "'")
            End If
        If RVariable.RecordCount > 0 Then
            LblVariable.Caption = RVariable(1)
        Else
            LblVariable.Caption = ""
        End If

End Sub

Private Sub TxtVar_DblClick()
        VVariable = True
        VAtributo = False
        VTipo = False
        VTipoVenta = False
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select * from VariablesDescripcion")
        Set DbGridBusqueda.DataSource = RBusqueda
        DbGridBusqueda.Columns(1).Width = "4000"
        FrameConsultas.Visible = True
        TxtBusqueda.SetFocus
End Sub

Private Sub TxtVar_GotFocus()
        TxtVar.SelStart = 0
        TxtVar.SelLength = Len(TxtVar.Text)
End Sub

Private Sub TxtVar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

        If KeyAscii = 43 Then
            VVariable = True
            VAtributo = False
            VTipo = False
            VTipoVenta = False
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select * from VariablesDescripcion")
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "4000"
            FrameConsultas.Visible = True
            TxtBusqueda.SetFocus
        End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
    If RFichaTecnica.RecordCount > 0 Then
                        If IsNull(RFichaTecnica!Esp_Tec) Then
                            TxtFicTec.Text = ""
                        Else
                            TxtFicTec.Text = RFichaTecnica!Esp_Tec
                        End If
                        
                        If IsNull(RFichaTecnica!Descrip) Then
                            TxtDes.Text = ""
                        Else
                            TxtDes.Text = RFichaTecnica!Descrip
                        End If
                        
                        If IsNull(RFichaTecnica!Tipo) Then
                            TxtGru.Text = ""
                        Else
                            TxtGru.Text = RFichaTecnica!Tipo
                        End If
                        
                        If IsNull(RFichaTecnica!Envases) Then
                            TxtEnv.Text = ""
                        Else
                            TxtEnv.Text = RFichaTecnica!Envases
                        End If
                        
                        If IsNull(RFichaTecnica!Nombre_Comercial) Then
                            TxtNomCom.Text = ""
                        Else
                            TxtNomCom.Text = RFichaTecnica!Nombre_Comercial
                        End If
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                                If RFichaTecnica!Imp_Defe = "Verdadero" Then
                                    ChkImpDef.Value = "1"
                                Else
                                    ChkImpDef.Value = "0"
                                End If
                        Else
                                If RFichaTecnica!Imp_Defe = "-1" Then
                                    ChkImpDef.Value = "1"
                                Else
                                    ChkImpDef.Value = "0"
                                End If
                        End If
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                                If RFichaTecnica!Imp_Cali = "Verdadero" Then
                                    chkImpCal.Value = "1"
                                Else
                                    chkImpCal.Value = "0"
                                End If
                        Else
                                If RFichaTecnica!Imp_Cali = "-1" Then
                                    chkImpCal.Value = "1"
                                Else
                                    chkImpCal.Value = "0"
                                End If
                        End If
                        
                        If IsNull(RFichaTecnica!Atributos) Then
                            TxtAtr.Text = ""
                        Else
                            TxtAtr.Text = RFichaTecnica!Atributos
                        End If
                        
                        If IsNull(RFichaTecnica!Variables) Then
                            TxtVar.Text = ""
                        Else
                            TxtVar.Text = RFichaTecnica!Variables
                        End If
                        
                        If IsNull(RFichaTecnica!Origen) Then
                            CboOrigen.Text = ""
                        Else
                            CboOrigen.Text = RFichaTecnica!Origen
                        End If
                        
                        If IsNull(RFichaTecnica!Usuario) Then
                            TxtUsuario.Text = ""
                        Else
                            TxtUsuario.Text = RFichaTecnica!Usuario
                        End If
                        
                        If IsNull(RFichaTecnica!unidadMedida) Then
                            TxtUniMed.Text = ""
                        Else
                            TxtUniMed.Text = RFichaTecnica!unidadMedida
                        End If
                        
                        If IsNull(RFichaTecnica!MaterialEmpaque) Then
                            TxtMatEmp.Text = ""
                        Else
                            TxtMatEmp.Text = RFichaTecnica!MaterialEmpaque
                        End If
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                                If RFichaTecnica!Activa = "Verdadero" Then
                                    Check1.Value = "1"
                                Else
                                    Check1.Value = "0"
                                End If
                        Else 'ORACLE
                                If RFichaTecnica!Activa = "-1" Then
                                    Check1.Value = "1"
                                Else
                                    Check1.Value = "0"
                                End If
                        End If
                                
                        If IsNull(RFichaTecnica!PesoxUnidad) Then
                            TxtPesUni.Text = ""
                        Else
                            TxtPesUni.Text = RFichaTecnica!PesoxUnidad
                        End If
                        
                        If IsNull(RFichaTecnica!UnidadesxLamina) Then
                            TxtUniLam.Text = ""
                        Else
                            TxtUniLam.Text = RFichaTecnica!UnidadesxLamina
                        End If
                        
                        If IsNull(RFichaTecnica!UnidadesxCaja) Then
                            TxtUniCaj.Text = ""
                        Else
                            TxtUniCaj.Text = RFichaTecnica!UnidadesxCaja
                        End If
                        
                        If IsNull(RFichaTecnica!TipoInventario) Then
                            CboTip.Text = ""
                        Else
                            CboTip.Text = RFichaTecnica!TipoInventario
                        End If
                        
                        If IsNull(RFichaTecnica!espesor) Then
                            TxtEsp.Text = ""
                        Else
                            TxtEsp.Text = RFichaTecnica!espesor
                        End If
                        If IsNull(RFichaTecnica!Fondo) Then
                            TxtFon.Text = ""
                        Else
                            TxtFon.Text = RFichaTecnica!Fondo
                        End If
                        If IsNull(RFichaTecnica!Tapa) Then
                            TxtTap.Text = ""
                        Else
                            TxtTap.Text = RFichaTecnica!Tapa
                        End If
                        If IsNull(RFichaTecnica!SegundoNombre) Then
                            TxtDes2.Text = ""
                        Else
                            TxtDes2.Text = RFichaTecnica!SegundoNombre
                        End If
                        If IsNull(RFichaTecnica!TercerNombre) Then
                            TxtDes3.Text = ""
                        Else
                            TxtDes3.Text = RFichaTecnica!TercerNombre
                        End If
                        If IsNull(RFichaTecnica!PesoxUnidadConTapa) Then
                            TxtPesUniTap.Text = ""
                        Else
                            TxtPesUniTap.Text = RFichaTecnica!PesoxUnidadConTapa
                        End If
                        If IsNull(RFichaTecnica!TipoVenta) Then
                            TxtTipVen.Text = ""
                        Else
                            TxtTipVen.Text = RFichaTecnica!TipoVenta
                        End If
                        If IsNull(RFichaTecnica!NumeroGrafica) Then
                            CboNumGra.Text = ""
                        Else
                            CboNumGra.Text = RFichaTecnica!NumeroGrafica
                        End If

    Else
        Limpia_Campos
    End If
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        
            TxtFicTec.Text = ""
            TxtDes.Text = ""
            TxtGru.Text = ""
            TxtEnv.Text = "0"
            TxtNomCom.Text = ""
            ChkImpDef.Value = "0"
            chkImpCal.Value = "0"
            TxtAtr.Text = ""
            TxtVar.Text = ""
            CboOrigen.Text = ""
            TxtUsuario.Text = ""
            TxtUniMed.Text = ""
            TxtMatEmp.Text = ""
            Check1.Value = "0"
            TxtPesUni.Text = "0"
            TxtUniLam.Text = "0"
            TxtUniCaj.Text = "0"
            CboTip.Text = ""
            TxtEsp.Text = 0
            TxtFon.Text = ""
            TxtTap.Text = ""
            TxtDes2.Text = ""
            TxtDes3.Text = ""
            TxtPesUniTap.Text = "0"
            TxtTipVen.Text = ""
        
End Sub

