VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CerrarBulto 
   BackColor       =   &H00008000&
   Caption         =   "Cerrar Bulto o Tarima"
   ClientHeight    =   8475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9495
   Icon            =   "CerrarBulto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   9495
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8415
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
      Begin MSDataGridLib.DataGrid DbGridConsultas 
         Height          =   7215
         Left            =   120
         TabIndex        =   45
         ToolTipText     =   "click en encabezado de columna para indexar"
         Top             =   1080
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   12726
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
         Left            =   8640
         Picture         =   "CerrarBulto.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   44
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   42
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   43
         Top             =   360
         Width           =   975
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Index           =   1
      Left            =   120
      MouseIcon       =   "CerrarBulto.frx":24B4
      Picture         =   "CerrarBulto.frx":28F6
      Style           =   1  'Graphical
      TabIndex        =   76
      ToolTipText     =   "Primer Registro"
      Top             =   7800
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Index           =   2
      Left            =   480
      MouseIcon       =   "CerrarBulto.frx":2E28
      Picture         =   "CerrarBulto.frx":326A
      Style           =   1  'Graphical
      TabIndex        =   75
      ToolTipText     =   "Registro Anterior"
      Top             =   7800
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Index           =   3
      Left            =   8640
      MouseIcon       =   "CerrarBulto.frx":379C
      Picture         =   "CerrarBulto.frx":3BDE
      Style           =   1  'Graphical
      TabIndex        =   74
      ToolTipText     =   "Siguiente Registro"
      Top             =   7800
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   555
      Index           =   4
      Left            =   9000
      MouseIcon       =   "CerrarBulto.frx":4110
      Picture         =   "CerrarBulto.frx":4552
      Style           =   1  'Graphical
      TabIndex        =   73
      ToolTipText     =   "Ultimo Registro"
      Top             =   7800
      Width           =   375
   End
   Begin TabDlg.SSTab TabIngresos 
      Height          =   7695
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   32768
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CerrarBulto.frx":4A84
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameNumerosIngresos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CerrarBulto.frx":4D9E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridCierreBulto"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CerrarBulto.frx":51F0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label2(8)"
      Tab(2).Control(1)=   "Label2(9)"
      Tab(2).Control(2)=   "LblBusCodMatPri"
      Tab(2).Control(3)=   "DtpFecIni"
      Tab(2).Control(4)=   "DtpFecFin"
      Tab(2).Control(5)=   "CmdBuscar"
      Tab(2).Control(6)=   "CmdActualizar"
      Tab(2).Control(7)=   "Txtbuscar"
      Tab(2).ControlCount=   8
      Begin MSDataGridLib.DataGrid DbGridCierreBulto 
         Height          =   6855
         Left            =   -74880
         TabIndex        =   69
         Top             =   720
         Width           =   9255
         _ExtentX        =   16325
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
         ColumnCount     =   21
         BeginProperty Column00 
            DataField       =   "Fecha"
            Caption         =   "Fecha"
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
            DataField       =   "Turno"
            Caption         =   "Turno"
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
            DataField       =   "Hora"
            Caption         =   "Hora"
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
            DataField       =   "Linea"
            Caption         =   "Linea"
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
            DataField       =   "FechaProduccion"
            Caption         =   "Fecha Produccion"
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
            DataField       =   "LineaProduccion"
            Caption         =   "Linea Produccion"
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
            DataField       =   "FichaTecnica"
            Caption         =   "Ficha Tecnica"
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
            DataField       =   "Tarima"
            Caption         =   "Tarima"
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
            DataField       =   "BodegaSalida"
            Caption         =   "Bodega Salida"
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
            DataField       =   "Existencia"
            Caption         =   "Existencia"
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
            DataField       =   "CantidadMas"
            Caption         =   "De +"
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
            DataField       =   "CantidadMenos"
            Caption         =   "De -"
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
            DataField       =   "ContadorInicial"
            Caption         =   "Contador Inicial"
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
            DataField       =   "ContadorFinal"
            Caption         =   "Contador Final"
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
            DataField       =   "CantidadProcesada"
            Caption         =   "Cantidad Procesada"
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
            DataField       =   "DesperdicioProceso"
            Caption         =   "Desperdicio Proceso"
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
            DataField       =   "DesperdicioProveedor"
            Caption         =   "Desperdicio Proveedor"
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
            DataField       =   "CantidadProcesadaReal"
            Caption         =   "Cantidad Procesada"
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
            DataField       =   "Total"
            Caption         =   "Total"
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
            DataField       =   "UsuarioAgregar"
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
            DataField       =   "Observaciones"
            Caption         =   "Observaciones"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1019.906
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1230.236
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column15 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column16 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column17 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column18 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column19 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   1844.787
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Txtbuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67680
         TabIndex        =   28
         Top             =   5040
         Width           =   1935
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Actualizar"
         Height          =   735
         Left            =   -67680
         Picture         =   "CerrarBulto.frx":5642
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   6360
         Width           =   1935
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Datos"
         Height          =   735
         Left            =   -67680
         Picture         =   "CerrarBulto.frx":594C
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   5520
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -67680
         TabIndex        =   27
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   53805059
         CurrentDate     =   37301
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -67680
         TabIndex        =   26
         Top             =   4200
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   53805059
         CurrentDate     =   37301
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
         Height          =   6855
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   9195
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            MaxLength       =   100
            TabIndex        =   17
            Top             =   6480
            Width           =   6735
         End
         Begin VB.TextBox TxtLinPro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   5
            Top             =   1440
            Width           =   1815
         End
         Begin MSMask.MaskEdBox MskCanMen 
            Height          =   285
            Left            =   2280
            TabIndex        =   16
            Top             =   6000
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16744576
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanMas 
            Height          =   285
            Left            =   2280
            TabIndex        =   15
            Top             =   5640
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16744576
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskNueExi 
            Height          =   525
            Left            =   6240
            TabIndex        =   19
            Top             =   4560
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   926
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            ForeColor       =   49152
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   3
            Top             =   600
            Width           =   1815
         End
         Begin MSMask.MaskEdBox MskCanProRea 
            Height          =   285
            Left            =   2280
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   5160
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTot 
            Height          =   525
            Left            =   6240
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   5160
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   926
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   255
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskDesProv 
            Height          =   285
            Left            =   2280
            TabIndex        =   13
            Top             =   4680
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskExi 
            Height          =   525
            Left            =   6240
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   3000
            Width           =   2775
            _ExtentX        =   4895
            _ExtentY        =   926
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   32768
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskDesPro 
            Height          =   285
            Left            =   2280
            TabIndex        =   12
            Top             =   4320
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanPro 
            Height          =   285
            Left            =   2280
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   3840
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskConFin 
            Height          =   285
            Left            =   2280
            TabIndex        =   10
            Top             =   3480
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBodSal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2280
            Locked          =   -1  'True
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   2640
            Width           =   1815
         End
         Begin VB.TextBox TxtTur 
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
            Left            =   4800
            MaxLength       =   1
            TabIndex        =   1
            Top             =   240
            Width           =   495
         End
         Begin MSMask.MaskEdBox MskConIni 
            Height          =   285
            Left            =   2280
            TabIndex        =   9
            Top             =   3120
            Width           =   1800
            _ExtentX        =   3175
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8438015
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFec 
            Height          =   285
            Left            =   2280
            TabIndex        =   0
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtUsu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   7800
            MaxLength       =   10
            TabIndex        =   39
            Top             =   240
            Width           =   1215
         End
         Begin MSMask.MaskEdBox MskHor 
            Height          =   285
            Left            =   5880
            TabIndex        =   2
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "hh:mm"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtFicTec 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            Left            =   2280
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox TxtTar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
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
            Left            =   2280
            TabIndex        =   7
            Top             =   2160
            Width           =   1815
         End
         Begin MSMask.MaskEdBox TxtFecPro 
            Height          =   285
            Left            =   2280
            TabIndex        =   4
            Top             =   1080
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Datos Del Bulto/Tarima"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   78
            Top             =   840
            Width           =   2010
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   77
            Top             =   6480
            Width           =   1275
         End
         Begin VB.Label LblLinPro 
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
            Left            =   4320
            TabIndex        =   72
            Top             =   1440
            Width           =   4695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            Caption         =   "Linea Entrada/Produc."
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
            Index           =   23
            Left            =   240
            TabIndex        =   71
            Top             =   1440
            Width           =   1950
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            Caption         =   "Fecha Entrada/Produc."
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
            Index           =   22
            Left            =   240
            TabIndex        =   70
            Top             =   1080
            Width           =   2010
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Cantidad de Cierres De Bulto"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   195
            Index           =   21
            Left            =   4440
            TabIndex        =   68
            Top             =   6000
            Width           =   2475
         End
         Begin VB.Label LblCanCieBul 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   6960
            TabIndex        =   67
            Top             =   5880
            Width           =   2055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   15
            Left            =   7080
            TabIndex        =   66
            Top             =   240
            Width           =   660
         End
         Begin VB.Label LblUnidadMedida 
            Alignment       =   1  'Right Justify
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
            Height          =   285
            Left            =   6240
            TabIndex        =   65
            Top             =   3600
            Width           =   2775
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
            Left            =   4320
            TabIndex        =   64
            Top             =   1800
            Width           =   4695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad de Menos"
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
            Left            =   240
            TabIndex        =   63
            Top             =   6000
            Width           =   1650
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad de Mas"
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
            TabIndex        =   62
            Top             =   5640
            Width           =   1440
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Existencia Nueva"
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
            Index           =   12
            Left            =   4440
            TabIndex        =   61
            Top             =   4680
            Width           =   1500
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
            Left            =   4200
            TabIndex        =   60
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Linea En Que Se Utilizo"
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
            Index           =   11
            Left            =   120
            TabIndex        =   59
            Top             =   600
            Width           =   2040
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Real Procesado"
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
            TabIndex        =   58
            Top             =   5160
            Width           =   1365
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000B&
            Caption         =   "Total a Descargar"
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
            Left            =   4440
            TabIndex        =   57
            Top             =   5280
            Width           =   1545
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Desperdicio Proveedor"
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
            Left            =   240
            TabIndex        =   56
            Top             =   4680
            Width           =   1950
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Existencia Actual"
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
            Left            =   4440
            TabIndex        =   55
            Top             =   3120
            Width           =   1485
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Desperdicio Proceso"
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
            TabIndex        =   54
            Top             =   4320
            Width           =   1770
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
            Left            =   4320
            TabIndex        =   53
            Top             =   2640
            Width           =   4695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cantidad Procesada"
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
            Index           =   20
            Left            =   240
            TabIndex        =   52
            Top             =   3840
            Width           =   1725
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Contador Final"
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
            Index           =   19
            Left            =   240
            TabIndex        =   51
            Top             =   3480
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bodega Actual"
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
            Index           =   18
            Left            =   120
            TabIndex        =   50
            Top             =   2640
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   17
            Left            =   4200
            TabIndex        =   49
            Top             =   240
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
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
            Left            =   240
            TabIndex        =   48
            Top             =   3120
            Width           =   1350
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
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
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   35
            Top             =   1800
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            Caption         =   "Bulto/Tarima"
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
            Index           =   1
            Left            =   240
            TabIndex        =   34
            Top             =   2160
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   0
            Left            =   5400
            TabIndex        =   33
            Top             =   240
            Width           =   420
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00C0C0C0&
            BackStyle       =   1  'Opaque
            Height          =   3375
            Left            =   120
            Top             =   3000
            Width           =   4095
         End
         Begin VB.Shape Shape2 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            Height          =   1575
            Left            =   120
            Top             =   960
            Width           =   4095
         End
      End
      Begin VB.Label LblBusCodMatPri 
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
         Left            =   -69000
         TabIndex        =   40
         Top             =   5040
         Width           =   1230
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
         TabIndex        =   38
         Top             =   4200
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
         TabIndex        =   37
         Top             =   4560
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   555
      Left            =   7200
      MouseIcon       =   "CerrarBulto.frx":5D8E
      Picture         =   "CerrarBulto.frx":61D0
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7800
      Width           =   1380
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   555
      Left            =   5640
      MouseIcon       =   "CerrarBulto.frx":8242
      Picture         =   "CerrarBulto.frx":8684
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "al borrar se regresa la existencia de nuevo"
      Top             =   7800
      Width           =   1500
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   4080
      MouseIcon       =   "CerrarBulto.frx":8BB6
      Picture         =   "CerrarBulto.frx":8FF8
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7800
      Width           =   1500
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   2520
      MouseIcon       =   "CerrarBulto.frx":952A
      Picture         =   "CerrarBulto.frx":996C
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7800
      Width           =   1500
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   555
      Left            =   960
      MouseIcon       =   "CerrarBulto.frx":9E9E
      Picture         =   "CerrarBulto.frx":A2E0
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7800
      Width           =   1500
   End
End
Attribute VB_Name = "CerrarBulto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim VTexto As String

Dim RBuscaTarima As New ADODB.Recordset
Dim RBuscaNumeroIngresoEntradas As New ADODB.Recordset
Dim RBuscaUnidad As New ADODB.Recordset
Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaTotalProcesado As New ADODB.Recordset
Dim RBuscaExistencia As New ADODB.Recordset
Dim RBuscaContadorInicial As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaLineaP As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim RCuentaBultos As New ADODB.Recordset
Dim RBuscaLiberado As New ADODB.Recordset
Dim RBuscaCuerposPorLamina As New ADODB.Recordset
Dim RBuscaEntradasMateriaPrima As New ADODB.Recordset
Dim RCierreBulto As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VFichaTecnica As String
Dim VTarima As Integer
Dim VFechaProduccion As Date
Dim VLinea As String

Dim VUltimaFicha As String
Dim VUltimaTarima As Integer
Dim VUltimaLinea As String

Dim VBodega As String
Dim VCantidadMateriaPrima As Currency
Dim VCantidadSalida As Currency
Dim VNuevaExistencia As Currency
Dim VTotalaDescargar As Currency

Dim BNumeroIngreso As Boolean
Dim BFichaTecnica As Boolean
Dim BLinea As Boolean


Sub botones()
    If Bandera = True Then
         FrameNumerosIngresos.Enabled = True
         CmdAgregar.Enabled = False
         
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtFicTec.SetFocus

         TxtBuscar.Visible = False
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         
         DbGridCierreBulto.Visible = False
    Else
         FrameNumerosIngresos.Enabled = False
         CmdAgregar.Enabled = True
         
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         
         TxtBuscar.Visible = True
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True

         
         DbGridCierreBulto.Visible = True
    End If
End Sub

Private Sub CmdActualizar_Click()
        Set RCierreBulto = New ADODB.Recordset
        Call Abrir_Recordset(RCierreBulto, "Select * From CierreBulto")
        
        Set DbGridCierreBulto.DataSource = RCierreBulto
        TabIngresos.Tab = 1
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        'BOTONES DISPONIBLES
        Bandera = True
        botones
        Limpia_Campos
     
        'ASIGNA FECHA Y HORA
        MskFec.Text = Date
        MskHor.Text = Format(Time, "hh:mm")
        
        TxtFicTec.Text = VUltimaFicha
        TxtTar.Text = VUltimaTarima + 1
        TxtLinPro.Text = VUltimaLinea
        
        'ASIGNA USUARIO AGREGAR
        TxtUsu.Text = GUsuario
        MskConIni.Text = "0"
        MskConFin.Text = "0"
        MskCanPro.Text = "0"
        MskDesPro.Text = "0"
        MskDesProv.Text = "0"
        MskCanMas.Text = "0"
        MskCanMen.Text = "0"
        MskFec.SetFocus
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

If GBorrar = False Then
       MsgBox "Usted No Tiene Acceso a Esta Funcion, Consulte Encargado De Amapro", vbOKOnly + vbInformation, "Informacion"
       Exit Sub
End If
            VFechaProduccion = Format(TxtFecPro.Text, "dd/mm/yyyy")
            VLinea = UCase(TxtLinPro.Text)
            VFichaTecnica = UCase(TxtFicTec.Text)
            VTarima = TxtTar.Text
            
            VBodega = UCase(TxtBodSal.Text)
            VCantidadMateriaPrima = MskExi.Text
            VCantidadSalida = MskTot.Text
            

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                        'REGRESA EL SALDO
                        Conexion.BeginTrans
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Update DetalleEntradasInventario Set Saldo = " & VCantidadMateriaPrima & " Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'"
                            Else 'ORACLE
                                Conexion.Execute "Update DetalleEntradasInventario Set Saldo = " & VCantidadMateriaPrima & " Where FechaProduccion = To_Date('" & VFechaProduccion & "', 'dd/mm/yyyy')" & " And Tarima = " & VTarima & " And UPPER(Linea) = '" & UCase(VLinea) & "' And UPPER(FichaTecnica) = '" & UCase(VFichaTecnica) & "'"
                            End If
                            
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Exit Sub
                            End If
                        
            
                        'BORRA EL REGISTRO
                        RCierreBulto.Delete
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147467259 Then
                                Conexion.RollbackTrans
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147467259 And Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                        
                        'TERMINA LA CONEXION SI NO HYA ERRORRES
                        Conexion.CommitTrans
                        
                        'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RCierreBulto.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RCierreBulto.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                        End If
                        
                        Llena_Campos
                    End If
                
                
                
End Sub


Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RCierreBulto.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RCierreBulto.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RCierreBulto.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RCierreBulto.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RCierreBulto.BOF Then
        RCierreBulto.MoveFirst
    ElseIf RCierreBulto.EOF Then
        RCierreBulto.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click()
        Set RCierreBulto = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RCierreBulto, "Select * From CierreBulto Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And FichaTecnica Like '" & TxtBuscar.Text & "%'")
            Else 'ORACLE
                Call Abrir_Recordset(RCierreBulto, "Select * From CierreBulto Where Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(FichaTecnica) Like '" & UCase(TxtBuscar.Text) & "%'")
            End If
            
            Set DbGridCierreBulto.DataSource = RCierreBulto
            TabIngresos.Tab = 1
        
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
        
        
        Bandera = False
        botones
        Llena_Campos
End Sub



Private Sub CmdGrabar_Click()
   On Error Resume Next
   
    MskFec.Text = Format(MskFec.Text, "dd/mm/yyyy")
    
    VBodega = TxtBodSal.Text
    VNuevaExistencia = MskNueExi.Text
    VCantidadMateriaPrima = MskTot.Text
    
    VFichaTecnica = UCase(TxtFicTec.Text)
    VTarima = TxtTar.Text
    VLinea = UCase(TxtLinPro.Text)
    TxtFecPro.Text = Format(TxtFecPro.Text, "dd/mm/yyyy")
    VFechaProduccion = TxtFecPro.Text
    
    VUltimaFicha = UCase(TxtFicTec.Text)
    VUltimaTarima = TxtTar.Text
    VUltimaLinea = UCase(TxtLinPro.Text)
            
    Set RBuscaFichaTecnica = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select TipoInventario From FichaTecnica where Esp_Tec = '" & TxtFicTec.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select TipoInventario From FichaTecnica where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            
            If RBuscaFichaTecnica.RecordCount > 0 Then
                If RBuscaFichaTecnica!TipoInventario = "MATERIA PRIMA" Then
                    'REVISA SI LA BODEGA ES DE NO CONFORME
                    Set RBuscaBodega = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaBodega, "Select EsBodegaDeNoConforme From BodegasInventario Where CodigoBodega = '" & TxtBodSal.Text & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaBodega, "Select EsBodegaDeNoConforme From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBodSal.Text) & "'")
                        End If
                        If RBuscaBodega.RecordCount > 0 Then
                            If RBuscaBodega!EsBodegadeNoConforme = -1 Then
                                    MsgBox "No Se Puede Grabar, Porque Este Bulto Esta En Una Bodega De No Conforme", vbOKOnly + vbInformation, "Informacion"
                                    Exit Sub
                            End If
                        Else
                            MsgBox "Bodega Actual Del Bulto No Existe", vbOKOnly + vbInformation, "Verifique"
                            Exit Sub
                        End If
                End If
            End If
                    
    'VERIFICA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta ", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    'VERIFICA EL TURNO
    If TxtTur.Text = "" Then
        MsgBox "Turno Incorrecto", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    'VERIFICA LA LINEA
    Set RBuscaLinea = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaLinea, "Select * From Lineas Where Linea = '" & TxtLin.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaLinea, "Select * From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
        End If
        If RBuscaLinea.RecordCount > 0 Then
        Else
            MsgBox "Linea No Existe", vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
        
                    'BUSCA SI EXISTE LA TARIMA
                    Set RBuscaTarima = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaTarima, "Select * From DetalleEntradasInventario Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaTarima, "Select * From DetalleEntradasInventario Where FechaProduccion = To_Date('" & VFechaProduccion & "', 'dd/mm/yyyy')" & " And Tarima = " & VTarima & " And UPPER(Linea) = '" & UCase(VLinea) & "' And UPPER(FichaTecnica) = '" & UCase(VFichaTecnica) & "'")
                            End If
                        
                        If RBuscaTarima.RecordCount > 0 Then
                        
                        Else
                            MsgBox "Bulto/Tarima No Existe", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                        
                        
                        
                            
    
    VNuevaExistencia = MskNueExi.Text
    VTotalaDescargar = MskTot.Text
    
    'REVISA QUE EL TOTAL NO SEA MAYOR QUE LO QUE VAN A DESCONTAR
    If VTotalaDescargar > VNuevaExistencia Then
        MsgBox "El Total A Descargar No Puede Ser Mayor Que La Existencia Modificada", vbOKOnly + vbInformation, "Informacion"
        MskCanMas.SetFocus
        Exit Sub
    End If
    
    'BUSCA QUE TODOS LOS TRASLADOS ESTEN LIBERADOS
    Set RBuscaLiberado = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaLiberado, "Select ET.Estado From DetalleTrasladosInventario DT, EncabezadoTrasladosInventario ET Where DT.Tarima = " & VTarima & " And DT.FichaTecnica = '" & VFichaTecnica & "' And FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And LineaProduccion = '" & VLinea & "' And DT.Documento = ET.Documento And ET.Estado = 'NO LIBERADO'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaLiberado, "Select ET.Estado From DetalleTrasladosInventario DT, EncabezadoTrasladosInventario ET Where DT.Tarima = " & VTarima & " And UPPER(DT.FichaTecnica) = '" & UCase(VFichaTecnica) & "' And FechaProduccion = TO_DATE('" & VFechaProduccion & "', 'dd/mm/yyyy')" & " And UPPER(LineaProduccion) = '" & UCase(VLinea) & "' And DT.Documento = ET.Documento And UPPER(ET.Estado) = 'NO LIBERADO'")
        End If
        If RBuscaLiberado.RecordCount > 0 Then
            MsgBox "No Se Puede Grabar, Todavia Hay Traslados Pendientes De Liberar Para Esta Tarima/Bulto", vbOKOnly + vbExclamation, "Verifique"
            Exit Sub
        End If
       
       
                        If GOrigenDeDatos = "AmaproAccess" Then
                             VTexto = "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                        Else 'ORACLE
                             VTexto = "To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                        End If
                        VTexto = VTexto & TxtTur.Text & "', '" 'TURNO
                        VTexto = VTexto & UCase(TxtLin.Text) & "', '" 'LINEA
                        VTexto = VTexto & UCase(TxtBodSal.Text) & "', " 'BODEGA SALIDA
                        VTexto = VTexto & MskExi.Text & ", " 'EXISTENCIA
                        VTexto = VTexto & MskCanMas.Text & ", " 'CANTIDAD DE +
                        VTexto = VTexto & MskCanMen.Text & ", " 'CANTIDAD DE -
                        VTexto = VTexto & MskNueExi.Text & ", " 'NUEVA EXITENCIA
                        VTexto = VTexto & MskConIni.Text & ", " 'CONTADOR INICIAL
                        VTexto = VTexto & MskConFin.Text & ", " 'CONTADOR FINAL
                        VTexto = VTexto & MskCanPro.Text & ", " 'CANTIDAD PROCESADA
                        VTexto = VTexto & MskDesPro.Text & ", " 'DESPERDICIO PROCESO
                        VTexto = VTexto & MskDesProv.Text & ", " 'DESPERDICIO PROVEEDOR
                        VTexto = VTexto & MskCanProRea.Text & ", " 'CANTIDAD PROCESADA REAL
                        VTexto = VTexto & MskTot.Text & ", '" 'TOTAL
                        VTexto = VTexto & UCase(GUsuario) & "', " 'USUARIO
                        If GOrigenDeDatos = "AmaproAccess" Then
                             VTexto = VTexto & "#" & Format(TxtFecPro.Text, "mm/dd/yyyy") & "#, '" 'FECHA PRODUCCION
                        Else 'ORACLE
                             VTexto = VTexto & "To_Date('" & TxtFecPro.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA PRODUCCION
                        End If
                        VTexto = VTexto & UCase(TxtLinPro.Text) & "', '" 'LINEA DE PRODUCCION
                        VTexto = VTexto & UCase(TxtFicTec.Text) & "', " 'FICHA TECNICA
                        VTexto = VTexto & TxtTar.Text & ", '" 'TARIMA
                        VTexto = VTexto & MskHor.Text & "', '" 'HORA
                        VTexto = VTexto & TxtObs.Text & "'" 'OBSERVACIONES
                                           
                        'INICIA UNA TRANSACCION
                       'SI ESTA GRABANDO UN REGISTRO NUEVO
                        Conexion.BeginTrans
                            Conexion.Execute "Insert Into CierreBulto Values(" & VTexto & ")"
                   
                                    'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                        If Err <> 0 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    Else 'ORACLE
                                        If Err <> 0 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    
                                    'MODIFICA EL SALDO DE LA TARIMA/BULTO CON LA NUEVA EXISTENCIA
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Conexion.Execute "Update DetalleEntradasInventario Set Saldo = " & VNuevaExistencia & " Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'"
                                        Else 'ORACLE
                                            Conexion.Execute "Update DetalleEntradasInventario Set Saldo = " & VNuevaExistencia & " Where FechaProduccion = To_Date('" & VFechaProduccion & "', 'dd/mm/yyyy')" & " And Tarima = " & VTarima & " And UPPER(Linea) = '" & UCase(VLinea) & "' And UPPER(FichaTecnica) = '" & UCase(VFichaTecnica) & "'"
                                        End If
                                            
                                            'SI SE DUPLICA LA LLAVE
                                             If GOrigenDeDatos = "AmaproAccess" Then
                                                If Err <> 0 Then
                                                    Conexion.RollbackTrans
                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                    Exit Sub
                                                End If
                                            Else 'ORACLE
                                                If Err <> 0 Then
                                                    Conexion.RollbackTrans
                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                    Exit Sub
                                                End If
                                            End If
                                    
                                    
                                    
                                    'MODIFICA EL SALDO DE LA TARIMA/BULTO Y REBAJA LA CANTIDAD PROCESADA
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo - " & VCantidadMateriaPrima & " Where FechaProduccion = #" & Format(VFechaProduccion, "mm/dd/yyyy") & "# And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FichaTecnica = '" & VFichaTecnica & "'"
                                        Else 'ORACLE
                                            Conexion.Execute "Update DetalleEntradasInventario Set Saldo = Saldo - " & VCantidadMateriaPrima & " Where FechaProduccion = To_Date('" & VFechaProduccion & "', 'dd/mm/yyyy')" & " And Tarima = " & VTarima & " And UPPER(Linea) = '" & UCase(VLinea) & "' And UPPER(FichaTecnica) = '" & UCase(VFichaTecnica) & "'"
                                        End If
                                        
                                            
                                            'SI SE DUPLICA LA LLAVE
                                             If GOrigenDeDatos = "AmaproAccess" Then
                                                If Err <> 0 Then
                                                    Conexion.RollbackTrans
                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                    Exit Sub
                                                End If
                                            Else 'ORACLE
                                                If Err <> 0 Then
                                                    Conexion.RollbackTrans
                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                    Exit Sub
                                                End If
                                            End If
                                            
                        'FINALIZA LA CONEXION
                        Conexion.CommitTrans
                                                
                                                
                        Bandera = False
                        botones
                        CmdAgregar.SetFocus
                        
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RCierreBulto.Requery
                        RCierreBulto.MoveLast
                        Llena_Campos
   
                    
         
   

End Sub


Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
        FrameConsultas.Visible = False
End Sub


Private Sub DBGridConsultas_DblClick()
    'MATERIAS PRIMAS
    If BFichaTecnica = True Then
        TxtFicTec.Text = DBGridConsultas.Columns(0).Text
        TxtFicTec.SetFocus
    'NUMERO DE INGRESO
    ElseIf BNumeroIngreso = True Then
        TxtTar.Text = DBGridConsultas.Columns(1).Text
        TxtTar.SetFocus
    'LINEA
    ElseIf BLinea = True Then
        TxtLin.Text = DBGridConsultas.Columns(0).Text
        TxtLin.SetFocus
    End If
        FrameConsultas.Visible = False
End Sub

Private Sub Dbgridconsultas_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        'MATERIAS PRIMAS
        If BFichaTecnica = True Then
            TxtFicTec.Text = DBGridConsultas.Columns(0).Text
            TxtFicTec.SetFocus
        'NUMERO DE INGRESO
        ElseIf BNumeroIngreso = True Then
            TxtTar.Text = DBGridConsultas.Columns(1).Text
            TxtTar.SetFocus
        'LINEA
        ElseIf BLinea = True Then
            TxtLin.Text = DBGridConsultas.Columns(0).Text
            TxtLin.SetFocus
        End If
        FrameConsultas.Visible = False
    End If
End Sub

Private Sub dbgridcierrebulto_HeadClick(ByVal ColIndex As Integer)
            RCierreBulto.Sort = RCierreBulto.Fields(ColIndex).Name
End Sub

Private Sub Form_Load()
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date
        
        Set RCierreBulto = New ADODB.Recordset
                
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RCierreBulto, "Select * From CierreBulto") '  Where Fecha = #" & Format(Date, "mm/dd/yyyy") & "#")
                Else 'ORACLE
                    Call Abrir_Recordset(RCierreBulto, "Select * From CierreBulto") ' Where Fecha = To_Date('" & Date & "', 'dd/mm/yyyy')")
                End If
                
            Set DbGridCierreBulto.DataSource = RCierreBulto
            Llena_Campos
            
End Sub

Private Sub MskCanMas_GotFocus()
        MskCanMas.SelStart = 0
        MskCanMas.SelLength = Len(MskCanMas.Text)
End Sub

Private Sub MskCanMas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskCanMas_LostFocus()
        Calcula
End Sub

Private Sub MskCanMen_GotFocus()
        MskCanMen.SelStart = 0
        MskCanMen.SelLength = Len(MskCanMen.Text)
End Sub

Private Sub MskCanMen_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskCanMen_LostFocus()
        Calcula
End Sub

Private Sub MskCanPro_GotFocus()
        MskCanPro.SelStart = 0
        MskCanPro.SelLength = Len(MskCanPro.Text)
End Sub

Private Sub MskCanPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub


Private Sub MskCanPro_LostFocus()
        Calcula
End Sub

Private Sub MskConFin_GotFocus()
        MskConFin.SelStart = 0
        MskConFin.SelLength = Len(MskConFin.Text)
End Sub

Private Sub MskConFin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskConFin_LostFocus()
        Calcula
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

Private Sub MskConIni_LostFocus()
        Calcula
End Sub

Private Sub MskDesPro_GotFocus()
        MskDesPro.SelStart = 0
        MskDesPro.SelLength = Len(MskDesPro.Text)
        
End Sub

Private Sub MskDesPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskDesPro_LostFocus()
        Calcula
End Sub

Private Sub MskDesProv_GotFocus()
        MskDesProv.SelStart = 0
        MskDesProv.SelLength = Len(MskDesProv.Text)
End Sub

Private Sub MskDesProv_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskDesProv_LostFocus()
        Calcula
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

Private Sub MskHor_GotFocus()
        MskHor.SelStart = 0
        MskHor.SelLength = Len(MskHor.Text)
End Sub

Private Sub MskHor_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub


Private Sub TabIngresos_Click(PreviousTab As Integer)
        If TabIngresos.Tab = 0 Then
            CmdBorrar.Enabled = True
            If CmdGrabar.Enabled = False Then
                Llena_Campos
            End If
        Else
            CmdBorrar.Enabled = False
        End If
        
        
End Sub

Private Sub TxtBodSal_Change()
    Set RBuscaBodega = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaBodega, "Select * From BodegasInventario Where CodigoBodega = '" & TxtBodSal.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaBodega, "Select * From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBodSal.Text) & "'")
        End If
        If RBuscaBodega.RecordCount > 0 Then
            LblBodega.Caption = RBuscaBodega!Descripcion
        Else
            LblBodega.Caption = ""
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


Private Sub TxtFecPro_GotFocus()
        TxtFecPro.SelStart = 0
        TxtFecPro.SelLength = Len(TxtFecPro.Text)
End Sub

Private Sub TxtFecPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtFicTec_Change()
        Set RBuscaFichaTecnica = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where Esp_Tec = '" & TxtFicTec.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            
            If RBuscaFichaTecnica.RecordCount > 0 Then
                LblMateriaPrima.Caption = RBuscaFichaTecnica!Descrip
            Else
                LblMateriaPrima.Caption = ""
            End If
            
End Sub

Private Sub TxtFicTec_DblClick()
        BNumeroIngreso = False
        BFichaTecnica = True
        BLinea = False
        Set RBusqueda = New ADODB.Recordset
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
        Set DBGridConsultas.DataSource = RBusqueda
        FrameConsultas.Visible = True
        TxtConsultas.SetFocus
        DBGridConsultas.Columns(1).Width = "3000"
End Sub

Private Sub TxtFicTec_GotFocus()
        TxtFicTec.SelStart = 0
        TxtFicTec.SelLength = Len(TxtFicTec.Text)
End Sub

Private Sub TxtFicTec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
            BNumeroIngreso = False
            BFichaTecnica = True
            BLinea = False
            'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            Set DBGridConsultas.DataSource = RBusqueda
            FrameConsultas.Visible = True
            TxtConsultas.SetFocus
            DBGridConsultas.Columns(1).Width = "3000"
        End If
End Sub

Private Sub TxtConsultas_Change()
    Set RBusqueda = New ADODB.Recordset
    'MATERIA PRIMA
    If BFichaTecnica = True Then
        If OptDes.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '%" & TxtConsultas.Text & "%' Order By Esp_Tec")
            Else 'ORACLE
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtConsultas.Text) & "%' Order By Esp_Tec")
            End If
        ElseIf OptCod.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBusqueda, "Select * From FichaTecnica Where Esp_Tec Like '%" & TxtConsultas.Text & "%' Order By Esp_Tec")
            Else 'ORACLE
                Call Abrir_Recordset(RBusqueda, "Select * From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtConsultas.Text) & "%' Order By Esp_Tec")
            End If
        End If
    'LINEA
    ElseIf BLinea = True Then
        If OptDes.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBusqueda, "Select * From Lineas Where Descrip Like '%" & TxtConsultas.Text & "%'")
            Else 'ORACLE
                Call Abrir_Recordset(RBusqueda, "Select * From Lineas Where UPPER(Descrip) Like '%" & UCase(TxtConsultas.Text) & "%'")
            End If
            
        ElseIf OptCod.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBusqueda, "Select * From Lineas Where Linea Like '%" & TxtConsultas.Text & "%'")
            Else 'ORACLE
                Call Abrir_Recordset(RBusqueda, "Select * From Lineas Where UPPER(Linea) Like '%" & UCase(TxtConsultas.Text) & "%'")
            End If
        End If
    End If
    
    Set DBGridConsultas.DataSource = RBusqueda
    
    
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

Private Sub TxtLin_Change()
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLin.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
                LblLinea.Caption = RBuscaLinea!Descrip
            Else
                LblLinea.Caption = ""
            End If
            
End Sub

Private Sub Txtlin_DblClick()
        BNumeroIngreso = False
        BFichaTecnica = False
        BLinea = True
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas")
        Set DBGridConsultas.DataSource = RBusqueda
        FrameConsultas.Visible = True
        TxtConsultas.SetFocus
        DBGridConsultas.Columns(1).Width = "3000"
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
            BNumeroIngreso = False
            BFichaTecnica = False
            BLinea = True
            'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas")
            Set DBGridConsultas.DataSource = RBusqueda
            FrameConsultas.Visible = True
            TxtConsultas.SetFocus
            DBGridConsultas.Columns(1).Width = "3000"
        End If
End Sub

Private Sub TxtLinPro_Change()
        Set RBuscaLineaP = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLineaP, "Select Descrip From Lineas Where Linea = '" & TxtLinPro.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLineaP, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinPro.Text) & "'")
            End If
            If RBuscaLineaP.RecordCount > 0 Then
                LblLinPro.Caption = RBuscaLineaP!Descrip
            Else
                LblLinPro.Caption = ""
            End If
End Sub

Private Sub TxtLinPro_GotFocus()
            TxtLinPro.SelStart = 0
            TxtLinPro.SelLength = Len(TxtLinPro.Text)
End Sub

Private Sub TxtLinPro_KeyPress(KeyAscii As Integer)
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

Private Sub TxtTar_Change()
        'BUSCA LA UNIDAD DE MEDIDA
        If IsNumeric(TxtTar.Text) And IsDate(TxtFecPro.Text) Then
                'BUSCA EL NUMERO DE INGRESO
                Set RBuscaUnidad = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaUnidad, "Select UnidadMedida From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaUnidad, "Select UnidadMedida From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
                    End If
                If RBuscaUnidad.RecordCount > 0 Then
                    If IsNull(RBuscaUnidad(0)) Then
                        LblUnidadMedida.Caption = ""
                    Else
                        LblUnidadMedida.Caption = RBuscaUnidad!unidadMedida
                        
                        Set RCuentaBultos = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RCuentaBultos, "Select Count(*) From CierreBulto Where FechaProduccion = #" & Format(TxtFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTar.Text & " And Linea = '" & TxtLin.Text & "' And FichaTecnica = '" & TxtFicTec.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RCuentaBultos, "Select Count(*) From CierreBulto Where FechaProduccion = To_Date('" & TxtFecPro.Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTar.Text & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtFicTec.Text) & "'")
                            End If
                                            
                            If RCuentaBultos.RecordCount > 0 Then
                                LblCanCieBul.Caption = RCuentaBultos(0)
                            Else
                                LblCanCieBul.Caption = "0"
                            End If
                        
                    End If
                Else
                    LblUnidadMedida.Caption = ""
                End If
                
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



Private Sub TxtTar_LostFocus()

                
                TxtFecPro.Text = Format(TxtFecPro.Text, "dd/mm/yyyy")
                
                    If IsNumeric(TxtTar.Text) Then
                        
                            If TxtFecPro.Text = "" Then
                                    Set RBuscaTarima = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RBuscaTarima, "Select FechaProduccion From DetalleEntradasInventario Where Tarima = " & TxtTar.Text & " And FichaTecnica = '" & TxtFicTec.Text & "' And Linea = '" & TxtLinPro.Text & "'")
                                            Else 'ORACLE
                                                Call Abrir_Recordset(RBuscaTarima, "Select FechaProduccion From DetalleEntradasInventario Where Tarima = " & TxtTar.Text & " And UPPER(FichaTecnica) = '" & UCase(TxtFicTec.Text) & "' And Linea = '" & TxtLinPro.Text & "'")
                                            End If
                                        
                                        If RBuscaTarima.RecordCount > 0 Then
                                                TxtFecPro.Text = RBuscaTarima!FechaProduccion
                                        Else
                                            MsgBox "Ficha Tecnica Con Este Bulto No Existe Y Linea", vbOKOnly + vbInformation, "Informacion"
                                            Exit Sub
                                        End If
                            End If
                        
                    End If
        

                
                If IsNumeric(TxtTar.Text) And IsDate(TxtFecPro.Text) Then
                    'BUSCA SI EXISTE LA TARIMA
                    Set RBuscaTarima = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaTarima, "Select Saldo, Bodega From DetalleEntradasInventario Where FechaProduccion = #" & Format(TxtFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTar.Text & " And Linea = '" & TxtLinPro.Text & "' And FichaTecnica = '" & TxtFicTec.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaTarima, "Select Saldo, Bodega From DetalleEntradasInventario Where FechaProduccion = To_Date('" & TxtFecPro.Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTar.Text & " And UPPER(Linea) = '" & UCase(TxtLinPro.Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtFicTec.Text) & "'")
                            End If
        
                                'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
                                If RBuscaTarima.RecordCount > 0 Then
                                    If IsNull(RBuscaTarima(0)) Then
                                        TxtBodSal.Text = ""
                                        MskExi.Text = ""
                                        LblUnidadMedida.Caption = ""
                                    Else
                                        TxtBodSal.Text = RBuscaTarima!Bodega
                                        MskExi.Text = RBuscaTarima!Saldo
                                        
                                        Set RCuentaBultos = New ADODB.Recordset
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RCuentaBultos, "Select Count(*) From CierreBulto Where FechaProduccion = #" & Format(TxtFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTar.Text & " And LineaProduccion = '" & TxtLinPro.Text & "' And FichaTecnica = '" & TxtFicTec.Text & "'")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RCuentaBultos, "Select Count(*) From CierreBulto Where FechaProduccion = To_Date('" & TxtFecPro.Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTar.Text & " And UPPER(LineaProduccion) = '" & UCase(TxtLinPro.Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtFicTec.Text) & "'")
                                                End If
                                            
                                                    If RCuentaBultos.RecordCount > 0 Then
                                                        LblCanCieBul.Caption = RCuentaBultos(0)
                                                    Else
                                                        LblCanCieBul.Caption = "0"
                                                    End If
                                    End If
                                 Else
                                    TxtBodSal.Text = ""
                                    MskExi.Text = "0"
                                    LblUnidadMedida.Caption = ""
                                    MsgBox "Tarima/Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                                End If
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

Sub Calcula()
        'CANTIDAD PROCESADA ES IGUAL AL CONTADOR FINAL MENOS EL CONTADOR INICIAL
        MskCanPro.Text = Val(MskConFin.Text) - Val(MskConIni.Text)
        
        'A LA EXISTENCIA ACTUAL LE SUMA LA CANTIDAD DE MENOS
        MskNueExi.Text = Val(MskExi.Text) + Val(MskCanMas.Text)
        'AL RESULTADO DE LA EXISTENCIA MAS LA CANTIDAD DE MENOS LE RESTA
        'LA CANTIDAD DE MENOS
        MskNueExi.Text = Val(MskNueExi.Text) - Val(MskCanMen.Text)
        
        'TOTAL A DESCONTAR DE INVENTARIO ES IGUAL A LO PROCESADO MAS EL DESPERDICIO DE PROVEEDOR
        MskTot.Text = Val(MskCanPro.Text) + Val(MskDesProv.Text)
        'REAL PROCESADO ES IGUAL A LA CANTIADAD PROCESADA MENOS EL DESPERDICIO DE PROCESO
        MskCanProRea.Text = Val(MskCanPro.Text) - Val(MskDesPro.Text)
End Sub


Public Sub Llena_Campos()
On Error Resume Next
        'FECHA
            If IsNull(RCierreBulto!fecha) Then
                MskFec.Text = ""
            Else
                MskFec.Text = RCierreBulto!fecha
            End If
        'TURNO
            If IsNull(RCierreBulto!Turno) Then
                TxtTur.Text = ""
            Else
                TxtTur.Text = RCierreBulto!Turno
            End If
        'LINEA
            If IsNull(RCierreBulto!Linea) Then
                TxtLin.Text = ""
            Else
                TxtLin.Text = RCierreBulto!Linea
            End If
        'BODEGA SALIDA
            If IsNull(RCierreBulto!BodegaSalida) Then
                TxtBodSal.Text = ""
            Else
                TxtBodSal.Text = RCierreBulto!BodegaSalida
            End If
        'EXISTENCIA
            If IsNull(RCierreBulto!Existencia) Then
                MskExi.Text = ""
            Else
                MskExi.Text = RCierreBulto!Existencia
            End If
        'DE MAS
            If IsNull(RCierreBulto!CantidadMas) Then
                MskCanMas.Text = ""
            Else
                MskCanMas.Text = RCierreBulto!CantidadMas
            End If
        'DE MENOS
            If IsNull(RCierreBulto!CantidadMenos) Then
                MskCanMen.Text = ""
            Else
                MskCanMen.Text = RCierreBulto!CantidadMenos
            End If
        'EXISTENICA NUEVA
            If IsNull(RCierreBulto!ExistenciaNueva) Then
                MskNueExi.Text = ""
            Else
                MskNueExi.Text = RCierreBulto!ExistenciaNueva
            End If
        'CONTADOR INICIAL
            If IsNull(RCierreBulto!ContadorInicial) Then
                MskConIni.Text = ""
            Else
                MskConIni.Text = RCierreBulto!ContadorInicial
            End If
        'CONTADOR FINAL
            If IsNull(RCierreBulto!ContadorFinal) Then
                MskConFin.Text = ""
            Else
                MskConFin.Text = RCierreBulto!ContadorFinal
            End If
        'CANTIDAD PROCESADA
            If IsNull(RCierreBulto!CantidadProcesada) Then
                MskCanPro.Text = ""
            Else
                MskCanPro.Text = RCierreBulto!CantidadProcesada
            End If
        'DESPERDICIO PROCESO
            If IsNull(RCierreBulto!DesperdicioProceso) Then
                MskDesPro.Text = ""
            Else
                MskDesPro.Text = RCierreBulto!DesperdicioProceso
            End If
        'DESPERDICIO PROVEEDOR
            If IsNull(RCierreBulto!DesperdicioProveedor) Then
                MskDesProv.Text = ""
            Else
                MskDesProv.Text = RCierreBulto!DesperdicioProveedor
            End If
        'CANTIDAD PROCESADA REAL
            If IsNull(RCierreBulto!CantidadProcesadaReal) Then
                MskCanProRea.Text = ""
            Else
                MskCanProRea.Text = RCierreBulto!CantidadProcesadaReal
            End If
        'TOTAL
            If IsNull(RCierreBulto!total) Then
                MskTot.Text = ""
            Else
                MskTot.Text = RCierreBulto!total
            End If
        'USUARIO
            If IsNull(RCierreBulto!Usuario) Then
                TxtUsu.Text = ""
            Else
                TxtUsu.Text = RCierreBulto!Usuario
            End If
        'FECHA PRODUCCION
            If IsNull(RCierreBulto!FechaProduccion) Then
                TxtFecPro.Text = ""
            Else
                TxtFecPro.Text = RCierreBulto!FechaProduccion
            End If
        'LINEA PRODUCICON
            If IsNull(RCierreBulto!LineaProduccion) Then
                TxtLinPro.Text = ""
            Else
                TxtLinPro.Text = RCierreBulto!LineaProduccion
            End If
        'FICHA TECNICA
            If IsNull(RCierreBulto!FichaTecnica) Then
                TxtFicTec.Text = ""
            Else
                TxtFicTec.Text = RCierreBulto!FichaTecnica
            End If
        'TARIMA
            If IsNull(RCierreBulto!Tarima) Then
                TxtTar.Text = ""
            Else
                TxtTar.Text = RCierreBulto!Tarima
            End If
        'HORA
            If IsNull(RCierreBulto!Hora) Then
                MskHor.Text = ""
            Else
                MskHor.Text = RCierreBulto!Hora
            End If
        'OBSERVACIONES
            If IsNull(RCierreBulto!Observaciones) Then
                TxtObs.Text = ""
            Else
                TxtObs.Text = RCierreBulto!Observaciones
            End If
                
        If Err <> 0 Then
            
        End If

End Sub

Public Sub Limpia_Campos()
        
                MskFec.Text = ""
                TxtTur.Text = ""
                TxtLin.Text = ""
                TxtBodSal.Text = ""
                MskExi.Text = 0
                MskCanMas.Text = 0
                MskCanMen.Text = 0
                MskNueExi.Text = 0
                MskConIni.Text = 0
                MskConFin.Text = 0
                MskCanPro.Text = 0
                MskDesPro.Text = 0
                MskDesProv.Text = 0
                MskCanProRea.Text = 0
                MskTot.Text = 0
                TxtUsu.Text = ""
                TxtFecPro.Text = ""
                TxtLinPro.Text = ""
                TxtFicTec.Text = ""
                TxtTar.Text = 0
                MskHor.Text = "00:00"
                TxtObs.Text = ""
        
        
End Sub






