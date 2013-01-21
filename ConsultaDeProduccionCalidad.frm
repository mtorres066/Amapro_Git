VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ConsultaDeProduccionCalidad 
   BackColor       =   &H000000FF&
   Caption         =   "Consulta De Captura De Produccion En Planta"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ConsultaDeProduccionCalidad.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
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
      Height          =   6855
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Visible         =   0   'False
      Width           =   8535
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   5655
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9975
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
         ColumnCount     =   3
         BeginProperty Column00 
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
            DataField       =   "Grupo"
            Caption         =   "Grupo"
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
               Locked          =   -1  'True
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   4529.764
            EndProperty
            BeginProperty Column02 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "ConsultaDeProduccionCalidad.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7695
      Left            =   0
      TabIndex        =   25
      Top             =   840
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   13573
      _Version        =   393216
      TabHeight       =   520
      BackColor       =   255
      TabCaption(0)   =   "Resumen Produccion"
      TabPicture(0)   =   "ConsultaDeProduccionCalidad.frx":21BC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DbgridMes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DBGridLineasFichaTecnica"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DbGridGerencia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle Produccion"
      TabPicture(1)   =   "ConsultaDeProduccionCalidad.frx":21D8
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridDia"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DataGridDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Detalle Paros"
      TabPicture(2)   =   "ConsultaDeProduccionCalidad.frx":21F4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DbGridParos"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DataGridParos"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).ControlCount=   2
      Begin MSDataGridLib.DataGrid DbGridGerencia 
         Height          =   7095
         Left            =   120
         TabIndex        =   26
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   12515
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
            Weight          =   700
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
         Caption         =   "Produccion Por Ficha Tecnica"
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "Esp_Tec"
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
            DataField       =   "Expr1002"
            Caption         =   "Tarimas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Expr1003"
            Caption         =   "Unidades"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4155.024
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   645.165
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DBGridLineasFichaTecnica 
         Height          =   2535
         Left            =   7800
         TabIndex        =   27
         Top             =   480
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
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
            Weight          =   700
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
         Caption         =   "Produccion x Linea"
         ColumnCount     =   4
         BeginProperty Column00 
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
            DataField       =   "Expr1002"
            Caption         =   "Tarimas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Expr1003"
            Caption         =   "Unidades"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   345.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1709.858
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   464.882
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DbgridMes 
         Height          =   4455
         Left            =   7800
         TabIndex        =   28
         Top             =   3120
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   7858
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
            Weight          =   700
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
         Caption         =   "Produccion x Año y Mes"
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
      Begin MSDataGridLib.DataGrid DataGridDetalle 
         Height          =   7215
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
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
            Weight          =   700
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
         Caption         =   "Produccion Por Linea, Fecha y Ficha Tecnica"
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
      Begin MSDataGridLib.DataGrid DbGridDia 
         Height          =   7215
         Left            =   -66960
         TabIndex        =   30
         Top             =   360
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   12726
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "Fec_Prd"
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
            DataField       =   "Expr1001"
            Caption         =   "Tarimas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Expr1002"
            Caption         =   "Unidades"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DataGridParos 
         Height          =   7095
         Left            =   -74880
         TabIndex        =   31
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   12515
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
         HeadLines       =   1
         RowHeight       =   15
         TabAcrossSplits =   -1  'True
         TabAction       =   2
         WrapCellPointer =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
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
         Caption         =   "Paros x Linea"
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
      Begin MSDataGridLib.DataGrid DbGridParos 
         Height          =   6975
         Left            =   -66960
         TabIndex        =   32
         Top             =   480
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   12303
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8438015
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
            Weight          =   700
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
         Caption         =   "Paros x Linea"
         ColumnCount     =   4
         BeginProperty Column00 
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
            DataField       =   "Expr1003"
            Caption         =   "Horas"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   345.26
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1874.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   239.811
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   705.26
            EndProperty
         EndProperty
      End
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H000000FF&
      Caption         =   "PNC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   2040
      TabIndex        =   16
      Top             =   600
      Width           =   735
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H000000FF&
      Caption         =   "PC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   1
      Left            =   2040
      TabIndex        =   15
      Top             =   360
      Width           =   855
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H000000FF&
      Caption         =   "Todos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   2040
      TabIndex        =   14
      Top             =   120
      Value           =   -1  'True
      Width           =   852
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Caption         =   "Tipo De Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton OptLinea 
         BackColor       =   &H000000FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   0
         TabIndex        =   18
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton OptGrupo 
         BackColor       =   &H000000FF&
         Caption         =   "Grupo De Linea"
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
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.TextBox TxtLinea 
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
      Left            =   3840
      MaxLength       =   10
      TabIndex        =   2
      ToolTipText     =   "doble click o signo '+' para ayuda"
      Top             =   480
      Width           =   1455
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   615
      Left            =   11160
      Picture         =   "ConsultaDeProduccionCalidad.frx":2210
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Salida"
      Top             =   120
      Width           =   615
   End
   Begin MSMask.MaskEdBox MskTotalEnvases 
      Height          =   255
      Left            =   9000
      TabIndex        =   7
      Top             =   480
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MskTotalTarimas 
      Height          =   255
      Left            =   9000
      TabIndex        =   6
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "#,###,##0"
      PromptChar      =   "_"
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   255
      Left            =   6000
      TabIndex        =   1
      Top             =   120
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   52035587
      CurrentDate     =   37248
   End
   Begin MSComCtl2.DTPicker DtpFecIni 
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   52035587
      CurrentDate     =   37248
   End
   Begin VB.Label LblLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   12
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label LblDescripcion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Grupo"
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
      Left            =   3240
      TabIndex        =   11
      Top             =   480
      Width           =   525
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Cantidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   8040
      TabIndex        =   9
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Tarimas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   8040
      TabIndex        =   8
      Top             =   120
      Width           =   765
   End
   Begin MSForms.CommandButton CmdGenera 
      Height          =   615
      Left            =   10440
      TabIndex        =   5
      ToolTipText     =   "Generar Datos"
      Top             =   120
      Width           =   615
      PicturePosition =   327683
      Size            =   "1085;1085"
      Picture         =   "ConsultaDeProduccionCalidad.frx":272B
      FontEffects     =   1073741825
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   510
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
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
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "ConsultaDeProduccionCalidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RTotal As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RTarimas As New ADODB.Recordset
Dim BLinea As Boolean
Dim BGrupo As Boolean

Dim RFicha As New ADODB.Recordset
Dim RLinea As New ADODB.Recordset
Dim RMes As New ADODB.Recordset
Dim RDia As New ADODB.Recordset
Dim RParos As New ADODB.Recordset
Dim RDetalle As New ADODB.Recordset
Dim RDetalle2 As New ADODB.Recordset
Dim RParos2 As New ADODB.Recordset




Private Sub CmdGenera_Click()
On Error Resume Next
MousePointer = 11

'_______________________________________________________________________________________________________________________
            'GRID DE FICHA TECNICA
            Set RFicha = New ADODB.Recordset
            If Opcion.Item(0).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Esp_Tec, F.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By P.Esp_Tec, F.Descrip")
                    Else
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Esp_Tec, F.Descrip")
                    End If
                End If
                
            ElseIf Opcion.Item(1).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') Group By P.Esp_Tec, F.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= TO_DATE('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') Group By P.Esp_Tec, F.Descrip")
                    End If
                End If
                
            ElseIf Opcion.Item(2).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' And (P.Calidad = 'R') Group By P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' And (UPPER(P.Calidad) = 'R') Group By P.Esp_Tec, F.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' And (P.Calidad = 'R') Group By P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RFicha, "SELECT P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' And (UPPER(P.Calidad) = 'R') Group By P.Esp_Tec, F.Descrip")
                    End If
                End If
            End If
            
            Set DbGridGerencia.DataSource = RFicha
            
            

'_______________________________________________________________________________________________________________________
            'EL GRID DE LINEAS Y FICHA TECNICA
            
            Set RLinea = New ADODB.Recordset
            If Opcion.Item(0).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By P.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Linea, L.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By P.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Linea, L.Descrip")
                    End If
                End If
            ElseIf Opcion.Item(1).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By P.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Linea, L.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By P.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Linea, L.Descrip")
                    End If
                End If
            ElseIf Opcion.Item(2).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By P.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Calidad) = 'R' And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Linea, L.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By P.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RLinea, "Select P.Linea, L.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L Inner Join Produccion P On P.Linea = L.Linea Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Calidad) = 'R' And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Linea, L.Descrip")
                    End If
                End If
            End If
            
            Set DBGridLineasFichaTecnica.DataSource = RLinea
            
            
            
            
'_______________________________________________________________________________________________________________________
            'EL GRID DE MES
            'TODOS
            Set RMes = New ADODB.Recordset
            
            If Opcion.Item(0).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RMes, "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RMes, "Select To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM'), Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM') Order By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM')")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RMes, "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion P, Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea Like '" & TxtLinea.Text & "%' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RMes, "Select To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM'), Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM') Order By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM')")
                    End If
                End If
            'PC
            ElseIf Opcion.Item(1).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RMes, "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RMes, "Select To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM'), Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') and UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM') Order By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM')")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RMes, "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RMes, "Select To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM'), Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM') Order By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM')")
                    End If
                End If
            'PNC
            ElseIf Opcion.Item(2).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RMes, "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RMes, "Select To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM'), Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Calidad) = 'R' and UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM') Order By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM')")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RMes, "Select month(P.Fec_Prd), Year(P.Fec_Prd), Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By month(P.Fec_Prd), year(P.Fec_Prd) Order By Year(P.Fec_Prd)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RMes, "Select To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM'), Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Calidad) = 'R' And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM') Order By To_Char(P.Fec_Prd,'YYYY'), To_Char(P.Fec_Prd,'MM')")
                    End If
                End If
            End If
            
            Set DbgridMes.DataSource = RMes
            
            DbgridMes.Columns(0).Width = "600"
            DbgridMes.Columns(1).Width = "600"
            DbgridMes.Columns(2).Width = "1000"
            DbgridMes.Columns(3).Width = "1000"
           
            DbgridMes.Columns(2).Alignment = dbgRight
            DbgridMes.Columns(3).Alignment = dbgRight
           
            DbgridMes.Columns(2).NumberFormat = "#,###,###"
            DbgridMes.Columns(3).NumberFormat = "#,###,###"
            
            DbgridMes.Columns(0).Caption = "Año"
            DbgridMes.Columns(1).Caption = "Mes"
            DbgridMes.Columns(2).Caption = "Tarimas"
            DbgridMes.Columns(3).Caption = "Unidades"
            
            
            
'_______________________________________________________________________________________________________________________
            
            Set RParos = New ADODB.Recordset
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo")
                    Else 'ORACLE
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) = '" & UCase(TxtLinea.Text) & "' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) Group By EP.Linea, L.Descrip, P.Tipo")
                        End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & TxtLinea.Text & "' And EP.Linea = L.Linea And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo")
                    Else 'ORACLE
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = '" & UCase(TxtLinea.Text) & "' And EP.Linea = L.Linea And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) Group By EP.Linea, L.Descrip, P.Tipo")
                    End If
                End If
            
            Set DbGridParos.DataSource = RParos

            
    
'_______________________________________________________________________________________________________________________
            'EL GRID DE DIA
            Set RDia = New ADODB.Recordset
            If Opcion.Item(0).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea Like '" & TxtLinea.Text & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    End If
                End If
            'PC
            ElseIf Opcion.Item(1).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') and UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    End If
                End If
            'PNC
            ElseIf Opcion.Item(2).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Calidad) = 'R' and UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDia, "Select P.Fec_Prd, Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Calidad) = 'R' And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Fec_Prd Order By P.Fec_Prd")
                    End If
                End If
            End If
            
            
            Set DbGridDia.DataSource = RDia
            
'_______________________________________________________________________________________________________________________
            'CUENTA LAS TARIMAS
            Set RTarimas = New ADODB.Recordset
            
            If Opcion.Item(0).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%'")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea Like '" & TxtLinea.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%'")
                    End If
                End If
            'PC
            ElseIf Opcion.Item(1).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') and UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%'")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%'")
                    End If
                End If
            'PNC
            ElseIf Opcion.Item(2).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Calidad) = 'R' and UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%'")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P, Lineas L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RTarimas, "Select Count(P.Tarima), Sum(P.Envases) From Produccion P Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Calidad) = 'R' And UPPER(P.Linea) Like '" & UCase(TxtLinea.Text) & "%'")
                    End If
                End If
            End If
            
'
            If RTarimas.RecordCount > 0 Then
                If Not IsNull(RTarimas(0)) Then
                    MskTotalTarimas.Text = RTarimas(0)
                Else
                    MskTotalTarimas.Text = 0
                End If
                
                If Not IsNull(RTarimas(1)) Then
                    MskTotalEnvases.Text = RTarimas(1)
                Else
                    MskTotalEnvases.Text = 0
                End If
            Else
                    MskTotalTarimas.Text = 0
                    MskTotalEnvases.Text = 0
            End If
'_______________________________________________________________________________________________________________________
            
            'SUMA LOS ENVASES
'            If Opcion.Item(0).Value = True Then
'                If OptTodos.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#")
'                ElseIf OptGrupo.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
'                ElseIf OptLinea.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
'                End If
'            ElseIf Opcion.Item(1).Value = True Then
'                If OptTodos.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C')")
'                ElseIf OptGrupo.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
'                ElseIf OptLinea.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
'                End If
'            ElseIf Opcion.Item(2).Value = True Then
'                If OptTodos.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R'")
'                ElseIf OptGrupo.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "'")
'                ElseIf OptLinea.Value = True Then
'                    Set RTotal = Db.OpenRecordset("Select Sum(P.Envases) From Produccion as P, Lineas as L Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And P.Calidad = 'R' And P.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "'")
'                End If
                
'            End If
            
'            If RTotal.RecordCount > 0 Then
'                If Not IsNull(RTotal(0)) Then
'                    MskTotalEnvases.Text = RTotal(0)
'                Else
'                    MskTotalEnvases.Text = "0"
'                End If
'            Else
'                MskTotalEnvases.Text = 0
'            End If
            
'_______________________________________________________________________________________________________________________
            'DETALLE
            Set RDetalle = New ADODB.Recordset
            If Opcion.Item(0).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    Else
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    End If
                End If
                
            ElseIf Opcion.Item(1).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' And (P.Calidad = 'A' OR P.Calidad = 'I' Or P.Calidad = 'C') Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= TO_DATE('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' And (UPPER(P.Calidad) = 'A' OR UPPER(P.Calidad) = 'I' Or UPPER(P.Calidad) = 'C') Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    End If
                End If
                
            ElseIf Opcion.Item(2).Value = True Then
                If OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Grupo Like '" & TxtLinea.Text & "%' And (P.Calidad = 'R') Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) Like '" & UCase(TxtLinea.Text) & "%' And (UPPER(P.Calidad) = 'R') Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# and P.Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And P.Linea = L.Linea And L.Linea Like '" & TxtLinea.Text & "%' And (P.Calidad = 'R') Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDetalle, "SELECT P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip, Count(P.Tarima), Sum(P.Envases) From Lineas L, Produccion P INNER JOIN FichaTecnica F ON P.Esp_Tec = F.ESP_TEC Where P.Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and P.Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(P.Linea) = UPPER(L.Linea) And UPPER(L.Linea) Like '" & UCase(TxtLinea.Text) & "%' And (UPPER(P.Calidad) = 'R') Group By P.Linea, P.Fec_Prd, P.Esp_Tec, F.Descrip")
                    End If
                End If
            End If
            
            Set DataGridDetalle.DataSource = RDetalle
            
            DataGridDetalle.Columns(0).Width = "300"
            DataGridDetalle.Columns(1).Width = "1000"
            DataGridDetalle.Columns(2).Width = "1100"
            DataGridDetalle.Columns(3).Width = "3500"
            DataGridDetalle.Columns(4).Width = "400"
            DataGridDetalle.Columns(5).Width = "700"
           
            DataGridDetalle.Columns(4).Alignment = dbgRight
            DataGridDetalle.Columns(5).Alignment = dbgRight
           
            DataGridDetalle.Columns(1).NumberFormat = "dd/mm/yyyy"
            DataGridDetalle.Columns(4).NumberFormat = "#,###,###"
            DataGridDetalle.Columns(5).NumberFormat = "#,###,###"
            
            
            DataGridDetalle.Columns(0).Caption = "Linea"
            DataGridDetalle.Columns(1).Caption = "Fecha"
            DataGridDetalle.Columns(2).Caption = "Ficha Tecnica"
            DataGridDetalle.Columns(3).Caption = "Descripcion"
            DataGridDetalle.Columns(4).Caption = "Tarimas"
            DataGridDetalle.Columns(5).Caption = "Unidades"
            
        'SOLO SI ES OPCION POR LINEA
        'If OptLinea.Value = True Then
                'PAROS DETALLE__________________________________________________________________________________________________________
                Set RParos2 = New ADODB.Recordset
                               If OptGrupo.Value = True Then
                                   If GOrigenDeDatos = "AmaproAccess" Then
                                       Call Abrir_Recordset(RParos2, "SELECT EP.Linea, L.Descrip, P.Tipo, DP.Paro, P.DescripcionParo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo, DP.Paro, P.DescripcionParo Order By EP.Linea, P.Tipo, Sum(DP.Minutos/60)")
                                   Else 'ORACLE
                                       Call Abrir_Recordset(RParos2, "SELECT EP.Linea, L.Descrip, P.Tipo, DP.Paro, P.DescripcionParo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) = '" & UCase(TxtLinea.Text) & "' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) Group By EP.Linea, L.Descrip, P.Tipo, DP.Paro, P.DescripcionParo Order By EP.Linea, P.Tipo, Sum(DP.Minutos/60)")
                                       End If
                               ElseIf OptLinea.Value = True Then
                                   If GOrigenDeDatos = "AmaproAccess" Then
                                       Call Abrir_Recordset(RParos2, "SELECT EP.Linea, L.Descrip, P.Tipo, DP.Paro, P.DescripcionParo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & TxtLinea.Text & "' And EP.Linea = L.Linea And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo, DP.Paro, P.DescripcionParo Order By EP.Linea, P.Tipo, Sum(DP.Minutos/60)")
                                   Else 'ORACLE
                                       Call Abrir_Recordset(RParos2, "SELECT EP.Linea, L.Descrip, P.Tipo, DP.Paro, P.DescripcionParo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = '" & UCase(TxtLinea.Text) & "' And EP.Linea = L.Linea And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) Group By EP.Linea, L.Descrip, P.Tipo, DP.Paro, P.DescripcionParo Order By EP.Linea, P.Tipo, Sum(DP.Minutos/60)")
                                   End If
                               End If
                           
                           Set DataGridParos.DataSource = RParos2
                           
                           DataGridParos.Columns(0).Width = "300"
                           DataGridParos.Columns(1).Width = "1200"
                           DataGridParos.Columns(2).Width = "200"
                           DataGridParos.Columns(3).Width = "600"
                           DataGridParos.Columns(4).Width = "4000"
                           DataGridParos.Columns(5).Width = "700"
                           
                           DataGridParos.Columns(5).Caption = "Horas"
                           
                           DataGridParos.Columns(5).Alignment = dbgRight
                           
                           DataGridParos.Columns(5).NumberFormat = "#,###,##0.00"
        'End If

            
            
            If Err <> 0 Then
                MsgBox Err.Description
            End If
            
MousePointer = 0
        
End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False

End Sub

Private Sub CmdSalida_Click()
            Unload Me
End Sub

Private Sub DataGridDetalle_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
        RDetalle.Sort = RDetalle.Fields(ColIndex).Name
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
                
        
End Sub

Private Sub DataGridParos_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
                    RDetalle2.Sort = RDetalle2.Fields(ColIndex).Name
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
        

End Sub

Private Sub DBGridBusqueda_DblClick()
            If BLinea = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(0)
            ElseIf BGrupo = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(2)
            End If
            FrameBusqueda.Visible = False
            TxtLinea.SetFocus
End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                If BLinea = True Then
                    TxtLinea.Text = DBGridBusqueda.Columns(0)
                ElseIf BGrupo = True Then
                    TxtLinea.Text = DBGridBusqueda.Columns(2)
                End If
                FrameBusqueda.Visible = False
                TxtLinea.SetFocus
            End If
End Sub

Private Sub DbGridDia_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
        RDia.Sort = RDia.Fields(ColIndex).Name
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
End Sub

Private Sub DbGridGerencia_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
        RFicha.Sort = RFicha.Fields(ColIndex).Name
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
        

End Sub

Private Sub DBGridLineasFichaTecnica_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
        RLinea.Sort = RLinea.Fields(ColIndex).Name
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
        
        
End Sub

Private Sub DbGridMes_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
        RTarimas.Sort = RTarimas.Fields(ColIndex).Name
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
End Sub

Private Sub dbgridparos_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
            RParos.Sort = RParos.Fields(ColIndex).Name
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
            
End Sub

Private Sub Form_Load()
            DtpFecIni.Value = Date
            DTPFecFin.Value = Date
            CmdGenera_Click
End Sub

Private Sub Form_Resize()
On Error Resume Next
            
            
            DbgridMes.Height = Me.ScaleHeight - 4200
            DbGridParos.Height = Me.ScaleHeight - 1600
            DbGridGerencia.Height = Me.ScaleHeight - 1600
            DataGridDetalle.Height = Me.ScaleHeight - 1600
            DataGridParos.Height = Me.ScaleHeight - 1600
            DbGridDia.Height = Me.ScaleHeight - 1600
            
            SSTab1.Height = Me.ScaleHeight - 1000
            
            
            'MskTotalTarimas.Move 10000, Me.Height - 1000
            'MskTotalEnvases.Move 10000, Me.Height - 700
            'Label2.Item(0).Move 8800, Me.Height - 1000
            'Label2.Item(1).Move 8800, Me.Height - 700
            If Err <> 0 Then
            End If
End Sub

Private Sub OptGrupo_Click()
            LblDescripcion.Caption = "Grupo"
            TxtLinea.Visible = True
            TxtLinea.SetFocus
End Sub

Private Sub OptLinea_Click()
            LblDescripcion.Caption = "Linea"
            TxtLinea.Visible = True
            TxtLinea.SetFocus
End Sub



Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess " Then
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas where Descrip Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess " Then
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas where Linea Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
                
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBusqueda_GotFocus()
        TxtBusqueda.SelStart = 0
        TxtBusqueda.SelLength = Len(TxtBusqueda.Text)
End Sub

Private Sub TxtBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "tab"
        End If
End Sub

Private Sub TxtLinea_Change()
        If OptLinea.Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLinea.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                End If
                    If RBuscaLinea.RecordCount > 0 Then
                        LblLinea.Caption = RBuscaLinea!Descrip
                    Else
                        LblLinea.Caption = ""
                    End If
        End If
            
End Sub

Private Sub TxtLinea_DblClick()
            Set RBusqueda = New ADODB.Recordset
            If OptGrupo.Value = True Then
                    BGrupo = True
                    BLinea = False
            ElseIf OptLinea.Value = True Then
                    BGrupo = False
                    BLinea = True
            End If
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
End Sub

Private Sub TxtLinea_GotFocus()
        TxtLinea.SelStart = 0
        TxtLinea.SelLength = Len(TxtLinea.Text)
End Sub

Private Sub TxtLinea_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            Set RBusqueda = New ADODB.Recordset
            If OptGrupo.Value = True Then
                    BGrupo = True
                    BLinea = False
            ElseIf OptLinea.Value = True Then
                    BGrupo = False
                    BLinea = True
            End If
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If
End Sub
