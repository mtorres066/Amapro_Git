VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CapturaDesperdicioMateriaPrima 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Desperdicio De Materia Prima"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9135
   Icon            =   "CapturaDesperdicioMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Caption         =   "Busqueda de Datos"
      Height          =   5895
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   9135
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4695
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   8281
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
      Begin VB.TextBox TxtBus 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         Top             =   720
         Width           =   6735
      End
      Begin VB.OptionButton OptBus 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   32
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptBus 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   8160
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Sale de Busqueda"
         Top             =   240
         Width           =   735
      End
      Begin VB.Label LblBus 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   720
         Width           =   975
      End
   End
   Begin TabDlg.SSTab TabDesperdicio 
      Height          =   6015
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   10610
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CapturaDesperdicioMateriaPrima.frx":3D6C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameDesperdicio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdBotones(5)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdBotones(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdBotones(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdBotones(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdBotones(1)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdBotones(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmdBotones2(4)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CmdBotones2(3)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "CmdBotones2(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "CmdBotones2(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaDesperdicioMateriaPrima.frx":4086
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DGridDesperdicio"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda"
      TabPicture(2)   =   "CapturaDesperdicioMateriaPrima.frx":44D8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusquedadeDatos"
      Tab(2).ControlCount=   1
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   1
         Left            =   480
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":492A
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":4D6C
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Primer Registro"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   2
         Left            =   840
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":529E
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":56E0
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Registro Anterior"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   3
         Left            =   7800
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":5C12
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":6054
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Siguiente Registro"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   4
         Left            =   8160
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":6586
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":69C8
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Ultimo Registro"
         Top             =   4920
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Agregar"
         Height          =   800
         Index           =   0
         Left            =   1320
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":6EFA
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":733C
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4800
         Width           =   1000
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Editar"
         Height          =   800
         Index           =   1
         Left            =   2400
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":786E
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":7CB0
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4800
         Width           =   1000
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   800
         Index           =   2
         Left            =   3480
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":81E2
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":8624
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4800
         Width           =   1000
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   800
         Index           =   3
         Left            =   4560
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":8B56
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":8F98
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4800
         Width           =   1000
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "B&orrar"
         Height          =   800
         Index           =   4
         Left            =   5640
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":94CA
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":990C
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4800
         Width           =   1000
      End
      Begin VB.CommandButton CmdBotones 
         Caption         =   "&Salida"
         Height          =   800
         Index           =   5
         Left            =   6720
         MouseIcon       =   "CapturaDesperdicioMateriaPrima.frx":9E3E
         Picture         =   "CapturaDesperdicioMateriaPrima.frx":A280
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   4800
         Width           =   1000
      End
      Begin MSDataGridLib.DataGrid DGridDesperdicio 
         Height          =   5175
         Left            =   -74880
         TabIndex        =   37
         Top             =   720
         Width           =   8895
         _ExtentX        =   15690
         _ExtentY        =   9128
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
         ColumnCount     =   8
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
         BeginProperty Column03 
            DataField       =   "CodigoProceso"
            Caption         =   "Proceso"
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
         BeginProperty Column05 
            DataField       =   "CuerposProceso"
            Caption         =   "Proceso"
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
            DataField       =   "CuerposProveedor"
            Caption         =   "Proveedor"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   404.787
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1065.26
            EndProperty
         EndProperty
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
         Height          =   4575
         Left            =   -74880
         TabIndex        =   15
         Top             =   720
         Width           =   8775
         Begin VB.CommandButton CmdBotones 
            Caption         =   "Seleccionar Todos"
            Height          =   732
            Index           =   7
            Left            =   6840
            Picture         =   "CapturaDesperdicioMateriaPrima.frx":C2F2
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   3600
            Width           =   1812
         End
         Begin VB.CommandButton CmdBotones 
            Caption         =   "Seleccionar Datos"
            Height          =   732
            Index           =   6
            Left            =   6840
            Picture         =   "CapturaDesperdicioMateriaPrima.frx":C5FC
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   2760
            Width           =   1812
         End
         Begin MSComCtl2.DTPicker DtpFecFin 
            Height          =   255
            Left            =   7200
            TabIndex        =   11
            Top             =   1800
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   61603843
            CurrentDate     =   37396
         End
         Begin MSComCtl2.DTPicker DtpFecIni 
            Height          =   255
            Left            =   7200
            TabIndex        =   10
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   61603843
            CurrentDate     =   37396
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            TabIndex        =   12
            Top             =   2400
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha Y Proceso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1215
            Index           =   1
            Left            =   1800
            Picture         =   "CapturaDesperdicioMateriaPrima.frx":E2F6
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   360
            Width           =   1335
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
            Height          =   1215
            Index           =   0
            Left            =   360
            Picture         =   "CapturaDesperdicioMateriaPrima.frx":E600
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.Label Label1 
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
            Index           =   1
            Left            =   6480
            TabIndex        =   26
            Top             =   1800
            Width           =   510
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
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
            Left            =   6480
            TabIndex        =   25
            Top             =   1440
            Width           =   555
         End
         Begin VB.Label LblBusqueda 
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
            Left            =   3240
            TabIndex        =   16
            Top             =   2400
            Width           =   2775
         End
      End
      Begin VB.Frame FrameDesperdicio 
         Caption         =   "Datos del Desperdicio"
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
         Height          =   3255
         Left            =   240
         TabIndex        =   0
         Top             =   1320
         Width           =   8655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1440
            MaxLength       =   1
            TabIndex        =   2
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000B&
            Height          =   285
            Index           =   8
            Left            =   1680
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   2760
            Width           =   1200
         End
         Begin MSMask.MaskEdBox MskFec 
            Height          =   285
            Left            =   1440
            TabIndex        =   1
            Top             =   240
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   1680
            TabIndex        =   7
            Top             =   2400
            Width           =   1200
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1680
            TabIndex        =   6
            Top             =   2040
            Width           =   1185
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1440
            MaxLength       =   12
            TabIndex        =   5
            ToolTipText     =   "Doble click o signo '+' para ayuda"
            Top             =   1680
            Width           =   1440
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   4
            ToolTipText     =   "Doble click o signo '+' para ayuda"
            Top             =   1320
            Width           =   1440
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   3
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   30
            Top             =   2760
            Width           =   540
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
            Left            =   3000
            TabIndex        =   29
            Top             =   960
            Width           =   5415
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   450
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
            Left            =   3000
            TabIndex        =   14
            Top             =   1680
            Width           =   5415
         End
         Begin VB.Label LblProceso 
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
            Left            =   3000
            TabIndex        =   13
            Top             =   1320
            Width           =   5415
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Cantidad Proveedor"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   24
            Top             =   2400
            Width           =   1410
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            BackColor       =   &H80000004&
            Caption         =   "Cantidad Proceso"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   23
            Top             =   2040
            Width           =   1260
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   22
            Top             =   1680
            Width           =   495
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Proceso"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   21
            Top             =   1320
            Width           =   1125
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Linea"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   20
            Top             =   1020
            Width           =   390
         End
         Begin VB.Label lblFieldLabel 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   19
            Top             =   645
            Width           =   420
         End
      End
   End
End
Attribute VB_Name = "CapturaDesperdicioMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim VMensaje As Integer

Dim BLinea As Boolean
Dim BProceso As Boolean
Dim BCodigo As Boolean
Dim BEditar As Boolean

Dim RDesperdicio As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaProceso As New ADODB.Recordset
Dim RBuscaCodigo As New ADODB.Recordset


Dim VUltimaLinea As String
Dim VUltimaFichaTecnica As String
Dim VUltimoTurno As String
Dim VUltimaFecha As String
Dim VTexto As String


Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    
    
        'AGREGAR
        If Index = 0 Then
                Bandera = True
                botones
                Limpia_Campos
                
                'HABILITA LA LLAVE
                MskFec.Enabled = False
                TxtTexto.Item(0).Enabled = True
                TxtTexto.Item(1).Enabled = True
                TxtTexto.Item(2).Enabled = True
                TxtTexto.Item(3).Enabled = True
                TxtTexto.Item(4).Enabled = True
                
                MskFec.Text = VUltimaFecha
                TxtTexto.Item(8).Text = GUsuario
                TxtTexto.Item(0).Text = VUltimoTurno
                TxtTexto.Item(1).Text = VUltimaLinea
                TxtTexto.Item(3).Text = VUltimaFichaTecnica
                MskFec.SetFocus
                BEditar = False
        'EDITAR
        ElseIf Index = 1 Then
        
                Bandera = True
                botones
                'DESABILITA LA LLAVE
                MskFec.Enabled = False
                TxtTexto.Item(0).Enabled = False
                TxtTexto.Item(1).Enabled = False
                TxtTexto.Item(2).Enabled = False
                TxtTexto.Item(3).Enabled = False
                TxtTexto.Item(4).Enabled = False
                
                TxtTexto.Item(5).SetFocus
                TxtTexto.Item(8).Text = GUsuario
                BEditar = True
        'GRABAR
        ElseIf Index = 2 Then
                        
                        'REVISA LA FECHA
                        If Not IsDate(MskFec.Text) Then
                            MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            MskFec.SetFocus
                            Exit Sub
                        End If
                        
                        VUltimoTurno = TxtTexto.Item(0).Text
                        VUltimaLinea = TxtTexto.Item(1).Text
                        VUltimaFichaTecnica = TxtTexto.Item(3).Text
                        VUltimaFecha = MskFec.Text
                        
                    'AGREGAR
                    If BEditar = False Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "Values(#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = "Values(To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            
                            VTexto = VTexto & "'" & TxtTexto.Item(0).Text & "', '" 'TURNO
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', '" 'LINEA
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', '" 'PROCESO
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(4).Text & ", " 'PROCESO
                            VTexto = VTexto & TxtTexto.Item(5).Text & ", '" 'PROVEEDOR
                            VTexto = VTexto & TxtTexto.Item(8).Text & "')" 'USUARIO
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into BodegasInventario " & VTexto
                    'EDITAR
                    Else
                            VTexto = "CuerposProceso = " & TxtTexto.Item(4).Text & ", " 'PROCESO
                            VTexto = VTexto & "CuerposProveedor = " & TxtTexto.Item(5).Text & ", " 'PROVEEDOR
                            VTexto = VTexto & "usuario = '" & TxtTexto.Item(8).Text & "' " ' USUARIO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "Where Fecha = #" & Format(MskFec.Text, "mm/dd/yyyy") & "# And "
                            Else 'ORACLE
                                VTexto = VTexto & "Where Fecha = TO_DATE('" & MskFec.Text & "', 'dd/mm/yyyy') And "
                            End If
                                VTexto = VTexto & "Turno = '" & TxtTexto.Item(0) & "' And Linea = '" & TxtTexto.Item(1) & "' And "
                                VTexto = VTexto & "CodigoProceso = '" & TxtTexto.Item(2) & "' And FichaTecnica = '" & TxtTexto.Item(3) & "'"
                        
                            Conexion.Execute "UPDATE BodegasInventario SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = 3022 Then
                            MsgBox "Codigo De Bodega Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> 3022 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217900 Then
                            MsgBox "Codigo De Bodega Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217900 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
                        
                        'HABILITA LA LLAVE
                        MskFec.Enabled = False
                        TxtTexto.Item(0).Enabled = True
                        TxtTexto.Item(1).Enabled = True
                        TxtTexto.Item(2).Enabled = True
                        TxtTexto.Item(3).Enabled = True
                        TxtTexto.Item(4).Enabled = True
                        
                        
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RDesperdicio.Requery

        'CANCELAR
        ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    Llena_Campos
                    'HABILITA LA LLAVE
                    MskFec.Enabled = False
                    TxtTexto.Item(0).Enabled = True
                    TxtTexto.Item(1).Enabled = True
                    TxtTexto.Item(2).Enabled = True
                    TxtTexto.Item(3).Enabled = True
                    TxtTexto.Item(4).Enabled = True
        ElseIf Index = 4 Then ' BORRAR
        
                On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RDesperdicio.Delete
                        
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
                        RDesperdicio.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RDesperdicio.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If

        ElseIf Index = 5 Then ' SALIDA
                Unload Me
        ElseIf Index = 6 Then 'SELECCIONAR DATOS
                    Set RDesperdicio = New ADODB.Recordset
                    If OptBusqueda.Item(0).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RDesperdicio, "Select * From CapturaDesperdicio Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#")
                        Else 'ORACLE
                            Call Abrir_Recordset(RDesperdicio, "Select * From CapturaDesperdicio Where Fecha >= TO_DATE('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= TO_DATE('" & DtpFecFin.Value & "', 'dd/mm/yyyy')")
                        End If
                    ElseIf OptBusqueda.Item(1).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RDesperdicio, "Select * From CapturaDesperdicio Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And CodigoProceso Like '" & TxtBusqueda.Text & "%'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RDesperdicio, "Select * From CapturaDesperdicio Where Fecha >= TO_DATE('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= TO_DATE('" & DtpFecFin.Value & "', 'dd/mm/yyyy') And UPPER(CodigoProceso) Like '" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                    End If
                    TabDesperdicio.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    Call Abrir_Recordset(RDesperdicio, "Select * From CapturaDesperdicio")
                    TabDesperdicio.Tab = 1
        End If
    
    

End Sub


Sub botones()
    If Bandera = True Then
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         FrameDesperdicio.Enabled = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         DGridDesperdicio.Visible = False
         FrameBusquedadeDatos.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         FrameDesperdicio.Enabled = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True

         DGridDesperdicio.Visible = True
         FrameBusquedadeDatos.Visible = True
    End If
End Sub


Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RDesperdicio.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RDesperdicio.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RDesperdicio.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RDesperdicio.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RDesperdicio.BOF Then
        RDesperdicio.MoveFirst
    ElseIf RDesperdicio.EOF Then
        RDesperdicio.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub


Private Sub DBGridBusqueda_DblClick()
    If BLinea = True Then
        TxtTexto.Item(1).Text = DbGridBusqueda.Columns(0).Text
        TxtTexto.Item(1).SetFocus
    ElseIf BProceso = True Then
        TxtTexto.Item(2).Text = DbGridBusqueda.Columns(0).Text
        TxtTexto.Item(2).SetFocus
    ElseIf BCodigo = True Then
        TxtTexto.Item(3).Text = DbGridBusqueda.Columns(0).Text
        TxtTexto.Item(3).SetFocus
    End If
        FrameBusqueda.Visible = False
        
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
            If BLinea = True Then
                TxtTexto.Item(1).Text = DbGridBusqueda.Columns(0).Text
                TxtTexto.Item(1).SetFocus
            ElseIf BProceso = True Then
                TxtTexto.Item(2).Text = DbGridBusqueda.Columns(0).Text
                TxtTexto.Item(2).SetFocus
            ElseIf BCodigo = True Then
                TxtTexto.Item(3).Text = DbGridBusqueda.Columns(0).Text
                TxtTexto.Item(3).SetFocus
            End If
                FrameBusqueda.Visible = False
    End If
End Sub

Private Sub dgriddesperdicio_HeadClick(ByVal ColIndex As Integer)
        
        RDesperdicio.Sort = RDesperdicio.Fields(ColIndex).Name
        Set DGridDesperdicio.DataSource = RDesperdicio
        
        
End Sub

Private Sub Form_Load()
    Set RDesperdicio = New ADODB.Recordset
    Call Abrir_Recordset(RDesperdicio, "Select * From CapturaDesperdicio")
    
    Set DGridDesperdicio.DataSource = RDesperdicio
    Llena_Campos
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
            LblBusqueda.Caption = ""
            TxtBusqueda.Visible = False
    ElseIf Index = 1 Then
            LblBusqueda.Caption = "Codigo De Proceso"
            TxtBusqueda.Visible = True
            TxtBusqueda.SetFocus
    End If
            
End Sub

Private Sub TabDesperdicio_Click(PreviousTab As Integer)
        If TabDesperdicio.Tab = 2 Then
            DtpFecIni.Value = Date
            DtpFecFin.Value = Date
        End If
End Sub

Private Sub TxtBus_Change()
            
            
                    'OPCION POR DESCRIPCION
                    If OptBus.Item(0).Value = True Then
                                If BLinea = True Then
                                    DataBusqueda.RecordSource = ("Select Linea, Descrip from Lineas Where Descrip Like '*" & TxtBus.Text & "*'")
                                ElseIf BProceso = True Then
                                    DataBusqueda.RecordSource = ("Select CodigoProceso, Descripcion, Grupo from ProcesosMateriaPrima Where Descripcion Like '*" & TxtBus.Text & "*'")
                                ElseIf BCodigo = True Then
                                    DataBusqueda.RecordSource = ("Select CodigoMateriaPrima, Descripcion from CorrelativosMateriaPrima Where Descripcion Like '*" & TxtBus.Text & "*'")
                                End If
                                    
                    'OPCION DE CODIGO
                    ElseIf OptBus.Item(1).Value = True Then
                                If BLinea = True Then
                                    DataBusqueda.RecordSource = ("Select Linea, Descrip from Lineas Where Linea Like '*" & TxtBus.Text & "*'")
                                ElseIf BProceso = True Then
                                    DataBusqueda.RecordSource = ("Select CodigoProceso, Descripcion, Grupo from ProcesosMateriaPrima Where CodigoProceso Like '*" & TxtBus.Text & "*'")
                                ElseIf BCodigo = True Then
                                    DataBusqueda.RecordSource = ("Select CodigoMateriaPrima, Descripcion from CorrelativosMateriaPrima Where CodigoMateriaPrima Like '*" & TxtBus.Text & "*'")
                                End If
                        
                    End If
                            DataBusqueda.Refresh
                            DbGridBusqueda.Refresh
                            DbGridBusqueda.Columns(1).Width = "5000"

End Sub

Private Sub TxtBus_GotFocus()
        TxtBus.SelStart = 0
        TxtBus.SelLength = Len(TxtBus.Text)
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
        'LINEA
        If Index = 1 Then
            Set RBuscaLinea = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        'PROCESO
        ElseIf Index = 2 Then
            Set RBuscaProceso = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaProceso, "Select Descripcion From ProcesosMateriaPrima Where CodigoProceso = '" & TxtTexto.Item(2).Text & "'")
                If RBuscaProceso.RecordCount > 0 Then
                    LblProceso.Caption = RBuscaProceso!Descripcion
                Else
                    LblProceso.Caption = ""
                End If
        'CODIGO
        ElseIf Index = 3 Then
            Set RBuscaCodigo = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaCodigo, "Select Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(3).Text & "'")
                If RBuscaCodigo.RecordCount > 0 Then
                    LblFichaTecnica.Caption = RBuscaCodigo!Descripcion
                Else
                    LblFichaTecnica.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
    If Index = 1 Or Index = 2 Or Index = 3 Then
        Set RBusqueda = New ADODB.Recordset
    End If
    
    'LINEA
    If Index = 1 Then
        BLinea = True
        BProceso = False
        BCodigo = False
        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
    'PROCESOS
    ElseIf Index = 2 Then
        BLinea = False
        BProceso = True
        BCodigo = False
        Call Abrir_Recordset(RBusqueda, "Select CodigoProceso, Descripcion, Grupo From ProcesosMateriaPrima")
    'CODIGO
    ElseIf Index = 3 Then
        BLinea = False
        BProceso = False
        BCodigo = True
        Call Abrir_Recordset(RBusqueda, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima")
    End If
        
    If Index = 1 Or Index = 2 Or Index = 3 Then
        Set DbGridBusqueda.DataSource = RBusqueda
        FrameBusqueda.Visible = True
        TxtBus.SetFocus
        DbGridBusqueda.Columns(0).Width = "1000"
        DbGridBusqueda.Columns(1).Width = "4000"
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
                If Index = 1 Or Index = 2 Or Index = 3 Then
                    Set RBusqueda = New ADODB.Recordset
                End If
                
                'LINEA
                If Index = 1 Then
                    BLinea = True
                    BProceso = False
                    BCodigo = False
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
                'PROCESOS
                ElseIf Index = 2 Then
                    BLinea = False
                    BProceso = True
                    BCodigo = False
                    Call Abrir_Recordset(RBusqueda, "Select CodigoProceso, Descripcion, Grupo From ProcesosMateriaPrima")
                'CODIGO
                ElseIf Index = 3 Then
                    BLinea = False
                    BProceso = False
                    BCodigo = True
                    Call Abrir_Recordset(RBusqueda, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima")
                End If
                    
                If Index = 1 Or Index = 2 Or Index = 3 Then
                    Set DbGridBusqueda.DataSource = RBusqueda
                    FrameBusqueda.Visible = True
                    TxtBus.SetFocus
                    DbGridBusqueda.Columns(0).Width = "1000"
                    DbGridBusqueda.Columns(1).Width = "4000"
                End If
                    
                
        End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
        'FECHA
        MskFec.Text = RDesperdicio!fecha
        'TURNO
            If IsNull(RDesperdicio!Turno) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = RDesperdicio!Turno
            End If
        'LINEA
            If IsNull(RDesperdicio!Linea) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = RDesperdicio!Linea
            End If
        'PROCESO
            If IsNull(RDesperdicio!CodigoProceso) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RDesperdicio!CodigoProceso
            End If
        'FICHA TECNICA
            If IsNull(RDesperdicio!FichaTecnica) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = RDesperdicio!FichaTecnica
            End If
            
        TxtTexto.Item(4).Text = RDesperdicio!CuerposProceso
        TxtTexto.Item(5).Text = RDesperdicio!CuerposProveedor
        TxtTexto.Item(8).Text = RDesperdicio!usuario
        
        If Err <> 0 Then
            MsgBox Err.Description
        End If

End Sub

Public Sub Limpia_Campos()
        
        MskFec.Text = ""
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(3).Text = ""
        TxtTexto.Item(4).Text = ""
        TxtTexto.Item(5).Text = ""
        TxtTexto.Item(8).Text = ""
        
End Sub
