VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CatalogosEspecificaciones 
   BackColor       =   &H000000FF&
   Caption         =   "Catalogos De Especificaciones De Rutinas"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "CatalogosEspecificaciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
      Caption         =   "Buscar Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   120
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbGridBuscar 
         Height          =   6735
         Left            =   120
         TabIndex        =   37
         Top             =   1200
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11880
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
         Height          =   615
         Left            =   10800
         Picture         =   "CatalogosEspecificaciones.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Sale de Lista"
         Top             =   480
         Width           =   735
      End
   End
   Begin MSDataGridLib.DataGrid DBGridDetalleCatalogos 
      Height          =   3735
      Left            =   240
      TabIndex        =   43
      ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
      Top             =   3240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6588
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
      ColumnCount     =   7
      BeginProperty Column00 
         DataField       =   "Rutina"
         Caption         =   "Rutina"
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
         DataField       =   "Cabezales"
         Caption         =   "Cabezales"
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
         DataField       =   "MinimoClienteMilimetros"
         Caption         =   "Minimo Cliente"
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
         DataField       =   "MinimoInternoMilimetros"
         Caption         =   "Minimo Interno"
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
         DataField       =   "MaximoInternoMilimetros"
         Caption         =   "Maximo Interno"
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
         DataField       =   "MaximoClienteMilimetros"
         Caption         =   "Maximo Cliente"
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
            ColumnWidth     =   870.236
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3165.166
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   884.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   1379.906
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   1289.764
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1409.953
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1440
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   240
      MouseIcon       =   "CatalogosEspecificaciones.frx":237C
      Picture         =   "CatalogosEspecificaciones.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Primer Registro"
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   600
      MouseIcon       =   "CatalogosEspecificaciones.frx":2CF0
      Picture         =   "CatalogosEspecificaciones.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Registro Anterior"
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   10920
      MouseIcon       =   "CatalogosEspecificaciones.frx":3664
      Picture         =   "CatalogosEspecificaciones.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Siguiente Registro"
      Top             =   7320
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   11280
      MouseIcon       =   "CatalogosEspecificaciones.frx":3FD8
      Picture         =   "CatalogosEspecificaciones.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Ultimo Registro"
      Top             =   7320
      Width           =   375
   End
   Begin VB.Frame FrameEncabezado 
      Caption         =   "Encabezado De Catalogo"
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
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.Frame FrameRequisiciones 
         Enabled         =   0   'False
         Height          =   615
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11535
         Begin VB.TextBox TxtDes 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   4200
            MaxLength       =   50
            TabIndex        =   3
            Top             =   240
            Width           =   7215
         End
         Begin VB.TextBox TxtCodCat 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   2
            ToolTipText     =   "doble click o signo '+' para ver catalogos"
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label Label4 
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
            Index           =   1
            Left            =   3120
            TabIndex        =   30
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Codigo Catalogo"
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
            TabIndex        =   29
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   855
         Left            =   8760
         Picture         =   "CatalogosEspecificaciones.frx":494C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   960
         Width           =   1400
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&EDITAR"
         Height          =   855
         Left            =   1560
         Picture         =   "CatalogosEspecificaciones.frx":4E86
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   960
         Width           =   1400
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&USCAR"
         Height          =   855
         Left            =   7320
         Picture         =   "CatalogosEspecificaciones.frx":525D
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   960
         Width           =   1400
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Caption         =   "&Salida"
         Height          =   855
         Left            =   10200
         Picture         =   "CatalogosEspecificaciones.frx":56E5
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Salida"
         Top             =   960
         Width           =   1400
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&BORRAR"
         Height          =   855
         Left            =   5880
         Picture         =   "CatalogosEspecificaciones.frx":5C00
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   960
         Width           =   1400
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&CANCELAR"
         Enabled         =   0   'False
         Height          =   855
         Left            =   4440
         Picture         =   "CatalogosEspecificaciones.frx":61C8
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   1400
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&GRABAR"
         Enabled         =   0   'False
         Height          =   855
         Left            =   3000
         Picture         =   "CatalogosEspecificaciones.frx":66FF
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   1400
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&AGREGAR"
         Height          =   855
         Left            =   120
         Picture         =   "CatalogosEspecificaciones.frx":6C5B
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   960
         Width           =   1400
      End
   End
   Begin VB.Frame FrameDetalle 
      Caption         =   "Detalle De Catalogo"
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
      Height          =   6015
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   11805
      Begin VB.CommandButton CmdEditar2 
         Caption         =   "Editar"
         Height          =   855
         Left            =   2880
         Picture         =   "CatalogosEspecificaciones.frx":6FD8
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton CmdBorrar2 
         Caption         =   "B&orrar"
         Height          =   855
         Left            =   7560
         Picture         =   "CatalogosEspecificaciones.frx":73AF
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton CmdCancelar2 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   855
         Left            =   6000
         Picture         =   "CatalogosEspecificaciones.frx":7977
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
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
         Height          =   855
         Left            =   9120
         Picture         =   "CatalogosEspecificaciones.frx":7EAE
         TabIndex        =   25
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton CmdGrabar2 
         Caption         =   "G&rabar"
         Enabled         =   0   'False
         Height          =   855
         Left            =   4440
         Picture         =   "CatalogosEspecificaciones.frx":83E0
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.CommandButton CmdAgregar2 
         Caption         =   "A&gregar"
         Height          =   855
         Left            =   1320
         Picture         =   "CatalogosEspecificaciones.frx":893C
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   5040
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Frame FrameDetalleRequisiciones 
         Enabled         =   0   'False
         Height          =   975
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   11535
         Begin VB.TextBox TxtCab 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   10440
            MaxLength       =   4
            TabIndex        =   16
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   240
            Width           =   1000
         End
         Begin MSMask.MaskEdBox MskMilimetros 
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   17
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtCodCatDet 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox TxtRut 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            MaxLength       =   4
            TabIndex        =   15
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   240
            Width           =   1000
         End
         Begin MSMask.MaskEdBox MskMilimetros 
            Height          =   285
            Index           =   1
            Left            =   5160
            TabIndex        =   18
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMilimetros 
            Height          =   285
            Index           =   2
            Left            =   7800
            TabIndex        =   19
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskMilimetros 
            Height          =   285
            Index           =   3
            Left            =   10440
            TabIndex        =   20
            Top             =   600
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cabezales"
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
            Left            =   9000
            TabIndex        =   44
            Top             =   240
            Width           =   885
         End
         Begin VB.Label LblRutina 
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
            Left            =   3720
            TabIndex        =   36
            Top             =   240
            Width           =   5055
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "MILIMETROS"
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
            Index           =   8
            Left            =   120
            TabIndex        =   35
            Top             =   600
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Maximo Interno"
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
            Left            =   6360
            TabIndex        =   34
            Top             =   600
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Maximo Cliente"
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
            Left            =   9000
            TabIndex        =   33
            Top             =   600
            Width           =   1290
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Minimo Interno"
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
            Left            =   3720
            TabIndex        =   32
            Top             =   600
            Width           =   1260
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Minimo Cliente"
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
            Left            =   1320
            TabIndex        =   31
            Top             =   600
            Width           =   1245
         End
         Begin VB.Label Label1 
            Caption         =   "Rutina"
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
            Left            =   1320
            TabIndex        =   28
            Top             =   240
            Width           =   735
         End
      End
   End
End
Attribute VB_Name = "CatalogosEspecificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VCatalogo As String
Dim vtexto As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean

Dim BCatalogos As Boolean
Dim BRutinas As Boolean

Dim RBuscaDetalle As New ADODB.Recordset
Dim RBuscaRutina As New ADODB.Recordset
Dim REncabezado As New ADODB.Recordset
Dim RDetalle As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VLlaveEncabezado As String
Dim VLlaveDetalle As String
Dim BEditarEncabezado As Boolean
Dim BEditarDetalle As Boolean


Sub Botones1()
    If Bandera = True Then
         FrameRequisiciones.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
    Else
         FrameRequisiciones.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleRequisiciones.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdEditar2.Enabled = False
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalleRequisiciones.Enabled = False
         CmdAgregar2.Enabled = True
         CmdEditar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
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
        
    Bandera2 = True
    Botones2
    Limpia_CamposDetalle

    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DBGridDetalleCatalogos.Enabled = False
    
    'SE ASIGNA AL DOCUMENTO DE DETALLE EL DOCUMENTO DEL ENCABEZADO
    BEditarDetalle = False
    TxtRut.Enabled = True
    TxtCodCatDet.Text = VCatalogo
    TxtRut.SetFocus
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        REncabezado.Delete
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147217873 Then
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                        
                        'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        REncabezado.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        REncabezado.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_CamposEncabezado
                        
                         Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where V.Codigo = '" & TxtCodCat.Text & "' And V.Rutina = R.Rutina")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where UPPER(V.Codigo) = '" & UCase(TxtCodCat.Text) & "' And V.Rutina = R.Rutina")
                                End If
                                    Llena_CamposDetalle
                                    Set DBGridDetalleCatalogos.DataSource = RDetalle
                    End If

End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Delete From VariablesMedia Where Codigo = '" & VCatalogo & "' And Rutina = '" & TxtRut.Text & "'"
                        Else 'ORACLE
                            Conexion.Execute "Delete From VariablesMedia Where UPPER(Codigo) = '" & UCase(VCatalogo) & "' And UPPER(Rutina) = '" & UCase(TxtRut.Text) & "'"
                        End If
                        
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
                        RDetalle.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RDetalle.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_CamposDetalle
                         
                    End If
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
On Error Resume Next
MousePointer = 11
    If Index = 1 Then
        REncabezado.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        REncabezado.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        REncabezado.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        REncabezado.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If REncabezado.BOF Then
        REncabezado.MoveFirst
    ElseIf REncabezado.EOF Then
        REncabezado.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_CamposEncabezado
    
    Set RDetalle = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where V.Codigo = '" & TxtCodCat.Text & "' And V.Rutina = R.Rutina")
            Else 'ORACLE
                Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where UPPER(V.Codigo) = '" & UCase(TxtCodCat.Text) & "' And V.Rutina = R.Rutina")
            End If
                Llena_CamposDetalle
                Set DBGridDetalleCatalogos.DataSource = RDetalle
                
    
    
MousePointer = 0


End Sub

Private Sub CmdBuscar_Click()
    On Error Resume Next
    mensaje = InputBox("Catalogo a Buscar")
    If mensaje <> "" Then
                
                REncabezado.MoveFirst
                    If GOrigenDeDatos = "AmaproAccess" Then
                        REncabezado.Find "CodigoVariable = '" & mensaje & "'"
                    Else
                        REncabezado.Find "CodigoVariable = '" & UCase(mensaje) & "'"
                    End If
                
                                
                
                Llena_CamposEncabezado
                
                Set RDetalle = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where V.Codigo = '" & TxtCodCat.Text & "' And V.Rutina = R.Rutina")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where UPPER(V.Codigo) = '" & UCase(TxtCodCat.Text) & "' And V.Rutina = R.Rutina")
                    End If
                        Llena_CamposDetalle
                        Set DBGridDetalleCatalogos.DataSource = RDetalle

    End If

End Sub

Private Sub CmdCancelar_Click()
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    Llena_CamposEncabezado
    FrameDetalle.Visible = True
    DBGridDetalleCatalogos.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
    
    DBGridDetalleCatalogos.Enabled = True
    Bandera2 = False
    Botones2
    Llena_CamposDetalle

End Sub

Private Sub CmdEditar_Click()
    BEditarEncabezado = True
    Bandera = True
    Botones1
    FrameDetalle.Visible = False
    DBGridDetalleCatalogos.Visible = False
    TxtCodCat.Enabled = False
End Sub


Private Sub CmdEditar2_Click()
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DBGridDetalleCatalogos.Enabled = False
    
    BEditarDetalle = True
    VLlaveDetalle = TxtRut.Text
    TxtRut.Enabled = False
    Bandera2 = True
    Botones2
    
End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
                   
        
    'REVISAMOS DATOS EN MILIMETROS
    If Not IsNumeric(MskMilimetros.Item(0).Text) Then
       MsgBox "Minimo Cliente No Es Numerica", vbOKOnly + vbCritical, "Error"
       MskMilimetros.Item(0).SetFocus
       Exit Sub
    End If
    
    'REVISAMOS DATOS EN MILIMETROS
    If Not IsNumeric(MskMilimetros.Item(1).Text) Then
       MsgBox "Minimo Interno No Es Numerica", vbOKOnly + vbCritical, "Error"
       MskMilimetros.Item(1).SetFocus
       Exit Sub
    End If
    
    'REVISAMOS DATOS EN MILIMETROS
    If Not IsNumeric(MskMilimetros.Item(2).Text) Then
       MsgBox "Maximo Interno No Es Numerica", vbOKOnly + vbCritical, "Error"
       MskMilimetros.Item(2).SetFocus
       Exit Sub
    End If
    
    'REVISAMOS DATOS EN MILIMETROS
    If Not IsNumeric(MskMilimetros.Item(3).Text) Then
       MsgBox "Maximo Cliente No Es Numerica", vbOKOnly + vbCritical, "Error"
       MskMilimetros.Item(3).SetFocus
       Exit Sub
    End If
    
    'CABEZALES
    If Not IsNumeric(TxtCab.Text) Then
       MsgBox "Cabezales Debe Ser Numerico", vbOKOnly + vbCritical, "Error"
       TxtCab.SetFocus
       Exit Sub
    End If
            
                'AGREGAR
                    If BEditarDetalle = False Then
                            vtexto = "Values('" & TxtCodCatDet.Text & "', " ' CODIGO
                            vtexto = vtexto & MskMilimetros.Item(0).Text & ", " 'MINIMO CLIENTE
                            vtexto = vtexto & MskMilimetros.Item(1).Text & ", " 'MINIMO INTERNO
                            vtexto = vtexto & MskMilimetros.Item(2).Text & ", " 'MAXIMO CLIENTE
                            vtexto = vtexto & MskMilimetros.Item(3).Text & ", '" 'MAXIMO INTERNO
                            vtexto = vtexto & TxtRut.Text & "', " 'RUTINA
                            vtexto = vtexto & TxtCab.Text & ")" 'CABEZALES
                            
                            Conexion.Execute "Insert Into VariablesMedia " & vtexto
                'EDITAR
                    Else
                            vtexto = "MinimoClienteMilimetros = " & MskMilimetros.Item(0).Text & ", " 'MINIMO CLIENTE
                            vtexto = vtexto & "MinimoInternoMilimetros = " & MskMilimetros.Item(1).Text & ", " 'MINIMO INTERNO
                            vtexto = vtexto & "MaximoInternoMilimetros = " & MskMilimetros.Item(2).Text & ", " 'MAXIMO INTERNNO
                            vtexto = vtexto & "MaximoClienteMilimetros = " & MskMilimetros.Item(3).Text & ", " 'MAXIMO CLIENTE
                            vtexto = vtexto & "Cabezales = " & TxtCab.Text 'MAXIMO CLIENTE
                            vtexto = vtexto & " Where Codigo = '" & VCatalogo & "' And Rutina = '" & VLlaveDetalle & "'"
                        
                            Conexion.Execute "UPDATE VariablesMedia SET " & vtexto
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
                            MsgBox "Catalogo y Rutina Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtRut.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera2 = False
                        Botones2
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RDetalle.Requery
                        RDetalle.MoveLast
                        Llena_CamposDetalle
    TxtRut.Enabled = True
    DBGridDetalleCatalogos.Enabled = True
    CmdAgregar2.SetFocus
End Sub


Private Sub CmdAgregar_Click()

    TxtCodCat.Enabled = True
    Bandera = True
    Botones1
    FrameDetalle.Visible = False
    DBGridDetalleCatalogos.Visible = False
    Limpia_CamposEncabezado
    
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    
    VCatalogo = TxtCodCat.Text
    
    
                    'AGREGAR
                    If BEditarEncabezado = False Then
                            vtexto = "Values('" & TxtCodCat.Text & "', '" ' CODIGO
                            vtexto = vtexto & TxtDes.Text & "', '" 'DESCRIPCION
                            vtexto = vtexto & GUsuario & "')" 'USUARIO
                            
                            Conexion.Execute "Insert Into VariablesDescripcion " & vtexto
                    'EDITAR
                    Else
                            vtexto = "DescripcionVariable = '" & TxtDes.Text & "', " 'DESCRIPCION
                            vtexto = vtexto & "usuario = '" & GUsuario & "' " ' USUARIO
                            vtexto = vtexto & "Where CodigoVariable = '" & TxtCodCat.Text & "'"
                        
                            Conexion.Execute "UPDATE VariablesDescripcion SET " & vtexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Codigo De Catalogo Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtCodCat.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo De Catalogo Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtCodCat.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        Botones1
                        TxtCodCat.Enabled = True
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        REncabezado.Requery
                        'REncabezado.MoveLast
                        
            'Set REncabezado = New ADODB.Recordset
            '            Call Abrir_Recordset(REncabezado, "Select * From VariablesDescripcion") 'Where CodigoVariable = '" & VCatalogo & "'")
            '
            '            REncabezado.MoveFirst
            '            If GOrigenDeDatos = "AmaproAccess" Then
            '                REncabezado.Find "CodigoVariable = '" & VCatalogo & "'"
            '            Else
            '                REncabezado.Find "CodigoVariable = '" & UCase(VCatalogo) & "'"
            '            End If
            '
            '        Llena_CamposEncabezado
   
            Set RDetalle = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where V.Codigo = '" & VCatalogo & "' And V.Rutina = R.Rutina")
                Else 'ORACLE
                    Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where UPPER(V.Codigo) = '" & UCase(VCatalogo) & "' And V.Rutina = R.Rutina")
                End If
                Llena_CamposDetalle
                Set DBGridDetalleCatalogos.DataSource = RDetalle
    
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    'BOTONES DE DATA
    CmdBotones2.Item(1).Visible = False
    CmdBotones2.Item(2).Visible = False
    CmdBotones2.Item(3).Visible = False
    CmdBotones2.Item(4).Visible = False
    FrameDetalle.Visible = True
    DBGridDetalleCatalogos.Visible = True
    FrameDetalle.Enabled = True
    FrameEncabezado.Enabled = False
    CmdAgregar2.SetFocus
End Sub

Private Sub CmdImprimir_Click()
MousePointer = 11
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "Variables.rpt"
                Else
                    GNombreReporte = "VariablesO.rpt"
                End If
                GCriteriaReporte = "{VariablesDescripcion.CodigoVariable} = '" & TxtCodCat.Text & "'"
                FrmReporte.Show
            
MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdTerminar_Click()
If CmdCancelar2.Enabled = True Then
     CmdCancelar2_Click
End If
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    CmdBotones2.Item(1).Visible = True
    CmdBotones2.Item(2).Visible = True
    CmdBotones2.Item(3).Visible = True
    CmdBotones2.Item(4).Visible = True
    
    FrameDetalle.Visible = True
    FrameEncabezado.Enabled = True
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle

End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub

Private Sub DBGridBuscar_DblClick()
            
        If BRutinas = True Then
            TxtRut.Text = DbGridBuscar.Columns(0)
            TxtRut.SetFocus
        ElseIf BCatalogos = True Then
            TxtCodCat.Text = DbGridBuscar.Columns(0)
            TxtCodCat.SetFocus
        End If
           
            FrameBuscar.Visible = False
End Sub

Private Sub DbGridBuscar_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                If BRutinas = True Then
                    TxtRut.Text = DbGridBuscar.Columns(0)
                    TxtRut.SetFocus
                ElseIf BCatalogos = True Then
                    TxtCodCat.Text = DbGridBuscar.Columns(0)
                    TxtCodCat.SetFocus
                End If
                    
                    FrameBuscar.Visible = False
                
        End If
End Sub

Private Sub DBGridDetalleCatalogos_HeadClick(ByVal ColIndex As Integer)
        RDetalle.Sort = RDetalle.Fields(ColIndex).Name
End Sub


Private Sub DBGridDetalleCatalogos_SelChange(Cancel As Integer)
        Llena_CamposDetalle
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
    

End Sub

Private Sub Form_Load()
        Set REncabezado = New ADODB.Recordset
            Call Abrir_Recordset(REncabezado, "Select * From VariablesDescripcion")
                Llena_CamposEncabezado
                
        Set RDetalle = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where V.Codigo = '" & TxtCodCat.Text & "' And V.Rutina = R.Rutina")
            Else 'ORACLE
                Call Abrir_Recordset(RDetalle, "Select V.Rutina, R.Descrip, V.Cabezales, V.MinimoClienteMilimetros, V.MinimoInternoMilimetros, V.MaximoInternoMilimetros, V.MaximoClienteMilimetros From VariablesMedia V, Rutinas R Where UPPER(V.Codigo) = '" & UCase(TxtCodCat.Text) & "' And V.Rutina = R.Rutina")
            End If
                Llena_CamposDetalle
                Set DBGridDetalleCatalogos.DataSource = RDetalle
End Sub

Private Sub MskMilimetros_GotFocus(Index As Integer)
        MskMilimetros.Item(Index).SelStart = 0
        MskMilimetros.Item(Index).SelLength = Len(MskMilimetros.Item(Index).Text)
End Sub

Private Sub MskMilimetros_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub



Private Sub TxtCab_GotFocus()
        TxtCab.SelStart = 0
        TxtCab.SelLength = Len(TxtCab.Text)
End Sub

Private Sub TxtCab_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtCodCat_DblClick()
            BRutinas = False
            BCatalogos = True
            FrameBuscar.Visible = True
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select * From VariablesDescripcion")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(0).Width = 1500
            DbGridBuscar.Columns(1).Width = 4000
            DbGridBuscar.SetFocus

End Sub

Private Sub txtDes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub


Private Sub TxtCodCat_GotFocus()
        TxtCodCat.SelStart = 0
        TxtCodCat.SelLength = Len(TxtCodCat.Text)
End Sub

Private Sub TxtCodCat_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BRutinas = False
            BCatalogos = True
            FrameBuscar.Visible = True
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select * From VariablesDescripcion")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(0).Width = 1500
            DbGridBuscar.Columns(1).Width = 4000
            DbGridBuscar.SetFocus
        End If
End Sub

Private Sub TxtDes_GotFocus()
        TxtDes.SelStart = 0
        TxtDes.SelLength = Len(TxtDes.Text)
End Sub

Private Sub TxtRut_Change()
        Set RBuscaRutina = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaRutina, "Select Descrip From Rutinas Where Rutina = '" & TxtRut.Text & "'")
            Else ' ORACLE
                Call Abrir_Recordset(RBuscaRutina, "Select Descrip From Rutinas Where UPPER(Rutina) = '" & UCase(TxtRut.Text) & "'")
            End If
            If RBuscaRutina.RecordCount > 0 Then
                LblRutina.Caption = RBuscaRutina!Descrip
            Else
                LblRutina.Caption = ""
            End If
End Sub

Private Sub TxtRut_DblClick()
            BRutinas = True
            BCatalogos = False
            FrameBuscar.Visible = True
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Rutina, Descrip From Rutinas")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(0).Width = 1500
            DbGridBuscar.Columns(1).Width = 4000
            DbGridBuscar.SetFocus
End Sub

Private Sub TxtRut_GotFocus()
        TxtRut.SelStart = 0
        TxtRut.SelLength = Len(TxtRut.Text)
End Sub

Private Sub TxtRut_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BRutinas = True
            BCatalogos = False
            FrameBuscar.Visible = True
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Rutina, Descrip From Rutinas")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(0).Width = 1500
            DbGridBuscar.Columns(1).Width = 4000
            DbGridBuscar.SetFocus
        End If
End Sub



Public Sub Llena_CamposEncabezado()
On Error Resume Next
            If REncabezado.RecordCount > 0 Then
                TxtCodCat.Text = REncabezado!CodigoVariable
                TxtDes.Text = REncabezado!DescripcionVariable
            Else
                TxtCodCat.Text = ""
                TxtDes.Text = ""
            End If
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Llena_CamposDetalle()
On Error Resume Next
            If RDetalle.RecordCount > 0 Then
                TxtCodCatDet.Text = RDetalle!Codigo
                MskMilimetros.Item(0).Text = RDetalle!MinimoClienteMilimetros
                MskMilimetros.Item(1).Text = RDetalle!MinimoInternoMilimetros
                MskMilimetros.Item(2).Text = RDetalle!MaximoInternoMilimetros
                MskMilimetros.Item(3).Text = RDetalle!MaximoClienteMilimetros
                TxtRut.Text = RDetalle!Rutina
                TxtCab.Text = RDetalle!cabezales
            Else
                Limpia_CamposDetalle
            End If
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Limpia_CamposEncabezado()
                TxtCodCat.Text = ""
                TxtDes.Text = ""
End Sub

Public Sub Limpia_CamposDetalle()
                TxtCodCatDet.Text = ""
                MskMilimetros.Item(0).Text = "0"
                MskMilimetros.Item(1).Text = "0"
                MskMilimetros.Item(2).Text = "0"
                MskMilimetros.Item(3).Text = "0"
                TxtRut.Text = ""
                TxtCab.Text = "0"
End Sub

