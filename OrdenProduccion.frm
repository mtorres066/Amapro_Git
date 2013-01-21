VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form OrdenProduccion 
   Caption         =   "Orden De Produccion"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "OrdenProduccion.frx":0000
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
      Height          =   7935
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6735
         Left            =   120
         TabIndex        =   19
         Top             =   960
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
         Left            =   10920
         Picture         =   "OrdenProduccion.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1680
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   6255
      End
   End
   Begin MSDataGridLib.DataGrid DbGridDetalle 
      Height          =   3255
      Left            =   240
      TabIndex        =   61
      Top             =   3960
      Width           =   11295
      _ExtentX        =   19923
      _ExtentY        =   5741
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
      ColumnCount     =   9
      BeginProperty Column00 
         DataField       =   "Documento"
         Caption         =   "Documento"
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
      BeginProperty Column02 
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
      BeginProperty Column03 
         DataField       =   "Pasada"
         Caption         =   "Pasada"
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
         DataField       =   "Requerido"
         Caption         =   "Requerido"
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
      BeginProperty Column05 
         DataField       =   "Entregado"
         Caption         =   "Entregado"
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
      BeginProperty Column06 
         DataField       =   "Saldo"
         Caption         =   "Saldo"
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
      BeginProperty Column07 
         DataField       =   "Desperdicio"
         Caption         =   "Desperdicio"
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
      BeginProperty Column08 
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
            ColumnWidth     =   915.024
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   3014.929
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   615.118
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnWidth     =   1019.906
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnWidth     =   1005.165
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnWidth     =   945.071
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1379.906
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "OrdenProduccion.frx":24B4
      Picture         =   "OrdenProduccion.frx":28F6
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Primer Registro"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "OrdenProduccion.frx":2E28
      Picture         =   "OrdenProduccion.frx":326A
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Registro Anterior"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   10920
      MouseIcon       =   "OrdenProduccion.frx":379C
      Picture         =   "OrdenProduccion.frx":3BDE
      Style           =   1  'Graphical
      TabIndex        =   58
      ToolTipText     =   "Siguiente Registro"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   11280
      MouseIcon       =   "OrdenProduccion.frx":4110
      Picture         =   "OrdenProduccion.frx":4552
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Ultimo Registro"
      Top             =   7440
      Width           =   375
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
      Left            =   9360
      Picture         =   "OrdenProduccion.frx":4A84
      TabIndex        =   52
      Top             =   7440
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame FrameDetalle 
      BorderStyle     =   0  'None
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
      Height          =   4815
      Left            =   120
      TabIndex        =   34
      Top             =   2520
      Width           =   11565
      Begin VB.Frame FrameDetalle2 
         Enabled         =   0   'False
         Height          =   1215
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   11295
         Begin MSMask.MaskEdBox MskDes 
            Height          =   288
            Left            =   7800
            TabIndex        =   43
            Top             =   720
            Width           =   1164
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            PromptChar      =   "_"
         End
         Begin VB.CheckBox ChkLam 
            Caption         =   "Lam. x Unid."
            Height          =   372
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   975
         End
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   9600
            MaxLength       =   15
            TabIndex        =   44
            Top             =   720
            Width           =   1572
         End
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   2
            TabIndex        =   37
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   360
            Width           =   492
         End
         Begin VB.TextBox TxtDocDet 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   6120
            MaxLength       =   15
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox TxtPas 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7800
            MaxLength       =   10
            TabIndex        =   38
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   360
            Width           =   1176
         End
         Begin MSMask.MaskEdBox MskReq 
            Height          =   285
            Left            =   2160
            TabIndex        =   40
            Top             =   720
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskSal 
            Height          =   285
            Left            =   5880
            TabIndex        =   42
            Top             =   720
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Enabled         =   0   'False
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskEnt 
            Height          =   285
            Left            =   4200
            TabIndex        =   41
            Top             =   720
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Desper."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   2
            Left            =   7080
            TabIndex        =   56
            Top             =   720
            Width           =   672
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
            Height          =   252
            Left            =   720
            TabIndex        =   53
            Top             =   360
            Width           =   6972
         End
         Begin VB.Label LblPasada 
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
            Height          =   252
            Left            =   9000
            TabIndex        =   51
            Top             =   360
            Width           =   2172
         End
         Begin VB.Label Label1 
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
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   50
            Top             =   120
            Width           =   732
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Saldo"
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
            Left            =   5400
            TabIndex        =   49
            Top             =   720
            Width           =   510
         End
         Begin VB.Label Label1 
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
            Left            =   1200
            TabIndex        =   48
            Top             =   720
            Width           =   885
         End
         Begin VB.Label Label1 
            Caption         =   "Pasada"
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
            Index           =   5
            Left            =   7800
            TabIndex        =   47
            Top             =   120
            Width           =   1092
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Obser."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   6
            Left            =   9000
            TabIndex        =   46
            Top             =   720
            Width           =   564
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Entregado"
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
            Left            =   3360
            TabIndex        =   45
            Top             =   720
            Width           =   870
         End
      End
   End
   Begin VB.CommandButton CmdEditar2 
      Caption         =   "Editar"
      Height          =   495
      Left            =   2640
      Picture         =   "OrdenProduccion.frx":4FB6
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7440
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton CmdBorrar2 
      Caption         =   "B&orrar"
      Height          =   495
      Left            =   7680
      Picture         =   "OrdenProduccion.frx":54E8
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   7440
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton CmdCancelar2 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6000
      Picture         =   "OrdenProduccion.frx":5A1A
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7440
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton CmdGrabar2 
      Caption         =   "G&rabar"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4320
      Picture         =   "OrdenProduccion.frx":5F4C
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   7440
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton CmdAgregar2 
      Caption         =   "A&gregar"
      Height          =   495
      Left            =   960
      Picture         =   "OrdenProduccion.frx":647E
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7440
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.Frame FrameEncabezado 
      Caption         =   "Encabezado de Orden"
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
      Height          =   2655
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   720
         Left            =   8760
         Picture         =   "OrdenProduccion.frx":69B0
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   720
         Left            =   1560
         Picture         =   "OrdenProduccion.frx":6EEA
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&uscar"
         Height          =   720
         Left            =   7320
         Picture         =   "OrdenProduccion.frx":72C1
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Height          =   720
         Left            =   10200
         Picture         =   "OrdenProduccion.frx":7749
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salida"
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&Borrar"
         Height          =   720
         Left            =   5880
         Picture         =   "OrdenProduccion.frx":7C64
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   4440
         Picture         =   "OrdenProduccion.frx":822C
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   720
         Left            =   3000
         Picture         =   "OrdenProduccion.frx":8763
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   720
         Left            =   120
         Picture         =   "OrdenProduccion.frx":8CBF
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1800
         Width           =   1400
      End
      Begin VB.Frame FrameEncabezado2 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   11535
         Begin VB.ComboBox CboEstado 
            BackColor       =   &H0080C0FF&
            Height          =   288
            ItemData        =   "OrdenProduccion.frx":903C
            Left            =   10320
            List            =   "OrdenProduccion.frx":9046
            TabIndex        =   55
            Text            =   "ABIERTA"
            Top             =   720
            Width           =   1212
         End
         Begin VB.TextBox TxtUsu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Usuario"
            DataSource      =   "DataEncabezado"
            Height          =   285
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1692
         End
         Begin VB.TextBox TxtCli 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   4
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   1080
            Width           =   1335
         End
         Begin VB.TextBox TxtFicTec 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   3
            Top             =   720
            Width           =   1332
         End
         Begin MSMask.MaskEdBox MskFec 
            Height          =   288
            Index           =   0
            Left            =   6120
            TabIndex        =   1
            Top             =   240
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtDoc 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   345
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   0
            Top             =   240
            Width           =   3132
         End
         Begin MSMask.MaskEdBox MskFec 
            Height          =   288
            Index           =   1
            Left            =   8880
            TabIndex        =   2
            Top             =   240
            Width           =   1332
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   3
            Left            =   9720
            TabIndex        =   54
            Top             =   720
            Width           =   600
         End
         Begin VB.Label LblFicTec 
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
            Height          =   252
            Left            =   3000
            TabIndex        =   28
            Top             =   720
            Width           =   6612
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Entrega"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   6
            Left            =   7560
            TabIndex        =   27
            Top             =   240
            Width           =   1224
         End
         Begin VB.Label Label6 
            Caption         =   "Cliente"
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
            TabIndex        =   26
            Top             =   1080
            Width           =   1332
         End
         Begin VB.Label LblCli 
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
            Height          =   252
            Left            =   3000
            TabIndex        =   6
            Top             =   1080
            Width           =   6612
         End
         Begin VB.Label Label6 
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
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   1332
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Orden"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   348
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   888
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Apertura"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   0
            Left            =   4800
            TabIndex        =   23
            Top             =   240
            Width           =   1284
         End
      End
   End
End
Attribute VB_Name = "OrdenProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As String

Dim VCantidadMateriaPrima As Double
Dim VCodigoMateriaPrima As String
Dim VBodega As String
Dim VNumeroPedido As String
Dim VCliente As String
Dim VFechaApertura As Date

Dim VCantidad As Long
Dim BEditarEncabezado As Boolean
Dim BEditarDetalle As Boolean

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim BLinea As Boolean
Dim BCliente As Boolean
Dim BFichaTecnica As Boolean
Dim BPasada As Boolean
Dim REncabezado As New ADODB.Recordset
Dim RDetalle As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim RBuscaCliente As New ADODB.Recordset
Dim RBuscaSigDoc As New ADODB.Recordset
Dim RBuscaDetalle As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaPasada As New ADODB.Recordset
Dim VUnidadesxLamina As Integer
Dim VTexto As String

Dim VLinea As String
Dim VPasada As String




Sub Botones1()
    If Bandera = True Then
         FrameEncabezado2.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
    Else
         FrameEncabezado2.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalle2.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdEditar2.Enabled = False
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalle2.Enabled = False
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

Sub BotonesVisiblesEncabezado()
    If Bandera4 = True Then
         CmdAgregar.Visible = True
         CmdEditar.Visible = True
         CmdGrabar.Visible = True
         CmdCancelar.Visible = True
         CmdBorrar.Visible = True
         CmdBuscar.Visible = True
         CmdImprimir.Visible = True
         CmdSalida.Visible = True
    Else
         CmdAgregar.Visible = False
         CmdEditar.Visible = False
         CmdGrabar.Visible = False
         CmdCancelar.Visible = False
         CmdBorrar.Visible = False
         CmdBuscar.Visible = False
         CmdImprimir.Visible = False
         CmdSalida.Visible = False
    End If
End Sub




Private Sub ChkLam_KeyPress(KeyAscii As Integer)
         If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub CmdAgregar2_Click()
On Error Resume Next
        
    Bandera2 = True
    Botones2
    Limpia_CamposDetalle
    MskReq.Text = VCantidad
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DbGridDetalle.Enabled = False
    'SE ASIGNA AL DOCUMENTO DE DETALLE EL DOCUMENTO DEL ENCABEZADO
    TxtDocDet.Text = VDocumento
    TxtLin.SetFocus
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next

            VDocumento = TxtDoc.Text

            mensaje = MsgBox("¿Esta Seguro De Borrar El Registro", vbOKCancel + vbExclamation + vbDefaultButton2, "Esta Seguro")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                                
                            REncabezado.Delete
                            
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    If Err <> 0 Then
                                        Conexion.RollbackTrans
                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                        Err.Clear
                                    End If
                                Else 'ORACLE
                                    'SI HAY ERRORES
                                    If Err = -2147217873 Then
                                        Conexion.RollbackTrans
                                        MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                        Err.Clear
                                    ElseIf Err <> -2147217873 And Err <> 0 Then
                                        Conexion.RollbackTrans
                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                        Err.Clear
                                    End If
                                End If
                                
                            'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                                REncabezado.Requery
                                'MUEVE AL SIGUIENTE REGISTRO
                                REncabezado.MoveLast
                                'SI HAY ERRORES
                                If Err <> 0 Then
                                    'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                    Err.Clear
                                End If
                                
                                Llena_CamposEncabezado
                                
                                    Set RDetalle = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Linea = L.Linea")
                                            Else 'ORACLE
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Linea) = (L.Linea)")
                                            End If
                                                Llena_CamposDetalle
                                                Set DbGridDetalle.DataSource = RDetalle
                
            End If
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
            
            VDocumento = TxtDocDet.Text
            VLinea = TxtLin.Text
            
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                        
                   'BORRA EL REGISTRO
                        Conexion.Execute "Delete From DetalleOrdenProduccion Where Documento = '" & VDocumento & "' And Linea = '" & VLinea & "'"
                    
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147217873 Then
                                Conexion.RollbackTrans
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            End If
                        End If
                        
                    
                    'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RDetalle.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RDetalle.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                                                
                        Llena_CamposDetalle
                        
                MousePointer = 0
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
                         Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Linea = L.Linea")
                    Else 'ORACLE
                         Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Linea) = (L.Linea)")
                    End If
                         Llena_CamposDetalle
                         Set DbGridDetalle.DataSource = RDetalle
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click()
    mensaje = InputBox("Orden a Buscar")
    If mensaje <> "" Then
                REncabezado.MoveFirst
                    If GOrigenDeDatos = "AmaproAccess" Then
                        REncabezado.Find "Documento = '" & mensaje & "'"
                    Else
                        REncabezado.Find "Documento = '" & UCase(mensaje) & "'"
                    End If
                
                                                
                Llena_CamposEncabezado
                
                        Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Linea = L.Linea")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Linea) = (L.Linea)")
                                End If
                                Llena_CamposDetalle
                                Set DbGridDetalle.DataSource = RDetalle
    End If
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    Llena_CamposEncabezado
    FrameDetalle.Visible = True
    DbGridDetalle.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    
    DbGridDetalle.Enabled = True
    Bandera2 = False
    Botones2
    Llena_CamposDetalle
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next

    'VALIDA SI TIENE ACCESO
    'If GEditar = True Then
    'Else
    '    If TxtEst.Text = "CERRADA" Then
    '        MsgBox "Esta Orden Ya Esta Cerrada, No Se Puede Modificar", vbOKOnly + vbInformation, "Informacion"
    '        Exit Sub
    '    End If
    'End If

    BEditarEncabezado = True
    Bandera = True
    Botones1
    TxtDoc.Enabled = False
    MskFec.Item(0).SetFocus
    TxtUsu.Text = GUsuario
    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False
End Sub


Private Sub CmdEditar2_Click()
On Error Resume Next
    

    'VALIDA SI TIENE ACCESO
'    If GEditar = True Then
'        MskReq.Enabled = True
'    Else
'        MskReq.Enabled = False
'    End If
    
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DbGridDetalle.Enabled = False
    BEditarDetalle = True
    Bandera2 = True
    Botones2
    VDocumento = TxtDocDet.Text
    VLinea = TxtLin.Text
    VPasada = TxtPas.Text
    MskReq.SetFocus
End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
                        

                        'REVISA LA LINEA
                        If TxtLin.Text = "" Then
                            MsgBox "Codigo De Linea No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                            TxtLin.SetFocus
                            Exit Sub
                        End If
                                            
                        'REVISA LA PASADA
                        If TxtPas.Text = "" Then
                            MsgBox "Codigo De Pasada No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                            TxtPas.SetFocus
                            Exit Sub
                        End If
                                  
                      
                        MskSal.Text = Val(MskReq.Text) - Val(MskEnt.Text)
                        VCantidad = MskReq.Text
                        
                        
                        
                            If BEditarDetalle = False Then
                                    VTexto = "'" & TxtDocDet.Text & "', '" ' DOCUMENTO
                                    VTexto = VTexto & TxtLin.Text & "', '" 'LINEA
                                    VTexto = VTexto & TxtPas.Text & "', " 'PASADA
                                    VTexto = VTexto & MskReq.Text & ", " 'REQUIZADO
                                    VTexto = VTexto & MskEnt.Text & ", " 'ENTREGADO
                                    VTexto = VTexto & MskSal.Text & ", " 'SALDO
                                    VTexto = VTexto & MskDes.Text & ", '" 'DESPERDICIO
                                    VTexto = VTexto & TxtObs.Text & "'" 'OBSERVACIONES
                                    
                                    Conexion.Execute "Insert Into DetalleOrdenProduccion Values(" & VTexto & ")"
                            Else 'SI ESTA EDITANDO
                                    'VTexto = "'" & TxtDocDet.Text & "', '" ' DOCUMENTO
                                    VTexto = "Linea = '" & TxtLin.Text & "', " 'LINEA
                                    VTexto = VTexto & "Pasada = '" & TxtPas.Text & "', " 'PASADA
                                    VTexto = VTexto & "Requerido = " & MskReq.Text & ", "  'REQUIZADO
                                    VTexto = VTexto & "Entregado = " & MskEnt.Text & ", " 'ENTREGADO
                                    VTexto = VTexto & "Saldo = " & MskSal.Text & ", " 'SALDO
                                    VTexto = VTexto & "Desperdicio = " & MskDes.Text & ", " 'DESPERDICIO
                                    VTexto = VTexto & "Observaciones = '" & TxtObs.Text & "'" 'OBSERVACIONES
                                    VTexto = VTexto & " Where Documento = '" & VDocumento & "' And Linea = '" & VLinea & "' And Pasada = '" & VPasada & "'"
                                    
                                    Conexion.Execute "Update DetalleOrdenProduccion Set " & VTexto
                            End If
                                        
                                    'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                        End If
                                    Else 'ORACLE
                                        If Err = -2147217873 Then
                                            MsgBox "Orden, Linea, y Pasada Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                            TxtLin.SetFocus
                                      'SI ES CUALQUIER OTRO ERROR
                                        ElseIf Err <> -2147217873 And Err <> 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                        End If
                                    End If
                                        
                        
                                    Bandera2 = False
                                    Botones2
                                    RDetalle.Requery
                                    RDetalle.MoveLast
                                    Llena_CamposDetalle
                                    DbGridDetalle.Enabled = True
                                    CmdAgregar2.SetFocus
      
    
End Sub


Private Sub CmdAgregar_Click()
    On Error Resume Next
    
    TxtDoc.Enabled = True
    Bandera = True
    Botones1
    BEditarEncabezado = False
    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False
    Limpia_CamposEncabezado
    TxtUsu.Text = GUsuario
    MskFec.Item(0).Text = Date
    CboEstado.Text = "ABIERTA"
    TxtDoc.SetFocus
    
    
    'BUSCA EL DOCUMENTO MAXIMO Y LE ASIGNA 1
    'Set RBuscaSigDoc = Db.OpenRecordset("Select Max(Documento) from EncabezadoOrdenProduccion")
    '    If RBuscaSigDoc.RecordCount > 0 Then
    '        If IsNull(RBuscaSigDoc(0)) Then
    '            TxtDoc.Text = "1"
    '        Else
    '            TxtDoc.Text = RBuscaSigDoc(0) + 1
    '        End If
    '    End If
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    
    'NUMERO DE ORDEN
    If TxtDoc.Text = "" Then
        MsgBox "Documento No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtDoc.SetFocus
        Exit Sub
    End If
    
    'FICHA TECNICA
    If TxtFicTec.Text = "" Then
        MsgBox "Ficha Tecnica No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtFicTec.SetFocus
        Exit Sub
    End If
    
    'CLIENTE
    If TxtCli.Text = "" Then
        MsgBox "Codigo De Cliente No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtCli.SetFocus
        Exit Sub
    End If
    
    'FECHA APERTURA
    If Not IsDate(MskFec.Item(0).Text) Then
        MsgBox "Fecha De Apertura Incorrecta", vbOKOnly + vbInformation, "Informacion"
        MskFec.Item(0).SetFocus
        Exit Sub
    End If
    
    'FECHA DE ENTREGA
    If Not IsDate(MskFec.Item(1).Text) Then
        MsgBox "Fecha De Entrega Incorrecta", vbOKOnly + vbInformation, "Informacion"
        MskFec.Item(1).SetFocus
        Exit Sub
    End If
    
    'VERIFICA EL ESTADO DE LA ORDEN
    If (CboEstado.Text <> "ABIERTA" And CboEstado.Text <> "CERRADA") Then
        MsgBox "El Estado De La Orden Es Incorrecto", vbOKOnly + vbInformation, "Informacion"
        CboEstado.SetFocus
        Exit Sub
    End If
    
    MskFec.Item(0).Text = Format(MskFec.Item(0).Text)
    MskFec.Item(1).Text = Format(MskFec.Item(1).Text)
    
    VDocumento = TxtDoc.Text
    VCliente = TxtCli.Text
    VFechaApertura = MskFec.Item(0).Text
        
    'GRABA DATOS
    'AGREGAR
                    If BEditarEncabezado = False Then
                            VTexto = "'" & TxtDoc.Text & "', '" 'DOCUMENTO
                            VTexto = VTexto & TxtFicTec.Text & "', " 'FICHA TECNICA
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "#" & Format(MskFec.Item(0).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = VTexto & "To_Date('" & MskFec.Item(0).Text & "', 'dd/mm/yyyy')" & ", "  'FECHA
                            End If
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "#" & Format(MskFec.Item(1).Text, "mm/dd/yyyy") & "#, '" 'FECHA
                            Else 'ORACLE
                                VTexto = VTexto & "To_Date('" & MskFec.Item(1).Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                            End If
                            VTexto = VTexto & TxtCli.Text & "', '" 'TIPO DE DOCUMENTO
                            VTexto = VTexto & CboEstado & "', '" 'NUMERO DE DOCUMENTO
                            VTexto = VTexto & GUsuario & "'" 'NUMERO DE DOCUMENTO
                            
                            Conexion.Execute "Insert Into EncabezadoOrdenProduccion Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            VTexto = "FichaTecnica = '" & UCase(TxtFicTec.Text) & "', " 'TIPO DE DOCUMENTO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "FechaApertura = #" & Format(MskFec.Item(0).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = VTexto & "FechaApertura = To_Date('" & MskFec.Item(0).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "FechaEntrega = #" & Format(MskFec.Item(1).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = VTexto & "FechaEntrega = To_Date('" & MskFec.Item(1).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "Cliente = '" & UCase(TxtCli.Text) & "', " 'CLIENTE
                            VTexto = VTexto & "Estado = '" & UCase(CboEstado) & "', " 'ESTADO
                            VTexto = VTexto & "Usuario = '" & GUsuario & "'" 'OBSERVACIONES
                            VTexto = VTexto & " Where Documento = '" & VDocumento & "'" 'DOCUMENTO
                            
                            Conexion.Execute "UPDATE EncabezadoOrdenProduccion SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Orden Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtDoc.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Orden Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtDoc.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        Botones1
                        TxtDoc.Enabled = True
                        
                        REncabezado.Requery
                        REncabezado.MoveFirst
                        REncabezado.Find "Documento = '" & VDocumento & "'"
                                            
                        Llena_CamposEncabezado
                        
                        Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where E.Documento = '" & VDocumento & "' And E.Documento = D.Documento And D.Linea = L.Linea")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where UPPER(E.Documento) = '" & UCase(VDocumento) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Linea) = (L.Linea)")
                                End If
                                Llena_CamposDetalle
                                Set DbGridDetalle.DataSource = RDetalle
                        
   
    'ESCONDE LOS BOTONES DEL ENCABEZADO
    Bandera4 = False
    BotonesVisiblesEncabezado
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    'BOTONES DE DATA
    CmdBotones2.Item(1).Visible = False
    CmdBotones2.Item(2).Visible = False
    CmdBotones2.Item(3).Visible = False
    CmdBotones2.Item(4).Visible = False
                
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    
    FrameDetalle.Visible = True
    DbGridDetalle.Visible = True
    FrameDetalle.Enabled = True
    FrameEncabezado.Enabled = False
    CmdAgregar2.SetFocus
End Sub

Private Sub CmdImprimir_Click()
MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDoc.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'gtituloreporte = "letras = '" & VLetras & "'"
                
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "OrdenProduccionDetalle.rpt"
                Else
                    GNombreReporte = "OrdenProduccionDetalleO.rpt"
                End If
                GCriteriaReporte = "{EncabezadoOrdenProduccion.Documento} = '" & TxtDoc.Text & "'"
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
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    CmdBotones2.Item(1).Visible = True
    CmdBotones2.Item(2).Visible = True
    CmdBotones2.Item(3).Visible = True
    CmdBotones2.Item(4).Visible = True
    FrameDetalle.Visible = True
    FrameEncabezado.Enabled = True
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesVisiblesEncabezado
    
    'ESCONDE LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle

    Set REncabezado = New ADODB.Recordset
            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoOrdenProduccion Order By Documento")
            REncabezado.MoveLast
                Llena_CamposEncabezado
                
            Set RDetalle = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Linea = L.Linea")
                    Else 'ORACLE
                        Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Linea) = (L.Linea)")
                    End If
                        Llena_CamposDetalle
                        Set DbGridDetalle.DataSource = RDetalle
                        
    
End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BLinea = True Then
            TxtLin.Text = DBGridBusqueda.Columns(0)
            TxtLin.SetFocus
        ElseIf BCliente = True Then
            TxtCli.Text = DBGridBusqueda.Columns(0)
            TxtCli.SetFocus
        ElseIf BFichaTecnica = True Then
            TxtFicTec.Text = DBGridBusqueda.Columns(0)
            TxtFicTec.SetFocus
        ElseIf BPasada = True Then
            TxtPas.Text = DBGridBusqueda.Columns(0)
            TxtPas.SetFocus
        End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BLinea = True Then
                TxtLin.Text = DBGridBusqueda.Columns(0)
                TxtLin.SetFocus
            ElseIf BCliente = True Then
                TxtCli.Text = DBGridBusqueda.Columns(0)
                TxtCli.SetFocus
            ElseIf BFichaTecnica = True Then
                TxtFicTec.Text = DBGridBusqueda.Columns(0)
                TxtFicTec.SetFocus
            ElseIf BPasada = True Then
                TxtPas.Text = DBGridBusqueda.Columns(0)
                TxtPas.SetFocus
            End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
        End If
End Sub


Private Sub DbGridDetalle_HeadClick(ByVal ColIndex As Integer)
        RDetalle.Sort = RDetalle.Fields(ColIndex).Name
End Sub


Private Sub DbGridDetalle_SelChange(Cancel As Integer)
                Llena_CamposDetalle
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If

End Sub

Private Sub Form_Load()
    
      Set REncabezado = New ADODB.Recordset
            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoOrdenProduccion Order By Documento")
                Llena_CamposEncabezado
                
    
                Set RDetalle = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                         Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Linea = L.Linea")
                    Else 'ORACLE
                         Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Linea, L.Descrip, D.Pasada, D.Requerido, D.Entregado, D.Saldo, D.Desperdicio, D.Observaciones From EncabezadoOrdenProduccion E, DetalleOrdenProduccion D, Lineas L Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Linea) = (L.Linea)")
                    End If
                         Llena_CamposDetalle
                         Set DbGridDetalle.DataSource = RDetalle
                        
        
End Sub

Private Sub MskDes_GotFocus()
        MskDes.SelStart = 0
        MskDes.SelLength = Len(MskDes.Text)
End Sub

Private Sub MskDes_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskEnt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskEnt_LostFocus()
        'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES
        If ChkLam.Value = 1 Then
            If IsNumeric(MskReq.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
                        End If
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                            
                        MskEnt.Text = Val(MskEnt.Text) * Val(VUnidadesxLamina)
                        
                         'SALDO = Req - Ent
                        MskSal.Text = Val(MskReq.Text) - Val(MskEnt.Text)
            End If
                            
                        
        End If

End Sub

Private Sub MskFec_GotFocus(Index As Integer)
        MskFec.Item(Index).SelStart = 0
        MskFec.Item(Index).SelLength = Len(MskFec.Item(Index))
End Sub

Private Sub MskFec_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskReq_GotFocus()
        MskReq.SelStart = 0
        MskReq.SelLength = Len(MskReq.Text)
End Sub

Private Sub MskReq_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub MskReq_LostFocus()
        'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES
        If ChkLam.Value = 1 Then
            If IsNumeric(MskReq.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
                        End If
                        
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                            
                        MskReq.Text = Val(MskReq.Text) * Val(VUnidadesxLamina)
            End If
                            
                        
        End If

End Sub

Private Sub MskSal_GotFocus()
        MskSal.SelStart = 0
        MskSal.SelLength = Len(MskSal.Text)
End Sub

Private Sub MskSal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub



Private Sub MskSal_LostFocus()
                'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES
        If ChkLam.Value = 1 Then
            If IsNumeric(MskReq.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
                        End If
                        
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                            
                        MskSal.Text = Val(MskSal.Text) * Val(VUnidadesxLamina)
            End If
                        
        End If

End Sub

Private Sub Txtbuscar_Change()
        
        Set RBusqueda = New ADODB.Recordset

        'LINEA
        If BLinea = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where Linea Like '%" & TxtBuscar.Text & "%' Order by Linea")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where UPPER(Linea) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Linea")
                    End If
            ElseIf OptDescripcion.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where Descrip Like '%" & TxtBuscar.Text & "%' Order by Linea")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Linea")
                    End If
            End If
        'CLIENTE
        ElseIf BCliente = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where CodigoCliente Like '%" & TxtBuscar.Text & "%' Order by CodigoCliente")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where UPPER(CodigoCliente) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoCliente")
                    End If
            ElseIf OptDescripcion.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by CodigoCliente")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoCliente")
                    End If
            End If
        'FICHA TECNICA
        ElseIf BFichaTecnica = True Then
                    'OPCION POR DESCRIPCION
                    If OptDescripcion.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Descrip Like '%" & TxtBuscar.Text & "%' And Activa = -1")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' And Activa = -1")
                        End If
                    'OPCION DE CODIGO
                    ElseIf OptCodigo = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Esp_Tec Like '%" & TxtBuscar.Text & "%' And Activa = -1")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscar.Text) & "%' And Activa = -1")
                        End If
                    End If
        'PASADAS
        ElseIf BPasada = True Then
                    'OPCION POR DESCRIPCION
                    If OptDescripcion.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Pasadas Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Pasadas Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                        End If
                    'OPCION DE CODIGO
                    ElseIf OptCodigo = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Pasadas Where Codigo Like '%" & TxtBuscar.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Pasadas Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                        End If
                    End If
        End If
            
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"

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


Private Sub TxtFicTec_Change()
        Set RBuscaFicha = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            If RBuscaFicha.RecordCount > 0 Then
                lblFicTec.Caption = RBuscaFicha!Descrip
            Else
                lblFicTec.Caption = ""
            End If
End Sub

Private Sub TxtFicTec_DblClick()
            BLinea = False
            BCliente = False
            BFichaTecnica = True
            BPasada = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial, Envases from FichaTecnica Where Activa = -1")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
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
            BLinea = False
            BCliente = False
            BFichaTecnica = True
            BPasada = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial, Envases from FichaTecnica Where Activa = -1")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
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

Private Sub TxtCli_Change()
            Set RBuscaCliente = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where CodigoCliente = '" & TxtCli.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where UPPER(CodigoCliente) = '" & UCase(TxtCli.Text) & "'")
                End If
                If RBuscaCliente.RecordCount > 0 Then
                    LblCli.Caption = RBuscaCliente!Descripcion
                Else
                    LblCli.Caption = ""
                End If
End Sub

Private Sub TxtCli_DblClick()
            BLinea = False
            BCliente = True
            BFichaTecnica = False
            BPasada = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Order by CodigoCliente")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtCli_GotFocus()
            TxtCli.SelStart = 0
            TxtCli.SelLength = Len(TxtCli.Text)
End Sub

Private Sub TxtCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
            BLinea = False
            BCliente = True
            BFichaTecnica = False
            BPasada = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Order by CodigoCliente")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
    End If
End Sub

Private Sub TxtLin_Change()
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLin.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                End If
                
                If RBuscaLinea.RecordCount > 0 Then
                        LblLinea.Caption = RBuscaLinea!Descrip
                Else
                        LblLinea.Caption = ""
                End If
End Sub
Private Sub Txtlin_DblClick()
            BLinea = True
            BCliente = False
            BFichaTecnica = False
            BPasada = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Order by Linea")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
End Sub
Private Sub TxtLin_GotFocus()
        TxtLin.SelStart = 0
        TxtLin.SelLength = Len(TxtLin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
        'SI PRECIONA ENTER
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
        'SI PRECIONA LA TECLA DE SIGNO +
        If KeyAscii = 43 Then
            BLinea = True
            BCliente = False
            BFichaTecnica = False
            BPasada = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Order by Linea")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
        
        End If
End Sub

Private Sub TxtDoc_GotFocus()
    TxtDoc.SelStart = 0
    TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Private Sub TxtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub


Private Sub TxtPas_Change()
        Set RBuscaPasada = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaPasada, "Select Descripcion From Pasadas Where Codigo = '" & TxtPas.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaPasada, "Select Descripcion From Pasadas Where UPPER(Codigo) = '" & UCase(TxtPas.Text) & "'")
            End If
            If RBuscaPasada.RecordCount > 0 Then
                LblPasada.Caption = RBuscaPasada!Descripcion
            Else
                LblPasada.Caption = ""
            End If
End Sub

Private Sub TxtPas_DblClick()
            BLinea = False
            BCliente = False
            BFichaTecnica = False
            BPasada = True
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Codigo, descripcion From Pasadas")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
End Sub

Private Sub TxtPas_GotFocus()
        TxtPas.SelStart = 0
        TxtPas.SelLength = Len(TxtPas.Text)
End Sub

Private Sub TxtPas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BLinea = False
            BCliente = False
            BFichaTecnica = False
            BPasada = True
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Codigo, descripcion From Pasadas")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
        End If
End Sub

Public Sub Llena_CamposEncabezado()
On Error Resume Next
            If REncabezado.RecordCount > 0 Then
                If IsNull(REncabezado!Documento) Then
                    TxtDoc.Text = ""
                Else
                    TxtDoc.Text = REncabezado!Documento
                End If
                If IsNull(REncabezado!FichaTecnica) Then
                    TxtFicTec.Text = ""
                Else
                    TxtFicTec.Text = REncabezado!FichaTecnica
                End If
                If IsNull(REncabezado!FechaApertura) Then
                    MskFec.Item(0).Text = ""
                Else
                    MskFec.Item(0).Text = REncabezado!FechaApertura
                End If
                If IsNull(REncabezado!FechaEntrega) Then
                    MskFec.Item(1).Text = ""
                Else
                    MskFec.Item(1).Text = REncabezado!FechaEntrega
                End If
                If IsNull(REncabezado!Cliente) Then
                    TxtCli.Text = ""
                Else
                    TxtCli.Text = REncabezado!Cliente
                End If
                If IsNull(REncabezado!Estado) Then
                    CboEstado.Text = ""
                Else
                    CboEstado.Text = REncabezado!Estado
                End If
                If IsNull(REncabezado!Usuario) Then
                    TxtUsu.Text = ""
                Else
                    TxtUsu.Text = REncabezado!Usuario
                End If
            Else
                TxtDoc.Text = ""
                TxtFicTec.Text = ""
                MskFec.Item(0).Text = ""
                MskFec.Item(1).Text = ""
                TxtCli.Text = ""
                TxtUsu.Text = ""
            End If
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Llena_CamposDetalle()
On Error Resume Next
            If RDetalle.RecordCount > 0 Then
                If IsNull(RDetalle!Documento) Then
                    TxtDocDet.Text = ""
                Else
                    TxtDocDet.Text = RDetalle!Documento
                End If
                If IsNull(RDetalle!Linea) Then
                    TxtLin.Text = ""
                Else
                    TxtLin.Text = RDetalle!Linea
                End If
                If IsNull(RDetalle!Pasada) Then
                    TxtPas.Text = ""
                Else
                    TxtPas.Text = RDetalle!Pasada
                End If
                If IsNull(RDetalle!Requerido) Then
                    MskReq.Text = 0
                Else
                    MskReq.Text = RDetalle!Requerido
                End If
                If IsNull(RDetalle!Entregado) Then
                    MskEnt.Text = 0
                Else
                    MskEnt.Text = RDetalle!Entregado
                End If
                If IsNull(RDetalle!Saldo) Then
                    MskSal.Text = 0
                Else
                    MskSal.Text = RDetalle!Saldo
                End If
                If IsNull(RDetalle!Desperdicio) Then
                    MskDes.Text = 0
                Else
                    MskDes.Text = RDetalle!Desperdicio
                End If
                If IsNull(RDetalle!Observaciones) Then
                    TxtObs.Text = ""
                Else
                    TxtObs.Text = RDetalle!Observaciones
                End If
            Else
                TxtDocDet.Text = ""
                TxtLin.Text = ""
                TxtPas.Text = ""
                MskReq.Text = 0
                MskEnt.Text = 0
                MskSal.Text = 0
                MskDes.Text = 0
                TxtObs.Text = ""
            End If
            
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Limpia_CamposEncabezado()
                TxtDoc.Text = ""
                TxtFicTec.Text = ""
                MskFec.Item(0).Text = ""
                MskFec.Item(1).Text = ""
                TxtCli.Text = ""
                TxtUsu.Text = ""
End Sub

Public Sub Limpia_CamposDetalle()
                TxtDocDet.Text = ""
                TxtLin.Text = ""
                TxtPas.Text = ""
                MskReq.Text = 0
                MskEnt.Text = 0
                MskSal.Text = 0
                MskDes.Text = 0
                TxtObs.Text = ""
End Sub



