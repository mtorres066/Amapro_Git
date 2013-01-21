VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form TrasladosMateriaPrimaProceso 
   BackColor       =   &H00008000&
   Caption         =   "Traslados De Inventario"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11490
   Icon            =   "TrasladosMateriaPrimaProcesos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   11490
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
      Height          =   8415
      Left            =   120
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   11295
      Begin MSDataGridLib.DataGrid DbGridBuscar 
         Height          =   7335
         Left            =   120
         TabIndex        =   41
         Top             =   960
         Width           =   11055
         _ExtentX        =   19500
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
         Height          =   615
         Left            =   10440
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1560
         TabIndex        =   39
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   40
         Top             =   600
         Width           =   4935
      End
   End
   Begin VB.CommandButton CmdImprimirCedula 
      Caption         =   "Cedula"
      Height          =   480
      Left            =   6960
      Picture         =   "TrasladosMateriaPrimaProcesos.frx":293C
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   7800
      Width           =   1100
   End
   Begin VB.Frame FrameEncabezado 
      Caption         =   "Encabezado"
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
      Height          =   2652
      Left            =   120
      TabIndex        =   46
      Top             =   0
      Width           =   11295
      Begin VB.CommandButton CmdBuscar2 
         Caption         =   "Siguiente Documento"
         Height          =   480
         Left            =   7440
         TabIndex        =   15
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   480
         Left            =   8640
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":2A86
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&Editar"
         Height          =   480
         Left            =   1440
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":2FB8
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&uscar Documento"
         Height          =   480
         Left            =   6240
         TabIndex        =   14
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Caption         =   "&Salida"
         Height          =   480
         Left            =   9840
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":34EA
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Salida"
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&Borrar"
         Height          =   480
         Left            =   5040
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":555C
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   480
         Left            =   3840
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":5A8E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   480
         Left            =   2640
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":5FC0
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   480
         Left            =   240
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":64F2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Frame FrameCompras 
         Enabled         =   0   'False
         Height          =   1692
         Left            =   120
         TabIndex        =   47
         Top             =   240
         Width           =   11055
         Begin VB.TextBox TxtBodSal 
            Appearance      =   0  'Flat
            Height          =   288
            Left            =   1560
            MaxLength       =   3
            TabIndex        =   8
            Top             =   1320
            Width           =   1452
         End
         Begin VB.TextBox TxtTipDoc 
            Appearance      =   0  'Flat
            DataField       =   "TipoDeDocumento"
            Height          =   285
            Left            =   4680
            MaxLength       =   10
            TabIndex        =   3
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TxtNumDoc 
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   2
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox TxtEncabezado 
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
            Index           =   0
            Left            =   9360
            Locked          =   -1  'True
            MaxLength       =   12
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Index           =   1
            Left            =   9360
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   960
            Width           =   1575
         End
         Begin MSMask.MaskEdBox MskFec 
            Height          =   285
            Left            =   1560
            TabIndex        =   0
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   4
            Top             =   960
            Width           =   4575
         End
         Begin VB.TextBox TxtEncabezado 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            Index           =   2
            Left            =   9360
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox TxtDocTra 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   4680
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label LblBodegaSalida 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000A&
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
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3120
            TabIndex        =   71
            Top             =   1320
            Width           =   3015
         End
         Begin VB.Label Label6 
            Caption         =   "Bodega Salida"
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
            Index           =   6
            Left            =   120
            TabIndex        =   70
            Top             =   1320
            Width           =   1452
         End
         Begin VB.Label LblTipDoc 
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
            Left            =   6240
            TabIndex        =   68
            Top             =   600
            Width           =   4695
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            Index           =   5
            Left            =   3120
            TabIndex        =   67
            Top             =   600
            Width           =   1410
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "# Documento "
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
            Left            =   120
            TabIndex        =   66
            Top             =   600
            Width           =   1215
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
            Index           =   3
            Left            =   8160
            TabIndex        =   55
            Top             =   1320
            Width           =   750
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
            Index           =   2
            Left            =   8160
            TabIndex        =   54
            Top             =   960
            Width           =   885
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
            Left            =   8160
            TabIndex        =   53
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   52
            Top             =   960
            Width           =   1455
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
            Left            =   3120
            TabIndex        =   51
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Traslado"
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
            TabIndex        =   50
            Top             =   240
            Width           =   1335
         End
      End
   End
   Begin VB.Frame FrameDetalle 
      Caption         =   "Detalle"
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
      Height          =   5655
      Left            =   120
      TabIndex        =   42
      Top             =   2760
      Width           =   11325
      Begin MSDataGridLib.DataGrid DbGridDetalle 
         Height          =   2895
         Left            =   120
         TabIndex        =   78
         Top             =   2040
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   5106
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
      Begin VB.CommandButton CmdEditar2 
         Caption         =   "Editar"
         Height          =   495
         Left            =   2040
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":6A24
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5040
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton CmdBorrar2 
         Caption         =   "B&orrar"
         Height          =   495
         Left            =   5640
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":6F56
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   5040
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton CmdCancelar2 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4440
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":7488
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   5040
         Visible         =   0   'False
         Width           =   1100
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
         Left            =   8040
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":79BA
         TabIndex        =   36
         Top             =   5040
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton CmdGrabar2 
         Caption         =   "G&rabar"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3240
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":7EEC
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   5040
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton CmdAgregar2 
         Caption         =   "A&gregar"
         Height          =   495
         Left            =   840
         Picture         =   "TrasladosMateriaPrimaProcesos.frx":841E
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   5040
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.Frame FrameDetalleCompras 
         Enabled         =   0   'False
         Height          =   1695
         Left            =   120
         TabIndex        =   43
         Top             =   240
         Width           =   11055
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   3240
            MaxLength       =   2
            TabIndex        =   20
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox TxtLamReq 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   288
            Left            =   9960
            Locked          =   -1  'True
            TabIndex        =   75
            TabStop         =   0   'False
            Top             =   720
            Width           =   972
         End
         Begin VB.TextBox TxtLamEnt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   288
            Left            =   9960
            Locked          =   -1  'True
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1320
            Width           =   972
         End
         Begin VB.CheckBox ChkLam 
            Caption         =   "Lam. x Unid."
            Height          =   375
            Left            =   3360
            TabIndex        =   72
            Top             =   1200
            Width           =   972
         End
         Begin VB.TextBox TxtOrd 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   120
            MaxLength       =   15
            TabIndex        =   18
            Top             =   360
            Width           =   1452
         End
         Begin VB.TextBox TxtUniMedSal 
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
            Left            =   6480
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   720
            Width           =   972
         End
         Begin MSMask.MaskEdBox MskDifReqCorMas 
            Height          =   285
            Left            =   4440
            TabIndex        =   26
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanRea 
            Height          =   285
            Left            =   8760
            TabIndex        =   30
            Top             =   1320
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskDifReqCor 
            Height          =   285
            Left            =   5520
            TabIndex        =   27
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanDesPro 
            Height          =   285
            Left            =   7680
            TabIndex        =   29
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanDes 
            Height          =   285
            Left            =   6600
            TabIndex        =   28
            Top             =   1320
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   8421631
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanSal 
            Height          =   285
            Left            =   8040
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   720
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   -2147483644
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBodEnt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            MaxLength       =   3
            TabIndex        =   25
            Top             =   1320
            Width           =   612
         End
         Begin VB.TextBox TxtNumIng 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   9960
            TabIndex        =   22
            Top             =   360
            Width           =   975
         End
         Begin VB.TextBox TxtDocDet 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   840
            MaxLength       =   15
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   492
         End
         Begin VB.TextBox TxtCodSal 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0E0FF&
            Height          =   285
            Left            =   3840
            MaxLength       =   15
            TabIndex        =   21
            Top             =   360
            Width           =   1575
         End
         Begin MSMask.MaskEdBox MskFecPro 
            Height          =   285
            Left            =   1680
            TabIndex        =   19
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   12640511
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   82
            Top             =   720
            Width           =   1260
         End
         Begin VB.Label LblDesSal 
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
            Left            =   5520
            TabIndex        =   81
            Top             =   360
            Width           =   4335
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha Produccion"
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
            Left            =   1560
            TabIndex        =   80
            Top             =   120
            Width           =   1560
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Index           =   1
            Left            =   3240
            TabIndex        =   79
            Top             =   120
            Width           =   480
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
            Left            =   1680
            TabIndex        =   77
            Top             =   720
            Width           =   3735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Laminas"
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
            Left            =   9240
            TabIndex        =   76
            Top             =   720
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Laminas"
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
            Left            =   9960
            TabIndex        =   74
            Top             =   1080
            Width           =   705
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Height          =   192
            Index           =   3
            Left            =   120
            TabIndex        =   69
            Top             =   120
            Width           =   516
         End
         Begin VB.Label LblBodegaEntrada 
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
            Left            =   840
            TabIndex        =   64
            Top             =   1320
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "U / Medida"
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
            Left            =   5520
            TabIndex        =   63
            Top             =   720
            Width           =   930
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "De Mas"
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
            Left            =   4440
            TabIndex        =   62
            Top             =   1080
            Width           =   645
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Index           =   8
            Left            =   8760
            TabIndex        =   61
            Top             =   1080
            Width           =   870
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "De Menos"
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
            Left            =   5520
            TabIndex        =   60
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Des.Prov."
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
            Left            =   7680
            TabIndex        =   59
            Top             =   1080
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Des.Proc"
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
            Left            =   6600
            TabIndex        =   58
            Top             =   1080
            Width           =   780
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   "Bodega Entrada"
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
            Left            =   120
            TabIndex        =   57
            Top             =   1080
            Width           =   1365
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Height          =   195
            Index           =   2
            Left            =   9960
            TabIndex        =   56
            Top             =   120
            Width           =   1110
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Index           =   0
            Left            =   7440
            TabIndex        =   49
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label1 
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
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
            Index           =   0
            Left            =   3840
            TabIndex        =   48
            Top             =   120
            Width           =   615
         End
      End
   End
End
Attribute VB_Name = "TrasladosMateriaPrimaProceso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As Long
Dim VDocumentoDetalle As Long
Dim VSumaEgresos As Double

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean

Dim BBodegaEntrada As Boolean
Dim BBodegaSalida As Boolean
Dim BCodigoSalida As Boolean
Dim BNumeroIngreso As Boolean
Dim BDocumento As Boolean
Dim BEditar As Boolean
Dim BEditarEncabezado As Boolean
Dim BEditarDetalle As Boolean


Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaMateriaPrimaSalida As New ADODB.Recordset
Dim RBuscaMateriaPrimaEntrada As New ADODB.Recordset
Dim RBuscaSigDoc As New ADODB.Recordset
Dim RBuscaTipoDocumento As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset
Dim RBuscaDetalle As New ADODB.Recordset
Dim RBuscaEncabezado As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RBuscaBodegaEntrada As New ADODB.Recordset
Dim RBuscaBodegaSalida As New ADODB.Recordset
Dim RBuscaFichaOrden As New ADODB.Recordset
Dim RBuscaOrden As New ADODB.Recordset

Dim VUltimoCodigo As String
Dim VUltimoBodegaEntrada As String
Dim VUltimaOrden As String
Dim VBodegaSalida As String
Dim VUnidadesxLamina As Integer

Dim RBusqueda As New ADODB.Recordset
Dim REncabezado As New ADODB.Recordset
Dim RDetalle As New ADODB.Recordset
Dim VTexto As String
Dim VTexto2 As String



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
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
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
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleCompras.Enabled = True
         CmdAgregar2.Enabled = False
         CmdEditar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdTerminar.Enabled = False
         CmdImprimirCedula.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalleCompras.Enabled = False
         CmdAgregar2.Enabled = True
         CmdEditar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdImprimirCedula.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
    End If

End Sub

Sub BotonesDetalleVisibles()
    If Bandera3 = True Then
         CmdAgregar2.Visible = True
         CmdEditar2.Visible = True
         CmdGrabar2.Visible = True
         CmdCancelar2.Visible = True
         CmdBorrar2.Visible = True
         'CmdImprimirCedula.Visible = True
         CmdTerminar.Visible = True
    Else
         CmdAgregar2.Visible = False
         CmdEditar2.Visible = False
         CmdGrabar2.Visible = False
         CmdCancelar2.Visible = False
         CmdBorrar2.Visible = False
         'CmdImprimirCedula.Visible = False
         CmdTerminar.Visible = False
    
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
    DbGridDetalle.Enabled = False
    TxtDocDet.Text = VDocumento
    
    MskFecPro.Enabled = True
    TxtLin.Enabled = True
    TxtCodSal.Enabled = True
    TxtNumIng.Enabled = True
    
    MskDifReqCorMas.Text = 0
    MskDifReqCor.Text = 0
    MskCanDes.Text = 0
    MskCanDesPro.Text = 0
    MskCanRea.Text = 0
    'ASIGNA EL ULTIMO CODIGO GUARDADO AL NUEVO
    TxtOrd.Text = VUltimaOrden
    TxtCodSal.Text = VUltimoCodigo
    TxtOrd.SetFocus
    TxtBodEnt.Text = VUltimoBodegaEntrada
    
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            If GBorrar = True Then
                'NO HACE NADA PORQUE SI TIENE ACCESO
            ElseIf TxtEncabezado.Item(0).Text = "LIBERADO" Then
                'VERIFICA SI YA FUE LIBERADA LA ENTRADA
                    MsgBox "Esta Recepcion No Se Puede BORRAR Porque Ya Fue Liberada", vbOKOnly + vbExclamation, "Informacion"
                    Exit Sub
            End If


            VDocumento = TxtDocTra.Text
            mensaje = MsgBox("¿Está Seguro De Borrar El Traslado?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                        'BORRA EL ENCABEZADO DE EL PEDIDO
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
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                                            Else 'ORACLE
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                                            End If
                                                Llena_CamposDetalle
                                                Set DbGridDetalle.DataSource = RDetalle
                
                
            End If
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
            VDocumentoDetalle = TxtDocDet.Text
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
            'SI CONTESTA QUE SI QUIERE BORRAR
            
            If mensaje = vbOK Then
                
                   'BORRA EL REGISTRO
                        RDetalle.Delete
                    
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
                  
            End If
            
End Sub

Private Sub CmdBuscar_Click()
    mensaje = InputBox("Transaccion a Buscar")
    If mensaje <> "" Then
                Set REncabezado = New ADODB.Recordset
                Call Abrir_Recordset(REncabezado, "Select * From EncabezadoTrasladosInventario Where Documento = " & mensaje & " Order By Documento")
                If REncabezado.RecordCount > 0 Then
                Else
                    MsgBox "Transaccion No Existe", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
                End If
                                                
                Llena_CamposEncabezado
                
                                 Set RDetalle = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                                            Else 'ORACLE
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
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

    If GEditar = True Then
        'NO HACE NADA PORQUE SI TIENE ACCESO
    ElseIf TxtEncabezado.Item(0).Text = "LIBERADO" Then
        'VERIFICA SI YA FUE LIBERADA LA ENTRADA
        MsgBox "Este Traslado No Se Puede EDITAR Porque Ya Fue Liberada", vbOKOnly + vbExclamation, "Informacion"
        Exit Sub
    End If
        
    BEditarEncabezado = True
    BEditar = True
    Bandera = True
    Botones1
    MskFec.SetFocus
    'ASIGNA AL CAMPO DE REQUERIDO EL USUARIO QUE LO ESTA EDITANDO
    TxtEncabezado.Item(1).Text = GUsuario
    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False
End Sub


Private Sub CmdEditar2_Click()
On Error Resume Next
    DbGridDetalle.Enabled = False
    BEditarDetalle = True
    Bandera2 = True
    Botones2
    VDocumento = TxtDocDet.Text
    DbGridDetalle.Enabled = False
    TxtOrd.SetFocus
    
    MskFecPro.Enabled = False
    TxtLin.Enabled = False
    TxtCodSal.Enabled = False
    TxtNumIng.Enabled = False

End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
        
    'GUARDAMOS EL ULTIMO CODIGO PARA DESPLEGARLO DESPUES A LA HORA DE AGREGAR
    VUltimoCodigo = TxtCodSal.Text
    VUltimoBodegaEntrada = TxtBodEnt.Text
    VUltimaOrden = TxtOrd.Text
        
    'BUSCA EL NUMERO DE INGRESO EN ENTRADAS DE BODEGA
    Set RBuscaTarima = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaTarima, "Select Tarima, Bodega From DetalleEntradasInventario Where FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Linea = '" & TxtLin.Text & "' and Tarima = " & TxtNumIng.Text & " And FichaTecnica = '" & TxtCodSal.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaTarima, "Select Tarima, Bodega From DetalleEntradasInventario Where FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and Tarima = " & TxtNumIng.Text & " And UPPER(FichaTecnica) = '" & UCase(TxtCodSal.Text) & "'")
        End If
        If RBuscaTarima.RecordCount > 0 Then
            If VBodegaSalida = RBuscaTarima!Bodega Then
            Else
                MsgBox "La Bodega De Salida No Coincide Con La Bodega Actual Donde Esta Ubicado El Bulto", vbOKOnly + vbInformation, "Verifique"
                Exit Sub
            End If
        Else
            MsgBox "Bulto/Tarima No Existe, En Inventario", vbOKOnly + vbInformation, "Informacion"
            TxtNumIng.SetFocus
            Exit Sub
        End If
    
            'BUSCA EL NUMERO DE INGRESO EN ENTRADAS DE BODEGA SI YA FUE LIBERADO
            Set RBuscaTarima = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaTarima, "Select DE.Tarima From DetalleEntradasInventario D, EncabezadoEntradasInventario EE Where DE.FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And DE.Linea = '" & TxtLin.Text & "' and DE.Tarima = " & TxtNumIng.Text & " And DE.FichaTecnica = '" & TxtCodSal.Text & "' And DE.Documento = EE.Documento And EE.Estado = 'LIBERADO'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaTarima, "Select DE.Tarima From DetalleEntradasInventario D, EncabezadoEntradasInventario EE Where DE.FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(DE.Linea) = '" & UCase(TxtLin.Text) & "' and DE.Tarima = " & TxtNumIng.Text & " And UPPER(DE.FichaTecnica) = '" & UCase(TxtCodSal.Text) & "' And DE.Documento = EE.Documento And UPPER(EE.Estado) = 'LIBERADO'")
            End If
            
            If RBuscaTarima.RecordCount > 0 Then
            Else
                MsgBox "Bulto/Tarima No Ha Sido Liberado Por Recepcion De Bodega", vbOKOnly + vbInformation, "Informacion"
                TxtNumIng.SetFocus
                Exit Sub
            End If
            
            'REVISAMOS LA CANTIDAD DE SALIDA
            If Not IsNumeric(MskCanSal.Text) Then
               MsgBox "Cantidad De SALIDA Incorrecta", vbOKOnly + vbCritical, "Error"
               MskCanSal.SetFocus
               Exit Sub
            End If
            
            'REVISA LA CANTIDAD REQUISADA DE MENOS
            If Not IsNumeric(MskDifReqCor.Text) Then
               MsgBox "Cantidad De Diferencia Req/Cor Incorrecta", vbOKOnly + vbCritical, "Error"
               MskDifReqCor.SetFocus
               Exit Sub
            End If
            
            'REVISA LA CANTIDAD REQUISADA DE MAS
            If Not IsNumeric(MskDifReqCorMas.Text) Then
               MsgBox "Cantidad De Diferencia Req/Cor Incorrecta", vbOKOnly + vbCritical, "Error"
               MskDifReqCor.SetFocus
               Exit Sub
            End If
               
            'REVISA LA CANTIDAD REAL A TRASLADAR
            If Not IsNumeric(MskCanRea.Text) Then
               MsgBox "Cantidad Real Incorrecta", vbOKOnly + vbCritical, "Error"
               MskCanRea.SetFocus
               Exit Sub
            End If
            
            'REVISA LA CANTIDAD REAL A TRASLADAR
            If MskCanRea.Text <= 0 Then
               MsgBox "La Cantidad Real No Puede Ser Cero", vbOKOnly + vbCritical, "Error"
               MskCanRea.SetFocus
               Exit Sub
            End If
        
            
            'REVISAMOS LA CANTIDAD DE ENTRADA
            If Not IsNumeric(MskCanSal.Text) Then
               MsgBox "Cantidad De ENTRADA Incorrecta", vbOKOnly + vbCritical, "Error"
               MskCanSal.SetFocus
               Exit Sub
            End If
            
            'REVISAMOS LA CANTIDAD DE DESPERDICIO
            If Not IsNumeric(MskCanDes.Text) Then
               MsgBox "Cantidad De DESPERDICIO Incorrecta", vbOKOnly + vbCritical, "Error"
               MskCanDes.SetFocus
               Exit Sub
            End If
            
            'REVISAMOS LA CANTIDAD DE DESPERDICIO
            If Not IsNumeric(MskCanDes.Text) Then
               MsgBox "Cantidad De DESPERDICIO Incorrecta", vbOKOnly + vbCritical, "Error"
               MskCanDes.SetFocus
               Exit Sub
            End If
            
            'REVISA LA BODEGA A DONDE VA A ENTRAR LA MATERIA PRIMA
            Set RBuscaBodegaEntrada = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaBodegaEntrada, "select * From BodegasInventario Where CodigoBodega = '" & TxtBodEnt.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaBodegaEntrada, "select * From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBodEnt.Text) & "'")
                End If
                If RBuscaBodegaEntrada.RecordCount > 0 Then
                Else
                    MsgBox "Bodega De Entrada No Existe", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
                End If
            
                'VERIFICA BODEGA SALIDA CON BODEGA DE ENTRADA
                If TxtBodSal.Text = TxtBodEnt.Text Then
                    MsgBox "La Bodega De Salida No Puede Ser Igual A La Bodega De Entrada", vbOKOnly + vbExclamation, "Informacion"
                    TxtBodEnt.SetFocus
                    Exit Sub
                End If
                
                'VERIFICA LA UNIDAD DE MEDIDA
                If TxtUniMedSal.Text = "" Then
                    MsgBox "Unidad De Medida Salida No Puede Estar Vacia", vbOKOnly + vbExclamation, "Informacion"
                    TxtUniMedSal.SetFocus
                    Exit Sub
                End If
            
                'SUMA LA CANTIDAD REAL
                VSumaEgresos = Val(MskDifReqCor.Text) + Val(MskCanDes.Text) + Val(MskCanDesPro.Text)
                MskCanRea.Text = ((Val(MskCanSal.Text) + Val(MskDifReqCorMas.Text)) - VSumaEgresos)
            
                If TxtOrd.Text <> "" Then
                        Set RBuscaOrden = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where Documento = '" & TxtOrd.Text & "'")
                            Else
                                Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtOrd.Text) & "'")
                            End If
                                
                                        If RBuscaOrden.RecordCount > 0 Then
                                        Else
                                            MsgBox "Numero De Orden No Existe", vbOKOnly + vbInformation, "Informacion"
                                            Exit Sub
                                        End If
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Conexion.Execute "Update DetalleEntradasInventario Set OrdenProduccion = '" & TxtOrd.Text & "' Where FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Linea = '" & TxtLin.Text & "' and Tarima = " & TxtNumIng.Text & " And FichaTecnica = '" & TxtCodSal.Text & "'"
                        Else 'ORACLE
                            Conexion.Execute "Update DetalleEntradasInventario Set OrdenProduccion = '" & UCase(TxtOrd.Text) & "' Where FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and Tarima = " & TxtNumIng.Text & " And UPPER(FichaTecnica) = '" & UCase(TxtCodSal.Text) & "'"
                        End If
                        
                End If
    
                            If BEditarDetalle = False Then
                                    VTexto = TxtDocDet.Text & ", " ' DOCUMENTO
                                    VTexto = VTexto & TxtNumIng.Text & ", '" 'TARIMA
                                    VTexto = VTexto & TxtCodSal.Text & "', " 'FICHA TECNICA
                                    VTexto = VTexto & MskCanSal.Text & ", '" 'CANTIDAD SALIDA
                                    VTexto = VTexto & TxtBodEnt.Text & "', " 'BODEGA ENTRADA
                                    VTexto = VTexto & MskDifReqCorMas.Text & ", " 'DE +
                                    VTexto = VTexto & MskDifReqCor.Text & ", " 'DE -
                                    VTexto = VTexto & MskCanDes.Text & ", " 'DESPERDICIO
                                    VTexto = VTexto & MskCanDesPro.Text & ", " 'DESÈRDICIO PROVEEDOR
                                    VTexto = VTexto & MskCanRea.Text & ", '" 'CANTIDAD REAL
                                    VTexto = VTexto & TxtOrd.Text & "', " 'ORDEN
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                            VTexto = VTexto & "#" & Format(MskFecPro.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                                    Else 'ORACLE
                                            VTexto = VTexto & "To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & ", '"  'FECHA
                                    End If
                                    VTexto = VTexto & TxtLin.Text & "'" 'LINEA
                                    
                                    Conexion.Execute "Insert Into DetalleTrasladosInventario Values(" & VTexto & ")"
                            Else 'SI ESTA EDITANDO
                                    VTexto = VTexto & "Tarima = " & TxtNumIng.Text & ", "  'TARIMA
                                    VTexto = VTexto & "FichaTecnica = '" & TxtCodSal.Text & "', " 'FICHA
                                    VTexto = VTexto & "CantidadSalida = " & MskCanSal.Text & ", " 'CANTIDAD SALIDA
                                    VTexto = VTexto & "BodegaEntrada = '" & TxtBodEnt.Text & "', " 'BODEGA ENTRADA
                                    VTexto = VTexto & "DiferenciaReqCorMas = " & MskDifReqCorMas.Text & ", " 'DE +
                                    VTexto = VTexto & "DiferenciaReqCor = " & MskDifReqCor.Text & ", " 'DE -
                                    VTexto = VTexto & "CantidadDesperdicio = " & MskCanDes.Text & ", " 'DESPERDICIO
                                    VTexto = VTexto & "CantidadDesperdicioProveedor = " & MskCanDesPro.Text & ", " 'DESPERDICIO PROVEEDOR
                                    VTexto = VTexto & "CantidadReal = " & MskCanRea.Text & ", " 'CANTIDAD REAL
                                    VTexto = VTexto & "Orden = '" & TxtOrd.Text & "', " 'ORDEN
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        VTexto = "Fecha = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "#, " 'FECHA
                                    Else 'ORACLE
                                        VTexto = "Fecha = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                                    End If
                                    VTexto = VTexto & "LineaProduccion = '" & TxtLin.Text & "', " 'LINEA
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        VTexto = VTexto & " Where Documento = " & TxtDocDet.Text & " And FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And LineaProduccion = '" & TxtLin.Text & "' And FichaTecnica = '" & TxtCodSal.Text & "' And Tarima = " & TxtNumIng.Text
                                    Else
                                        VTexto = VTexto & " Where Documento = " & TxtDocDet.Text & " And FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(LineaProduccion) = '" & UCase(TxtLin.Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtCodSal.Text) & "' And Tarima = " & TxtNumIng.Text
                                    End If
                                    
                                    Conexion.Execute "Update DetalleTrasladosInventario Set " & VTexto
                            End If
                                        
                                    'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                        'iI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                Exit Sub
                                        End If
                                    Else 'ORACLE
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err = -2147217873 Then
                                                MsgBox "Tarima/Bulto Ya Existe En Este Documento De Traslado", vbOKOnly + vbInformation, "Informacion"
                                                Exit Sub
                                        ElseIf Err <> -2147217873 And Err <> 0 Then
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                Exit Sub
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
    
    'AGREGA UN REGISTRO
    DataEncabezadoTraslados.Recordset.AddNew
    
    'SI HAY ERRORES
    If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbCritical + vbExclamation, "Error """
        Exit Sub
    End If
    
    BEditar = False
    Bandera = True
    Botones1
    
    'ASIGNA EL USUARIO A EL CAMPO DE REQUERIDO
    TxtEncabezado.Item(1).Text = GUsuario
    MskFec.Text = Date
    MskFec.SetFocus
    'ASIGNA EL ESTADO DE EL TRASLADO
    TxtEncabezado.Item(0).Text = "NO LIBERADO"
        
    'BUSCA EL MAXIMO DE DOCUMENTO Y LE SUMA 1
    Set RBuscaSigDoc = Db.OpenRecordset("Select Max(Documento) from EncabezadoTrasladosInventario")
        If RBuscaSigDoc.RecordCount > 0 Then
            If IsNull(RBuscaSigDoc(0)) Then
                TxtDocTra.Text = "1"
            Else
                TxtDocTra.Text = Val(RBuscaSigDoc(0)) + 1
            End If
        End If

    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False

End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next

    'REVISA EL TIPO DE DOCUMENTO
    If TxtTipDoc.Text = "" Then
        MsgBox "Tipo Documento No Puede Ser Vacio", vbOKOnly + vbInformation, "Informacion"
        TxtTipDoc.SetFocus
        Exit Sub
    End If
    'REVISA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
        MskFec.SetFocus
        Exit Sub
    End If
    
    'OSEA QUE SI ESTA AGREGANDO UN REGISTRO
    If BEditar = False Then
            'BUSCA SI YA EXISTE EL NUMERO DE DOCUMENTO PARA ESTE TIPO DE DOCUMENTO
            Set RBuscaDocumento = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaDocumento, "Select * From EncabezadoTrasladosInventario Where TipoDeDocumento = '" & TxtTipDoc.Text & "' And NumeroDocumento = '" & TxtNumDoc.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaDocumento, "Select * From EncabezadoTrasladosInventario Where UPPER(TipoDeDocumento) = '" & UCase(TxtTipDoc.Text) & "' And UPPER(NumeroDocumento) = '" & UCase(TxtNumDoc.Text) & "'")
                End If
                    If RBuscaDocumento.RecordCount > 0 Then
                        MsgBox "Numero Documento Para Este Tipo De Documento Ya Existe", vbOKOnly + vbInformation, "Informacion"
                        TxtTipDoc.SetFocus
                        Exit Sub
                    End If
    End If

    
    VDocumento = TxtDocTra.Text
    VBodegaSalida = TxtBodSal.Text
        
    'GRABA DATOS
    DataEncabezadoTraslados.Recordset.Update
    
    If Err = 3022 Then
        MsgBox "Documento De Traslado Ya Existe ", vbOKOnly + vbCritical, "Informacion"
        TxtDocTra.SetFocus
        Exit Sub
    ElseIf Err <> 0 And Err <> 3022 Then
        MsgBox "Error Numero " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        Exit Sub
    End If
    
                    If BEditarEncabezado = False Then
                            VTexto = TxtDocTra.Text & ", " 'DOCUMENTO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                            Else 'ORACLE
                                VTexto = VTexto & "To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", '"  'FECHA
                            End If
                            VTexto = VTexto & TxtTipDoc.Text & "', '" 'TIPO DOCUMENTO
                            VTexto = VTexto & TxtNumDoc.Text & "', '" 'NUMERO DOCUMENTO
                            VTexto = VTexto & TxtBodSal.Text & "', '" 'BODEGA SALIDA
                            VTexto = VTexto & TxtEncabezado.Item(1).Text & "', '" 'RQUERIDO
                            VTexto = VTexto & TxtEncabezado.Item(2).Text & "', '" 'LIBERADO
                            VTexto = VTexto & TxtEncabezado.Item(3).Text & "', '" 'OBSERVACIONES
                            VTexto = VTexto & TxtEncabezado.Item(4).Text & "'" 'ESTADO
                            
                            Conexion.Execute "Insert Into EncabezadoTrasladosInventario Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = "Fecha = #" & Format(MskFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = "Fecha = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "TipoDeDocumento = '" & UCase(TxtTipDoc.Text) & "', " 'TIPO DE DOCUMENTO
                            VTexto = VTexto & "NumeroDocumento = '" & UCase(TxtNumDoc.Text) & "', " 'NUMERO DOCUEMNTO
                            VTexto = VTexto & "BodegaSalida = '" & TxtBodSal.Text & "', " 'BODEGA SALIDA
                            VTexto = VTexto & "Requerido = '" & TxtEncabezado.Item(1).Text & "', " 'REQUERIDO
                            VTexto = VTexto & "Liberado = '" & TxtEncabezado.Item(2).Text & "', " 'LIBERADO
                            VTexto = VTexto & "Observaciones = '" & TxtEncabezado.Item(3).Text & "', " 'OBSERVACIONES
                            VTexto = VTexto & "Estado = '" & TxtEncabezado.Item(0).Text & "', " 'ESTADO
                            
                            VTexto = VTexto & " Where Documento = " & VDocumento 'DOCUMENTO
                            
                            Conexion.Execute "UPDATE EncabezadoTrasladosInventario SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Transaccion Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            MousePointer = 0
                            MskFec.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            MousePointer = 0
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Transaccion Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            MousePointer = 0
                            MskFec.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            MousePointer = 0
                            Exit Sub
                        End If
                    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
                    Set REncabezado = New ADODB.Recordset
                        Call Abrir_Recordset(REncabezado, "Select * From EncabezadoTrasladosInventario Where Documento = " & VDocumento & " Order By Documento")
                        
                        Llena_CamposEncabezado
                        
                        Set RDetalle = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                                            Else 'ORACLE
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                                            End If
                                                Llena_CamposDetalle
                                                Set DbGridDetalle.DataSource = RDetalle
                
                      
    'VIZUALIZA EL DETALLE DE TRASLADO
    FrameDetalle.Visible = True
    
    'VISUALIZA TODOS LOS BOTONES DE DETALLE
    Bandera3 = True
    BotonesDetalleVisibles
    
    'NO VISUALIZA EL DATA DE ENCABEZADO DE TRASLADOS
    'ESCONDE EL DATA
    CmdBotones2.Item(1).Visible = False
    CmdBotones2.Item(2).Visible = False
    CmdBotones2.Item(3).Visible = False
    CmdBotones2.Item(4).Visible = False
    
    
    'ESCONDE LOS BOTONES DEL ENCABEZADO
    Bandera4 = False
    BotonesEncabezadoVisibles
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = True
    FrameDetalle.Visible = True
    FrameEncabezado.Enabled = False
    DbGridDetalle.Visible = True
    CmdAgregar2.SetFocus
End Sub

Private Sub CmdImprimir_Click()
MousePointer = 11
                    'MUESTRA EL REPORTE
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioTraslados.rpt"
                    Else
                        GNombreReporte = "InventarioTrasladosO.rpt"
                    End If
                    GCriteriaReporte = "{EncabezadoTrasladosInventario.Documento} = " & TxtDocTra.Text
                    FrmReporte.Show 1
        
MousePointer = 0

End Sub

Private Sub CmdImprimirCedula_Click()
On Error Resume Next
MousePointer = 11
                    'MUESTRA EL REPORTE
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "CedulaMateriaPrimaTraslados.rpt"
                    Else
                        GNombreReporte = "CedulaMateriaPrimaTrasladosO.rpt"
                    End If
                    GCriteriaReporte = "{DetalleEntradasInventario.FechaProduccion} = " & MskFecPro.Text & " And {DetalleEntradasInventario.LineaProduccion} = '" & TxtLin.Text & "' And {DetalleEntradasInventario.FichaTecnica} = '" & TxtCodSal.Text & "' And {DetalleEntradasInventario.Tarima} = " & TxtNumIng.Text
                    FrmReporte.Show 1
                                    
                
MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdTerminar_Click()
If CmdCancelar2.Enabled = True Then
     CmdCancelar2_Click
End If
    
     
    'VISUALIZA EL DATA DE ENCABEZADO DE TRASLADOS
    CmdBotones2.Item(1).Visible = True
    CmdBotones2.Item(2).Visible = True
    CmdBotones2.Item(3).Visible = True
    CmdBotones2.Item(4).Visible = True
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = False
    'FrameDetalle.Visible = False
    FrameEncabezado.Enabled = True
    
    'VISUALIZA TODOS LOS BOTONES DE DETALLE
    Bandera3 = False
    BotonesDetalleVisibles
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesEncabezadoVisibles

    Set REncabezado = New ADODB.Recordset
            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoTrasladosInventario Order By Documento")
            REncabezado.MoveLast
                Llena_CamposEncabezado
    
                                    Set RDetalle = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                                            Else 'ORACLE
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                                            End If
                                                Llena_CamposDetalle
                                                Set DbGridDetalle.DataSource = RDetalle

End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub


Private Sub DBGridBuscar_DblClick()
    'BODEGA ENTRADA
    If BBodegaEntrada = True Then
        TxtBodEnt.Text = DbGridBuscar.Columns(0)
        TxtBodEnt.SetFocus
    'MATERIA PRIMA SALIDA
    ElseIf BCodigoSalida = True Then
        TxtCodSal.Text = DbGridBuscar.Columns(0)
        TxtCodSal.SetFocus
    'NUMERO INGRESO
    ElseIf BNumeroIngreso = True Then
        TxtNumIng.Text = DbGridBuscar.Columns(2)
        TxtNumIng.SetFocus
    'DOCUMENTO
    ElseIf BDocumento = True Then
        TxtTipDoc.Text = DbGridBuscar.Columns(0)
        TxtTipDoc.SetFocus
    'BODEGA SALIDA
    ElseIf BBodegaSalida = True Then
        TxtBodSal.Text = DbGridBuscar.Columns(0)
        TxtBodSal.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
    'BODEGA ENTRADA
    If BBodegaEntrada = True Then
        TxtBodEnt.Text = DbGridBuscar.Columns(0)
        TxtBodEnt.SetFocus
    'MATERIA PRIMA SALIDA
    ElseIf BCodigoSalida = True Then
        TxtCodSal.Text = DbGridBuscar.Columns(0)
        TxtCodSal.SetFocus
    'NUMERO INGRESO
    ElseIf BNumeroIngreso = True Then
        TxtNumIng.Text = DbGridBuscar.Columns(2)
        TxtNumIng.SetFocus
    'DOCUMENTO
    ElseIf BDocumento = True Then
        TxtTipDoc.Text = DbGridBuscar.Columns(0)
        TxtTipDoc.SetFocus
    'BODEGA SALIDA
    ElseIf BBodegaSalida = True Then
        TxtBodSal.Text = DbGridBuscar.Columns(0)
        TxtBodSal.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
End Sub

Private Sub Form_Load()
        Set REncabezado = New ADODB.Recordset
            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoTrasladosInventario Order By Documento")
                Llena_CamposEncabezado
        
                    Set RDetalle = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                        Else 'ORACLE
                            Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.LineaProduccion, D.FichaTecnica, D.Tarima, D.Orden, D.CantidadSalida, D.BodegaEntrada, D.DiferenciaReqCorMas, D.DiferenciaReqCor, D.CantidadDesperdicio, D.CantidadDesperdicioProveedor, D.CantidadReal From EncabezadoTrasladosInventario E, DetalleTrasladosInventario D Where E.Documento = " & TxtDocTra.Text & " And E.Documento = D.Documento")
                        End If
                            Llena_CamposDetalle
                            Set DbGridDetalle.DataSource = RDetalle
                
        
End Sub



Private Sub MskCanDes_GotFocus()
    MskCanDes.SelStart = 0
    MskCanDes.SelLength = Len(MskCanDes.Text)
End Sub

Private Sub MskCanDes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskCanDes_LostFocus()
'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES
        If ChkLam.Value = 1 Then
            If IsNumeric(MskCanDes.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtCodSal.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodSal.Text) & "'")
                        End If
                        
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                            
                        MskCanDes.Text = MskCanDes.Text * VUnidadesxLamina
            End If
                            
                        
        End If

End Sub

Private Sub MskCanDesPro_GotFocus()
    MskCanDesPro.SelStart = 0
    MskCanDesPro.SelLength = Len(MskCanDesPro.Text)
End Sub

Private Sub MskCanDesPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskCanDesPro_LostFocus()
    
    
    'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES
        If ChkLam.Value = 1 Then
            If IsNumeric(MskCanDesPro.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtCodSal.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodSal.Text) & "'")
                        End If
                        
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                        
                            
                        MskCanDesPro.Text = MskCanDesPro.Text * VUnidadesxLamina
            End If
                            
                        
        End If
        
        
            VSumaEgresos = Val(MskDifReqCor.Text) + Val(MskCanDes.Text) + Val(MskCanDesPro.Text)
            MskCanRea.Text = ((Val(MskCanSal.Text) + Val(MskDifReqCorMas.Text)) - VSumaEgresos)


            If IsNumeric(MskCanRea.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtCodSal.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodSal.Text) & "'")
                        End If
                        
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                        
                            
                        If VUnidadesxLamina > 0 Then
                            TxtLamEnt.Text = MskCanRea / VUnidadesxLamina
                        Else
                        End If
                End If

End Sub



Private Sub MskCanRea_GotFocus()
    MskCanRea.SelStart = 0
    MskCanRea.SelLength = Len(MskCanRea.Text)
End Sub

Private Sub MskCanRea_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskCanSal_GotFocus()
    MskCanSal.SelStart = 0
    MskCanSal.SelLength = Len(MskCanSal.Text)
End Sub

Private Sub MskCanSal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskDifReqCor_GotFocus()
    MskDifReqCor.SelStart = 0
    MskDifReqCor.SelLength = Len(MskDifReqCor.Text)
End Sub

Private Sub MskDifReqCor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub MskDifReqCor_LostFocus()
    'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES
        If ChkLam.Value = 1 Then
            If IsNumeric(MskDifReqCor.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtCodSal.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodSal.Text) & "'")
                        End If
                        
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                            
                        MskDifReqCor.Text = MskDifReqCor.Text * VUnidadesxLamina
            End If
        End If

End Sub

Private Sub MskDifReqCorMas_GotFocus()
    MskDifReqCorMas.SelStart = 0
    MskDifReqCorMas.SelLength = Len(MskDifReqCorMas.Text)
End Sub

Private Sub MskDifReqCorMas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub MskDifReqCorMas_LostFocus()
        'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES
        If ChkLam.Value = 1 Then
            If IsNumeric(MskDifReqCorMas.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtCodSal.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodSal.Text) & "'")
                        End If
                        
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                        
                            
                        MskDifReqCorMas.Text = MskDifReqCorMas.Text * VUnidadesxLamina
            End If
                            
                        
        End If

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


Private Sub TxtBodEnt_Change()
    'REVISA LA BODEGA A DONDE VA A ENTRAR LA MATERIA PRIMA
            Set RBuscaBodegaEntrada = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaBodegaEntrada, "select * From BodegasInventario Where CodigoBodega = '" & TxtBodEnt.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaBodegaEntrada, "select * From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBodEnt.Text) & "'")
                End If
                If RBuscaBodegaEntrada.RecordCount > 0 Then
                    LblBodegaEntrada.Caption = RBuscaBodegaEntrada!Descripcion
                Else
                    LblBodegaEntrada.Caption = ""
                End If
    
End Sub

Private Sub TxtBodEnt_DblClick()
        BBodegaEntrada = True
        BCodigoSalida = False
        BNumeroIngreso = False
        BDocumento = False
        BBodegaSalida = False
        Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
End Sub

Private Sub TxtBodEnt_GotFocus()
    TxtBodEnt.SelStart = 0
    TxtBodEnt.SelLength = Len(TxtBodEnt.Text)
End Sub

Private Sub TxtBodEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        BBodegaEntrada = True
        BCodigoSalida = False
        BNumeroIngreso = False
        BDocumento = False
        BBodegaSalida = False
        Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
    End If
End Sub

Private Sub TxtBodSal_Change()
    'REVISA LA BODEGA A DONDE VA A ENTRAR LA MATERIA PRIMA
    Set RBuscaBodegaSalida = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBusqueda, "select * From BodegasInventario Where CodigoBodega = '" & TxtBodSal.Text & "'")
        Else
            Call Abrir_Recordset(RBusqueda, "select * From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBodSal.Text) & "'")
        End If
        If RBuscaBodegaSalida.RecordCount > 0 Then
            LblBodegaSalida.Caption = RBuscaBodegaSalida!Descripcion
        Else
            LblBodegaSalida.Caption = ""
        End If

End Sub

Private Sub TxtBodSal_DblClick()
        BBodegaEntrada = False
        BCodigoSalida = False
        BNumeroIngreso = False
        BDocumento = False
        BBodegaSalida = True
        Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                Set DbGridBuscar.DataSource = RBusqueda
                DbGridBuscar.Columns(1).Width = "4000"
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus

End Sub

Private Sub TxtBodSal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BBodegaEntrada = False
            BCodigoSalida = False
            BNumeroIngreso = False
            BDocumento = False
            BBodegaSalida = True
            Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                Set DbGridBuscar.DataSource = RBusqueda
                DbGridBuscar.Columns(1).Width = "4000"
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
        End If

End Sub

Private Sub Txtbuscar_Change()
    Set RBusqueda = New ADODB.Recordset
    'BODEGA ENTRADA
   If BBodegaEntrada = True Or BBodegaSalida = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion from BodegasInventario Where CodigoBodega Like '%" & TxtBuscar.Text & "%' Order by CodigoBodega")
            Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion from BodegasInventario Where UPPER(CodigoBodega) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoBodega")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion from BodegasInventario Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by CodigoBodega")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion from BodegasInventario Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoBodega")
            End If
            
        End If
        
    'CODIGO MATERIA PRIMA SALIDA
    ElseIf BCodigoSalida = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Esp_Tec Like '%" & TxtBuscar.Text & "%' Order by Esp_Tec")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Esp_Tec")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Descrip Like '%" & TxtBuscar.Text & "%' Order by Esp_Tec")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Esp_Tec")
            End If
        End If
    'CODIGO DOCUMENTO
    ElseIf BDocumento = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where CodigoDocumento Like '%" & TxtBuscar.Text & "%' Order by CodigoDocumento")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where UPPER(CodigoDocumento) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoDocumento")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by CodigoDocumento")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoDocumento")
            End If
        End If
    End If
    
    If BBodegaEntrada = True Or BDocumento = True Or BCodigoSalida = True Then
        Set DbGridBuscar.DataSource = RBusqueda
        DbGridBuscar.Columns(1).Width = "4000"
    End If

End Sub


Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub




Private Sub TxtCodSal_Change()
        'BUSCA LA MATERIA PRIMA DE ACUERDO A LA BODEGA DE SALIDA
        Set RBuscaMateriaPrimaSalida = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaMateriaPrimaSalida, "Select Descripcion, UnidadMedida From FichaTecnica Where Esp_Tec = '" & TxtCodSal.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaMateriaPrimaSalida, "Select Descripcion, UnidadMedida From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodSal.Text) & "'")
            End If
        If RBuscaMateriaPrimaSalida.RecordCount > 0 Then
            LblDesSal.Caption = RBuscaMateriaPrimaSalida!Descripcion
            TxtUniMedSal.Text = RBuscaMateriaPrimaSalida!unidadMedida
        Else
            LblDesSal.Caption = ""
            TxtUniMedSal.Text = ""
        End If
End Sub

Private Sub TxtCodSal_DblClick()
        BBodegaEntrada = False
        BCodigoSalida = True
        BNumeroIngreso = False
        BDocumento = False
        BBodegaSalida = False
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
        Set DbGridBuscar.DataSource = RBusqueda
        DbGridBuscar.Columns(1).Width = "4000"
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
End Sub

Private Sub TxtCodSal_GotFocus()
    TxtCodSal.SelStart = 0
    TxtCodSal.SelLength = Len(TxtCodSal.Text)
End Sub

Private Sub TxtCodSal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        BBodegaEntrada = False
        BCodigoSalida = True
        BNumeroIngreso = False
        BDocumento = False
        BBodegaSalida = False
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
    End If
End Sub

Private Sub TxtDocTra_GotFocus()
            TxtDocTra.SelStart = 0
            TxtDocTra.SelLength = Len(TxtDocTra.Text)
End Sub

Private Sub TxtDocTra_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If
End Sub

Private Sub TxtEncabezado_GotFocus(Index As Integer)
            TxtEncabezado.Item(Index).SelStart = 0
            TxtEncabezado.Item(Index).SelLength = Len(TxtEncabezado.Item(Index).Text)
End Sub

Private Sub TxtEncabezado_KeyPress(Index As Integer, KeyAscii As Integer)
            If KeyAscii = 13 Then
               SendKeys "{tab}"
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


Private Sub TxtNumIng_GotFocus()
    TxtNumIng.SelStart = 0
    TxtNumIng.SelLength = Len(TxtNumIng.Text)
End Sub

Private Sub TxtNumIng_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub TxtNumIng_LostFocus()
On Error Resume Next
        If IsNumeric(TxtNumIng.Text) Then
            'BUSCA EL NUMERO DE INGRESO Y ASIGNA LA BODEGA, CODIGO Y CANTIDAD DE ACUERDO COMO ENTRO A LA BODEGA
            Set RBuscaTarima = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTarima, "Select D.Saldo, D.FichaTecnica, B.Descripcion, D.OrdenProduccion From DetalleEntradasInventario D, BodegasInventario B Where D.FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And D.Linea = '" & TxtLin.Text & "' And D.FichaTecnica = '" & TxtCodSal.Text & "' And D.Tarima = " & TxtNumIng.Text & " And D.Bodega = B.CodigoBodega")
                Else
                    Call Abrir_Recordset(RBuscaTarima, "Select D.Saldo, D.FichaTecnica, B.Descripcion, D.OrdenProduccion From DetalleEntradasInventario D, BodegasInventario B Where D.FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(D.Linea) = '" & UCase(TxtLin.Text) & "' And UPPER(D.FichaTecnica) = '" & UCase(TxtCodSal.Text) & "' And D.Tarima = " & TxtNumIng.Text & " And UPPER(D.Bodega) = UPPER(B.CodigoBodega)")
                End If
            'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
            If RBuscaTarima.RecordCount > 0 Then
                MskCanSal.Text = RBuscaTarima(0)
                TxtCodSal.Text = RBuscaTarima(1)
                'SI DIGITAN LA ORDEN
                If TxtOrd.Text <> "" Then
                Else
                    TxtOrd.Text = RBuscaTarima(3)
                End If
                LblBodega.Caption = RBuscaTarima(2)
                
            'SI NO ENCUENTRA EL INGRESO DEJA EN BLANCO
                        'SI ESTA CHEQUEADO EL CHK DE LAMINAS A UNIDADES

              If IsNumeric(MskCanSal.Text) Then
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select CuerposPorLamina From FichaTecnica Where Esp_Tec = '" & TxtCodSal.Text & "'")
                        Else
                            Call Abrir_Recordset(RBuscaFicha, "Select CuerposPorLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodSal.Text) & "'")
                        End If
                        If RBuscaFicha.RecordCount > 0 Then
                              VUnidadesxLamina = RBuscaFicha(0)
                        Else
                              VUnidadesxLamina = 0
                        End If
                        
                        If VUnidadesxLamina > 0 Then
                            TxtLamReq.Text = MskCanSal.Text / VUnidadesxLamina
                        Else
                        End If
                End If
                            
                        


            
            Else
                MskCanSal.Text = 0
                TxtCodSal.Text = ""
                'TxtBodSal.Text = ""
                TxtBodEnt.Text = ""
            End If
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
                Set RBuscaFichaOrden = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFichaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtOrd.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaFichaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtOrd.Text) & "'")
                    End If
                    If RBuscaFichaOrden.RecordCount > 0 Then
                        TxtCodSal.Text = RBuscaFichaOrden!FichaTecnica
                    Else
                            
                    End If
            
            
End Sub

Private Sub TxtTipDoc_Change()
            Set RBuscaTipoDocumento = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTipoDocumento, "Select Descripcion from Documentos Where CodigoDocumento = '" & TxtTipDoc.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaTipoDocumento, "Select Descripcion from Documentos Where UPPER(CodigoDocumento) = '" & UCase(TxtTipDoc.Text) & "'")
                End If
                If RBuscaTipoDocumento.RecordCount > 0 Then
                    LblTipDoc.Caption = RBuscaTipoDocumento!Descripcion
                Else
                    LblTipDoc.Caption = ""
                End If
End Sub

Private Sub TxtTipDoc_DblClick()
            BBodegaEntrada = False
            BCodigoSalida = False
            BNumeroIngreso = False
            BDocumento = True
            BBodegaSalida = False
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion From Documentos")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            
End Sub

Private Sub TxtTipDoc_GotFocus()
            TxtTipDoc.SelStart = 0
            TxtTipDoc.SelLength = Len(TxtTipDoc.Text)
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
            
    If KeyAscii = 13 Then
            SendKeys "{TAB}"
    End If
       
    If KeyAscii = 43 Then
            BBodegaEntrada = False
            BCodigoSalida = False
            BNumeroIngreso = False
            BDocumento = True
            BBodegaSalida = False
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion From Documentos")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
    End If

End Sub

Private Sub TxtUniMedSal_GotFocus()
    TxtUniMedSal.SelStart = 0
    TxtUniMedSal.SelLength = Len(TxtUniMedSal.Text)
End Sub

Private Sub TxtUniMedSal_KeyPress(KeyAscii As Integer)
    'SI PRECIONA ENTER
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub


Sub Columnas()
    DbGridBuscar.Columns(0).Caption = "Bodega"
    DbGridBuscar.Columns(0).Width = "500"
    DbGridBuscar.Columns(1).Caption = "Descripcion"
    DbGridBuscar.Columns(1).Width = "3000"
    DbGridBuscar.Columns(2).Caption = "# Bulto"
    DbGridBuscar.Columns(2).Width = "1000"
    DbGridBuscar.Columns(3).Caption = "Inicio"
    DbGridBuscar.Columns(4).Caption = "Salidas"
    DbGridBuscar.Columns(5).Caption = "Existencia"
    
End Sub


Public Sub BotonesEncabezadoVisibles()
    If Bandera4 = True Then
         CmdAgregar.Visible = True
         CmdEditar.Visible = True
         CmdGrabar.Visible = True
         CmdBorrar.Visible = True
         CmdCancelar.Visible = True
         CmdBuscar.Visible = True
         CmdBuscar2.Visible = True
         CmdSalida.Visible = True
         CmdImprimir.Visible = True
    ElseIf Bandera4 = False Then
         CmdAgregar.Visible = False
         CmdEditar.Visible = False
         CmdGrabar.Visible = False
         CmdBorrar.Visible = False
         CmdCancelar.Visible = False
         CmdBuscar.Visible = False
         CmdBuscar2.Visible = False
         CmdSalida.Visible = False
         CmdImprimir.Visible = False
    End If
    
End Sub

Public Sub Llena_CamposEncabezado()
On Error Resume Next
            If REncabezado.RecordCount > 0 Then
                If IsNull(REncabezado!Documento) Then
                    TxtDocTra.Text = ""
                Else
                    TxtDocTra.Text = REncabezado!Documento
                End If
                If IsNull(REncabezado!fecha) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = REncabezado!fecha
                End If
                If IsNull(REncabezado!TipoDeDocumento) Then
                    TxtTipDoc.Text = ""
                Else
                    TxtTipDoc.Text = REncabezado!TipoDeDocumento
                End If
                If IsNull(REncabezado!NumeroDocumento) Then
                    TxtNumDoc.Text = ""
                Else
                    TxtNumDoc.Text = REncabezado!NumeroDocumento
                End If
                If IsNull(REncabezado!BodegaSalida) Then
                    TxtBodSal.Text = ""
                Else
                    TxtBodSal.Text = REncabezado!BodegaSalida
                End If
                If IsNull(REncabezado!Requerido) Then
                    TxtEncabezado.Item(1).Text = ""
                Else
                    TxtEncabezado.Item(1).Text = REncabezado!Requerido
                End If
                If IsNull(REncabezado!Liberado) Then
                    TxtEncabezado.Item(2).Text = ""
                Else
                    TxtEncabezado.Item(2).Text = REncabezado!Liberado
                End If
                If IsNull(REncabezado!Obsercaciones) Then
                    TxtEncabezado.Item(3).Text = ""
                Else
                    TxtEncabezado.Item(3).Text = REncabezado!Observaciones
                End If
                If IsNull(REncabezado!Estado) Then
                    TxtEncabezado.Item(0).Text = ""
                Else
                    TxtEncabezado.Item(0).Text = REncabezado!Observaciones
                End If
            Else
                TxtDocTra.Text = ""
                MskFec.Text = ""
                TxtTipDoc.Text = ""
                TxtNumDoc.Text = ""
                TxtBodSal.Text = ""
                TxtEncabezado.Item(0).Text = ""
                TxtEncabezado.Item(1).Text = ""
                TxtEncabezado.Item(2).Text = ""
                TxtEncabezado.Item(3).Text = ""
                
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
                If IsNull(RDetalle!Tarima) Then
                    TxtNumIng.Text = ""
                Else
                    TxtNumIng.Text = RDetalle!Tarima
                End If
                If IsNull(RDetalle!FichaTecnica) Then
                    TxtcodsalText = ""
                Else
                    TxtCodSal.Text = RDetalle!FichaTecnica
                End If
                If IsNull(RDetalle!CantidadSalida) Then
                    MskCanSal.Text = ""
                Else
                    MskCanSal.Text = RDetalle!CantidadSalida
                End If
                If IsNull(RDetalle!BodegaEntrada) Then
                    TxtBodEnt.Text = 0
                Else
                    TxtBodEnt.Text = RDetalle!BodegaEntrada
                End If
                If IsNull(RDetalle!DiferenciaReqCorMas) Then
                    MskDifReqCorMas.Text = 0
                Else
                    MskDifReqCorMas.Text = RDetalle!DiferenciaReqCorMas
                End If
                If IsNull(RDetalle!DiferenciaReqCor) Then
                    MskDifReqCor.Text = ""
                Else
                    MskDifReqCor.Text = RDetalle!DiferenciaReqCor
                End If
                If IsNull(RDetalle!CantidadDesperdicio) Then
                    MskCanDes.Text = ""
                Else
                    MskCanDes.Text = RDetalle!CantidadDesperdicio
                End If
                If IsNull(RDetalle!CantidadDesperdicioProveedor) Then
                    MskCanDesPro.Text = ""
                Else
                    MskCanDesPro.Text = RDetalle!CantidadDesperdicioProveedor
                End If
                If IsNull(RDetalle!CantidadReal) Then
                    MskCanRea.Text = ""
                Else
                    MskCanRea.Text = RDetalle!CantidadReal
                End If
                If IsNull(RDetalle!Orden) Then
                    TxtOrd.Text = ""
                Else
                    TxtOrd.Text = RDetalle!Orden
                End If
                If IsNull(RDetalle!FechaProduccion) Then
                    MskFecPro.Text = ""
                Else
                    MskFecPro.Text = RDetalle!FechaProduccion
                End If
                If IsNull(RDetalle!LineaProduccion) Then
                    TxtLin.Text = ""
                Else
                    TxtLin.Text = RDetalle!LineaProduccion
                End If
            Else
                TxtDocDet.Text = ""
                TxtNumIng.Text = ""
                TxtCodSal.Text = ""
                MskCanSal.Text = ""
                TxtBodEnt.Text = ""
                MskDifReqCorMas.Text = 0
                MskDifReqCor.Text = 0
                MskCanDes.Text = 0
                MskCanDesPro.Text = 0
                MskCanRea.Text = 0
                TxtOrd.Text = ""
                MskFecPro.Text = ""
                TxtLin.Text = ""
            End If
            
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Limpia_CamposEncabezado()
                TxtDocTra.Text = ""
                MskFec.Text = ""
                TxtTipDoc.Text = ""
                TxtNumDoc.Text = ""
                TxtBodSal.Text = ""
                TxtEncabezado.Item(0).Text = ""
                TxtEncabezado.Item(1).Text = ""
                TxtEncabezado.Item(2).Text = ""
                TxtEncabezado.Item(3).Text = ""
End Sub

Public Sub Limpia_CamposDetalle()
                TxtDocDet.Text = ""
                TxtNumIng.Text = ""
                TxtCodSal.Text = ""
                MskCanSal.Text = ""
                TxtBodEnt.Text = ""
                MskDifReqCorMas.Text = 0
                MskDifReqCor.Text = 0
                MskCanDes.Text = 0
                MskCanDesPro.Text = 0
                MskCanRea.Text = 0
                TxtOrd.Text = ""
                MskFecPro.Text = ""
                TxtLin.Text = ""
End Sub





