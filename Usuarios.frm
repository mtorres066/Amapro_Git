VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Usuarios 
   Caption         =   "Mantenimiento de Usuarios"
   ClientHeight    =   8205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9555
   Icon            =   "Usuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   9555
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   8175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   14420
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "Usuarios.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameAgencias"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdEditar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdAgregar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdGrabar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdCancelar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdBorrar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "CmdSalida"
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
      TabPicture(1)   =   "Usuarios.frx":075C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGrid1"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   1
         Left            =   120
         MouseIcon       =   "Usuarios.frx":0BAE
         Picture         =   "Usuarios.frx":0FF0
         Style           =   1  'Graphical
         TabIndex        =   73
         ToolTipText     =   "Primer Registro"
         Top             =   7440
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   2
         Left            =   480
         MouseIcon       =   "Usuarios.frx":1522
         Picture         =   "Usuarios.frx":1964
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Registro Anterior"
         Top             =   7440
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   3
         Left            =   8640
         MouseIcon       =   "Usuarios.frx":1E96
         Picture         =   "Usuarios.frx":22D8
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Siguiente Registro"
         Top             =   7440
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   4
         Left            =   9000
         MouseIcon       =   "Usuarios.frx":280A
         Picture         =   "Usuarios.frx":2C4C
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Ultimo Registro"
         Top             =   7440
         Width           =   375
      End
      Begin MSDataGridLib.DataGrid DbGrid1 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   69
         Top             =   720
         Width           =   9255
         _ExtentX        =   16325
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
      Begin VB.CommandButton CmdSalida 
         Caption         =   "&Salida"
         Height          =   700
         Left            =   7560
         MouseIcon       =   "Usuarios.frx":317E
         Picture         =   "Usuarios.frx":35C0
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   7320
         Width           =   960
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "B&orrar"
         Height          =   700
         Left            =   6240
         MouseIcon       =   "Usuarios.frx":5632
         Picture         =   "Usuarios.frx":5A74
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   7320
         Width           =   1200
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   700
         Left            =   4920
         MouseIcon       =   "Usuarios.frx":5FA6
         Picture         =   "Usuarios.frx":63E8
         Style           =   1  'Graphical
         TabIndex        =   62
         Top             =   7320
         Width           =   1200
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   700
         Left            =   3600
         MouseIcon       =   "Usuarios.frx":691A
         Picture         =   "Usuarios.frx":6D5C
         Style           =   1  'Graphical
         TabIndex        =   61
         Top             =   7320
         Width           =   1200
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   700
         Left            =   960
         MouseIcon       =   "Usuarios.frx":728E
         Picture         =   "Usuarios.frx":76D0
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   7320
         Width           =   1200
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "Editar"
         Height          =   700
         Left            =   2280
         Picture         =   "Usuarios.frx":7C02
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   7320
         Width           =   1200
      End
      Begin VB.Frame FrameAgencias 
         Enabled         =   0   'False
         Height          =   6495
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   9195
         Begin VB.CheckBox Check45 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Reclamos Proveedor"
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
            Left            =   5640
            TabIndex        =   77
            Top             =   5880
            Width           =   2265
         End
         Begin VB.CheckBox Check44 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Captura Desperdicio"
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
            Left            =   5640
            TabIndex        =   76
            Top             =   5640
            Width           =   2265
         End
         Begin VB.CheckBox Check43 
            BackColor       =   &H00FF8080&
            Caption         =   "Inspeccion"
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
            Left            =   3120
            TabIndex        =   75
            Top             =   1800
            Width           =   1545
         End
         Begin VB.CheckBox Check42 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Reportes Formatos"
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
            Left            =   5640
            TabIndex        =   74
            Top             =   5400
            Width           =   2265
         End
         Begin VB.CheckBox Check35 
            BackColor       =   &H00FF8080&
            Caption         =   "Consulta Transito"
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
            Left            =   3120
            TabIndex        =   68
            Top             =   4440
            Width           =   2265
         End
         Begin VB.CheckBox Check38 
            BackColor       =   &H0080C0FF&
            Caption         =   "Captura Faltas"
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
            Left            =   5640
            TabIndex        =   67
            Top             =   1560
            Width           =   1770
         End
         Begin VB.CheckBox Check39 
            BackColor       =   &H0080C0FF&
            Caption         =   "Captura Cursos"
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
            Left            =   5640
            TabIndex        =   66
            Top             =   1800
            Width           =   1770
         End
         Begin VB.CheckBox Check40 
            BackColor       =   &H0080C0FF&
            Caption         =   "Captura Aumentos"
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
            Left            =   5640
            TabIndex        =   65
            Top             =   2040
            Width           =   2010
         End
         Begin VB.CheckBox Check14 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ordenes Produccion y Pasadas"
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
            Left            =   3120
            TabIndex        =   57
            Top             =   5160
            Width           =   2145
         End
         Begin VB.CheckBox Check21 
            BackColor       =   &H0000C000&
            Caption         =   "Editar Pedidos"
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
            Left            =   5640
            TabIndex        =   56
            Top             =   4200
            Width           =   2265
         End
         Begin VB.CheckBox Check22 
            BackColor       =   &H0000C000&
            Caption         =   "Borrar Pedidos"
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
            Left            =   5640
            TabIndex        =   55
            Top             =   4440
            Width           =   2265
         End
         Begin VB.CheckBox Check9 
            BackColor       =   &H000080FF&
            Caption         =   "Editar Captura Paros"
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
            Left            =   360
            TabIndex        =   54
            Top             =   3720
            Width           =   2136
         End
         Begin VB.CheckBox check10 
            BackColor       =   &H000080FF&
            Caption         =   "Borrar Captura Paros"
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
            Left            =   360
            TabIndex        =   53
            Top             =   3960
            Width           =   2136
         End
         Begin VB.CheckBox Check41 
            BackColor       =   &H0080C0FF&
            Caption         =   "Reportes"
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
            Left            =   5640
            TabIndex        =   51
            Top             =   2280
            Width           =   1410
         End
         Begin VB.CheckBox Check37 
            BackColor       =   &H0080C0FF&
            Caption         =   "Configuracion"
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
            Left            =   5640
            TabIndex        =   50
            Top             =   1320
            Width           =   2136
         End
         Begin VB.CheckBox Check16 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Reportes"
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
            Left            =   3120
            TabIndex        =   23
            Top             =   6120
            Width           =   2136
         End
         Begin VB.CheckBox Check30 
            BackColor       =   &H00FF8080&
            Caption         =   "Liberacion Traslados"
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
            Left            =   3120
            TabIndex        =   49
            Top             =   3240
            Width           =   2148
         End
         Begin VB.CheckBox Check15 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Inventario y Ventas y Reporte Ejecutivo"
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
            Left            =   3120
            TabIndex        =   20
            Top             =   5640
            Width           =   2175
         End
         Begin VB.TextBox TxtCon 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7920
            TabIndex        =   47
            Top             =   600
            Width           =   975
         End
         Begin MSMask.MaskEdBox MskFecUltAcc 
            Height          =   285
            Left            =   6480
            TabIndex        =   45
            Top             =   600
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.CheckBox Check32 
            BackColor       =   &H00FF8080&
            Caption         =   "Consultas y Graficas"
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
            Left            =   3120
            TabIndex        =   44
            Top             =   3720
            Width           =   2145
         End
         Begin VB.CheckBox Check31 
            BackColor       =   &H00FF8080&
            Caption         =   "Liberacion Salidas"
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
            Left            =   3120
            TabIndex        =   43
            Top             =   3480
            Width           =   2265
         End
         Begin VB.CheckBox Check27 
            BackColor       =   &H00FF8080&
            Caption         =   "Cambios Ubicacion"
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
            Left            =   3120
            TabIndex        =   42
            Top             =   2520
            Width           =   2145
         End
         Begin MSMask.MaskEdBox MskFecAlt 
            Height          =   285
            Left            =   6480
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   240
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.CheckBox Check34 
            BackColor       =   &H00FF8080&
            Caption         =   "Captura Transito"
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
            Left            =   3120
            TabIndex        =   39
            Top             =   4200
            Width           =   2265
         End
         Begin VB.CheckBox Check36 
            BackColor       =   &H00C0E0FF&
            Caption         =   "% Confor. Entrada Inv."
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
            Left            =   5640
            TabIndex        =   38
            Top             =   5160
            Width           =   2265
         End
         Begin VB.CheckBox Check6 
            BackColor       =   &H000080FF&
            Caption         =   "Configuracion"
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
            Left            =   360
            TabIndex        =   19
            Top             =   3000
            Width           =   1935
         End
         Begin VB.CheckBox Check23 
            BackColor       =   &H00FF8080&
            Caption         =   "Configuracion"
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
            Left            =   3120
            TabIndex        =   34
            Top             =   1320
            Width           =   1785
         End
         Begin VB.CheckBox Check24 
            BackColor       =   &H00FF8080&
            Caption         =   "Entradas"
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
            Left            =   3120
            TabIndex        =   33
            Top             =   1560
            Width           =   2145
         End
         Begin VB.CheckBox Check29 
            BackColor       =   &H00FF8080&
            Caption         =   "Liberacion Entradas"
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
            Left            =   3120
            TabIndex        =   32
            Top             =   3000
            Width           =   2145
         End
         Begin VB.CheckBox Check28 
            BackColor       =   &H00FF8080&
            Caption         =   "Cierre Bulto/Tarima"
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
            Left            =   3120
            TabIndex        =   31
            Top             =   2760
            Width           =   2145
         End
         Begin VB.CheckBox Check26 
            BackColor       =   &H00FF8080&
            Caption         =   "Salidas"
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
            Left            =   3120
            TabIndex        =   30
            Top             =   2280
            Width           =   1545
         End
         Begin VB.CheckBox Check25 
            BackColor       =   &H00FF8080&
            Caption         =   "Traslados"
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
            Left            =   3120
            TabIndex        =   29
            Top             =   2040
            Width           =   2145
         End
         Begin VB.CheckBox Check33 
            BackColor       =   &H00FF8080&
            Caption         =   "Reportes"
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
            Left            =   3120
            TabIndex        =   28
            Top             =   3960
            Width           =   1545
         End
         Begin VB.CheckBox Check17 
            BackColor       =   &H0000C000&
            Caption         =   "Pedidos de Clientes"
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
            Left            =   5640
            TabIndex        =   27
            Top             =   3240
            Width           =   2145
         End
         Begin VB.CheckBox Check18 
            BackColor       =   &H0000C000&
            Caption         =   "Pedidos a Proveedores"
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
            Left            =   5640
            TabIndex        =   26
            Top             =   3480
            Width           =   2385
         End
         Begin VB.CheckBox Check20 
            BackColor       =   &H0000C000&
            Caption         =   "Cerrar Pedido Proveedo"
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
            Left            =   5640
            TabIndex        =   25
            Top             =   3960
            Width           =   2385
         End
         Begin VB.CheckBox Check19 
            BackColor       =   &H0000C000&
            Caption         =   "Cerrar Pedido Clientes"
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
            Left            =   5640
            TabIndex        =   24
            Top             =   3720
            Width           =   2265
         End
         Begin VB.CheckBox check7 
            BackColor       =   &H000080FF&
            Caption         =   "Captura De Paros"
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
            Left            =   360
            TabIndex        =   21
            Top             =   3240
            Width           =   2055
         End
         Begin VB.CheckBox check8 
            BackColor       =   &H000080FF&
            Caption         =   "Reportes "
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
            Left            =   360
            TabIndex        =   22
            Top             =   3480
            Width           =   2415
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H000000FF&
            Caption         =   "Configuracion"
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
            Left            =   360
            TabIndex        =   14
            Top             =   1320
            Width           =   2500
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H000000FF&
            Caption         =   "Menu De Produccion"
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
            Left            =   360
            TabIndex        =   15
            Top             =   1560
            Width           =   2500
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H000000FF&
            Caption         =   "Especificacione, Rutinas, Defectos, Atributos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   360
            TabIndex        =   16
            Top             =   1800
            Width           =   2505
         End
         Begin VB.CheckBox Check4 
            BackColor       =   &H000000FF&
            Caption         =   "Reportes"
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
            Left            =   360
            TabIndex        =   17
            Top             =   2280
            Width           =   2500
         End
         Begin VB.CheckBox Check5 
            BackColor       =   &H80000004&
            Caption         =   "Ajustes Inventario"
            Height          =   195
            Left            =   240
            TabIndex        =   18
            Top             =   5640
            Width           =   1665
         End
         Begin VB.Frame Frame3 
            Caption         =   "Opciones Avanzadas "
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
            Height          =   1335
            Left            =   120
            TabIndex        =   11
            Top             =   4680
            Width           =   2415
            Begin VB.CheckBox check13 
               Caption         =   "Borrar"
               Height          =   195
               Left            =   120
               TabIndex        =   5
               Top             =   720
               Width           =   975
            End
            Begin VB.CheckBox check12 
               Caption         =   "Editar"
               Height          =   195
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   855
            End
            Begin VB.CheckBox check11 
               Caption         =   "Usuarios"
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4106
                  SubFormatType   =   0
               EndProperty
               Height          =   195
               Left            =   120
               TabIndex        =   12
               Top             =   240
               Width           =   1935
            End
         End
         Begin VB.TextBox Txtusuario 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   960
            MaxLength       =   20
            TabIndex        =   2
            ToolTipText     =   "Limite 10 Caracteres"
            Top             =   240
            Width           =   1875
         End
         Begin VB.TextBox TxtPassword 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3000
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
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
            IMEMode         =   3  'DISABLE
            Left            =   4080
            MaxLength       =   5
            PasswordChar    =   "*"
            TabIndex        =   3
            ToolTipText     =   "Limite 5 Caracteres"
            Top             =   240
            Width           =   975
         End
         Begin VB.TextBox TxtNombres 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   960
            TabIndex        =   4
            Top             =   600
            Width           =   4095
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00C0E0FF&
            FillStyle       =   0  'Solid
            Height          =   1335
            Index           =   2
            Left            =   5520
            Top             =   5040
            Width           =   2535
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00C0C0C0&
            FillStyle       =   0  'Solid
            Height          =   1335
            Index           =   6
            Left            =   3000
            Top             =   5040
            Width           =   2415
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ordenes De Produccion"
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
            Left            =   3000
            TabIndex        =   58
            Top             =   4800
            Width           =   2040
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Menu De Empleados"
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
            Left            =   5640
            TabIndex        =   52
            Top             =   960
            Width           =   1755
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H0080C0FF&
            FillStyle       =   0  'Solid
            Height          =   1455
            Index           =   5
            Left            =   5520
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Cantidad De Accesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   7800
            TabIndex        =   48
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Acceso"
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
            Left            =   5160
            TabIndex        =   46
            Top             =   600
            Width           =   1230
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H0000C000&
            FillStyle       =   0  'Solid
            Height          =   1575
            Index           =   4
            Left            =   5520
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Alta"
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
            Left            =   5160
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Menu De Eficiencia"
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
            TabIndex        =   37
            Top             =   2640
            Width           =   1680
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Menu De Inventario"
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
            Left            =   3000
            TabIndex        =   36
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Menu De Calidad"
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
            Left            =   120
            TabIndex        =   35
            Top             =   960
            Width           =   1470
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H00FF8080&
            FillStyle       =   0  'Solid
            Height          =   3615
            Index           =   3
            Left            =   3000
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H000080FF&
            FillStyle       =   0  'Solid
            Height          =   1335
            Index           =   1
            Left            =   120
            Top             =   2880
            Width           =   2775
         End
         Begin VB.Shape Shape1 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   1335
            Index           =   0
            Left            =   120
            Top             =   1200
            Width           =   2775
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Label2"
            Height          =   15
            Left            =   840
            TabIndex        =   9
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "Password"
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
            Left            =   3000
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Nombres"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim buscar As String
Dim mensaje As String
Dim Bandera As Boolean
Dim VLargo As Integer
Dim VEncriptado As String
Dim Cont As Integer
Dim VNumeroAscii As String
Dim REquipos As New ADODB.Recordset
Dim RUsuarios As New ADODB.Recordset
Dim BEditar As Boolean
Dim VTexto As String

Private Sub CmdAgregar_Click()
 On Error Resume Next
    
   Bandera = True
   botones
   Limpia_Campos
   TxtTexto.Enabled = True
   Txtusuario.Enabled = True
   Txtusuario.SetFocus
   BEditar = False
   MskFecAlt.Text = Date
   
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RUsuarios.Delete
                        
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
                        RUsuarios.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RUsuarios.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If
   

End Sub

Private Sub CmdBotones2_Click(Index As Integer)
On Error Resume Next
MousePointer = 11
    If Index = 1 Then
        RUsuarios.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RUsuarios.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RUsuarios.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RUsuarios.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RUsuarios.BOF Then
        RUsuarios.MoveFirst
    ElseIf RUsuarios.EOF Then
        RUsuarios.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub


Private Sub CmdCancelar_Click()
    Bandera = False
    botones
    Llena_Campos
    Txtusuario.Enabled = True
End Sub

Private Sub CmdGrabar_Click()
    On Error Resume Next
    
        If Txtusuario.Text = "" Then
            MsgBox "Usuario No Puede Estar Vacio", vbOKOnly, "Error"
            Txtusuario.SetFocus
            Exit Sub
        End If
    
    
        If TxtTexto.Text = "" Then
            MsgBox "Password No Puede Estar Vacio", vbOKOnly, "Error"
            TxtPassword.SetFocus
            Exit Sub
        End If
        
        VTexto = ""
        
        'Set REquipos = Db.OpenRecordset("Select Equipo From Equipos Where Equipo = '" & TxtEquipo.Text & "'")
        'If REquipos.RecordCount > 0 Then
        'Else
        '        MsgBox "Equipo No Existe ", vbOKOnly
        '        TxtEquipo.SetFocus
        '        Exit Sub
        'End If
                
        If BEditar = False Then
                    '----------------------------------------------------------------------------------------------------
                    'PROCESO PARA ENCRIPTAR EL PASSWORD AGARRAMOS CADA LETRA DEL PASSWORD Y LE ASIGNAMOS EL CODIGO ASSCII
                    'Y TAMBIEN LE AGREGAMOS UN NUMERO CUALQUIERA (0110) PARA QUE SEA UN POCO MAS DIFICIL DE LEERLO
                    
                    Cont = 1
                    VLargo = Len(TxtTexto.Text)
                    VNumeroAscii = ""
                    VEncriptado = ""
                    
                    Do While Cont <= VLargo
                       VNumeroAscii = Asc(Mid(TxtTexto.Text, Cont, 1))
                       VEncriptado = VEncriptado & VNumeroAscii & "0110"
                       Cont = Cont + 1
                    Loop
                    
                    TxtPassword.Text = VEncriptado
                    
                            VTexto = "'" & Txtusuario.Text & "', '" ' USUARIO
                            VTexto = VTexto & TxtNombres.Text & "', '" 'NOMBRES
                            VTexto = VTexto & TxtPassword.Text & "', " 'CLAVE
                            
                            If Check1.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check2.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check3.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check4.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check5.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check6.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If check7.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If check8.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check9.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If check10.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If check11.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If check12.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If check13.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check14.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check15.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check16.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check17.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check18.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check19.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check20.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check21.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check22.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check23.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check24.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check25.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check26.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check27.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check28.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check29.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check30.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check31.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check32.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check33.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check34.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check35.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check36.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " '% conforme entradas inventario
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check37.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check38.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check39.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check40.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'MENU PRODUCCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check41.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'configuracion calidad
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            
                            
                            
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "#" & Format(MskFecAlt.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "To_Date('" & Format(MskFecAlt.Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "#" & Format(Date, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "To_Date('" & Format(Date, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            
                            VTexto = VTexto & TxtCon.Text & ", '" 'CONTADOR ACCESSOS
                            VTexto = VTexto & GUsuario & "', " 'USUARIO AGREGAR
                            
                            If Check42.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'reportes formatos
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check43.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'INSPECCION
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check44.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'CAPTURA DESPERDICIO
                            Else
                                VTexto = VTexto & "0" & ", "
                            End If
                            If Check45.Value = "1" Then 'RECLAMOS PROVEEODR
                                VTexto = VTexto & "-1"
                            Else
                                VTexto = VTexto & "0"
                            End If
                            
                            Conexion.Execute "Insert Into Usuarios Values(" & VTexto & ")"
        Else
                                    
                                    
                            If Check1.Value = "1" Then
                                VTexto = VTexto & "ConfiguracionCalidad = -1" & ", "
                            Else
                                VTexto = VTexto & "ConfiguracionCalidad = 0" & ", "
                            End If
                            If Check2.Value = "1" Then
                                VTexto = VTexto & "Produccion = -1" & ", "
                            Else
                                VTexto = VTexto & "Produccion = 0" & ", "
                            End If
                            If Check3.Value = "1" Then
                                VTexto = VTexto & "Especificaciones = -1" & ", "
                            Else
                                VTexto = VTexto & "Especificaciones = 0" & ", "
                            End If
                            If Check4.Value = "1" Then
                                VTexto = VTexto & "ReportesCalidad = -1" & ", "
                            Else
                                VTexto = VTexto & "ReportesCalidad = 0" & ", "
                            End If
                            If Check5.Value = "1" Then
                                VTexto = VTexto & "GraficasCalidad = -1" & ", "
                            Else
                                VTexto = VTexto & "GraficasCalidad = 0" & ", "
                            End If
                            If Check6.Value = "1" Then
                                VTexto = VTexto & "ConfiguracionEficiencia = -1" & ", "
                            Else
                                VTexto = VTexto & "ConfiguracionEficiencia = 0" & ", "
                            End If
                            If check7.Value = "1" Then
                                VTexto = VTexto & "CapturaParos = -1" & ", "
                            Else
                                VTexto = VTexto & "CapturaParos = 0" & ", "
                            End If
                            If check8.Value = "1" Then
                                VTexto = VTexto & "ReportesEficiencia = -1" & ", "
                            Else
                                VTexto = VTexto & "ReportesEficiencia = 0" & ", "
                            End If
                            If Check9.Value = "1" Then
                                VTexto = VTexto & "EditarEficiencia = -1" & ", "
                            Else
                                VTexto = VTexto & "EditarEficiencia = 0" & ", "
                            End If
                            If check10.Value = "1" Then
                                VTexto = VTexto & "BorrarEficiencia = -1" & ", "
                            Else
                                VTexto = VTexto & "BorrarEficiencia = 0" & ", "
                            End If
                            If check11.Value = "1" Then
                                VTexto = VTexto & "Usuarios = -1" & ", "
                            Else
                                VTexto = VTexto & "Usuarios = 0" & ", "
                            End If
                            If check12.Value = "1" Then
                                VTexto = VTexto & "Editar = -1" & ", "
                            Else
                                VTexto = VTexto & "Editar = 0" & ", "
                            End If
                            If check13.Value = "1" Then
                                VTexto = VTexto & "Borrar = -1" & ", "
                            Else
                                VTexto = VTexto & "Borrar = 0" & ", "
                            End If
                            If Check14.Value = "1" Then
                                VTexto = VTexto & "OrdenProduccion = -1" & ", "
                            Else
                                VTexto = VTexto & "OrdenProduccion = 0" & ", "
                            End If
                            If Check15.Value = "1" Then
                                VTexto = VTexto & "InvVenRepEje = -1" & ", "
                            Else
                                VTexto = VTexto & "InvVenRepEje = 0" & ", "
                            End If
                            If Check16.Value = "1" Then
                                VTexto = VTexto & "ReportesOrdenes = -1" & ", "
                            Else
                                VTexto = VTexto & "ReportesOrdenes = 0" & ", "
                            End If
                            If Check17.Value = "1" Then
                                VTexto = VTexto & "PedidosClientes = -1" & ", "
                            Else
                                VTexto = VTexto & "PedidosClientes = 0" & ", "
                            End If
                            If Check18.Value = "1" Then
                                VTexto = VTexto & "PedidosProveedores = -1" & ", "
                            Else
                                VTexto = VTexto & "PedidosProveedores = 0" & ", "
                            End If
                            If Check19.Value = "1" Then
                                VTexto = VTexto & "CerrarPedidosClientes = -1" & ", "
                            Else
                                VTexto = VTexto & "CerrarPedidosClientes = 0" & ", "
                            End If
                            If Check20.Value = "1" Then
                                VTexto = VTexto & "CerrarPedidosProveedores = -1" & ", "
                            Else
                                VTexto = VTexto & "CerrarpedidosProveedores = 0" & ", "
                            End If
                            If Check21.Value = "1" Then
                                VTexto = VTexto & "EditarPedidos = -1" & ", "
                            Else
                                VTexto = VTexto & "EditarPedidos = 0" & ", "
                            End If
                            If Check22.Value = "1" Then
                                VTexto = VTexto & "BorrarPedidos = -1" & ", "
                            Else
                                VTexto = VTexto & "BorrarPedidos = 0" & ", "
                            End If
                            If Check23.Value = "1" Then
                                VTexto = VTexto & "ConfiguracionInventario = -1" & ", "
                            Else
                                VTexto = VTexto & "ConfiguracionInventario = 0" & ", "
                            End If
                            If Check24.Value = "1" Then
                                VTexto = VTexto & "Entradas = -1" & ", "
                            Else
                                VTexto = VTexto & "Entradas = 0" & ", "
                            End If
                            If Check25.Value = "1" Then
                                VTexto = VTexto & "Traslados = -1" & ", "
                            Else
                                VTexto = VTexto & "Traslados = 0" & ", "
                            End If
                            If Check26.Value = "1" Then
                                VTexto = VTexto & "Salidas = -1" & ", "
                            Else
                                VTexto = VTexto & "Salidas = 0" & ", "
                            End If
                            If Check27.Value = "1" Then
                                VTexto = VTexto & "CambiosUbicacion = -1" & ", "
                            Else
                                VTexto = VTexto & "CambiosUbicacion = 0" & ", "
                            End If
                            If Check28.Value = "1" Then
                                VTexto = VTexto & "CierreBulto = -1" & ", "
                            Else
                                VTexto = VTexto & "CierreBulto = 0" & ", "
                            End If
                            If Check29.Value = "1" Then
                                VTexto = VTexto & "LiberacionEntradas = -1" & ", "
                            Else
                                VTexto = VTexto & "LiberacionEntradas = 0" & ", "
                            End If
                            If Check30.Value = "1" Then
                                VTexto = VTexto & "LiberacionTraslados = -1" & ", "
                            Else
                                VTexto = VTexto & "LiberacionTraslados = 0" & ", "
                            End If
                            If Check31.Value = "1" Then
                                VTexto = VTexto & "LiberacionSalidas = -1" & ", "
                            Else
                                VTexto = VTexto & "LiberacionSalidas = 0" & ", "
                            End If
                            If Check32.Value = "1" Then
                                VTexto = VTexto & "GraficasInventario = -1" & ", "
                            Else
                                VTexto = VTexto & "GraficasInventario = 0" & ", "
                            End If
                            If Check33.Value = "1" Then
                                VTexto = VTexto & "ReportesInventario = -1" & ", "
                            Else
                                VTexto = VTexto & "ReportesInventario = 0" & ", "
                            End If
                            If Check34.Value = "1" Then
                                VTexto = VTexto & "CapturaTransito = -1" & ", "
                            Else
                                VTexto = VTexto & "CapturaTransito = 0" & ", "
                            End If
                            If Check35.Value = "1" Then
                                VTexto = VTexto & "ConsultaTransito = -1" & ", "
                            Else
                                VTexto = VTexto & "ConsultaTransito = 0" & ", "
                            End If
                            If Check36.Value = "1" Then
                                VTexto = VTexto & "PorConEntInv = -1" & ", "
                            Else
                                VTexto = VTexto & "PorConEntInv = 0" & ", "
                            End If
                            If Check37.Value = "1" Then
                                VTexto = VTexto & "ConfiguracionEmpleados = -1" & ", "
                            Else
                                VTexto = VTexto & "ConfiguracionEmpleados = 0" & ", "
                            End If
                            If Check38.Value = "1" Then
                                VTexto = VTexto & "CapturaFaltas = -1" & ", "
                            Else
                                VTexto = VTexto & "CapturaFaltas = 0" & ", "
                            End If
                            If Check39.Value = "1" Then
                                VTexto = VTexto & "CapturaCursos = -1" & ", "
                            Else
                                VTexto = VTexto & "CapturaCursos = 0" & ", "
                            End If
                            If Check40.Value = "1" Then
                                VTexto = VTexto & "CapturaAumentos = -1" & ", "
                            Else
                                VTexto = VTexto & "CapturaAumentos = 0" & ", "
                            End If
                            If Check41.Value = "1" Then
                                VTexto = VTexto & "ReportesEmpleados = -1" & ", "
                            Else
                                VTexto = VTexto & "ReportesEmpleados = 0" & ", "
                            End If
                            If Check42.Value = "1" Then
                                VTexto = VTexto & "ReportesFormatos = -1" & ", "
                            Else
                                VTexto = VTexto & "ReportesFormatos = 0" & ", "
                            End If
                            If Check43.Value = "1" Then
                                VTexto = VTexto & "Inspeccion = -1" & ", "
                            Else
                                VTexto = VTexto & "Inspeccion = 0" & ", "
                            End If
                            If Check44.Value = "1" Then
                                VTexto = VTexto & "CapturaDesperdicio = -1" & ", "
                            Else
                                VTexto = VTexto & "CapturaDesperdicio = 0" & ", "
                            End If
                            If Check45.Value = "1" Then
                                VTexto = VTexto & "ReclamosProveedores = -1" & ", "
                            Else
                                VTexto = VTexto & "ReclamosProveedores = 0" & ", "
                            End If
                            
                            
                            
                            VTexto = VTexto & "usuarioagregar = '" & GUsuario & "' " ' USUARIO
                            VTexto = VTexto & "Where Usuario = '" & Txtusuario.Text & "'"
                     
                            Conexion.Execute "UPDATE Usuarios SET " & VTexto
        End If
     
                        'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Codigo De Usuario Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo De Usuario Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            
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
                        Txtusuario.Enabled = True
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RUsuarios.Requery
                        RUsuarios.MoveLast
                        Llena_Campos
   
       
    
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Sub botones()
    If Bandera = True Then
      FrameAgencias.Enabled = True
      CmdAgregar.Enabled = False
      CmdGrabar.Enabled = True
      CmdEditar.Enabled = False
      
      CmdBorrar.Enabled = False
      CmdCancelar.Enabled = True
      CmdSalida.Enabled = False
    Else
      FrameAgencias.Enabled = False
      CmdAgregar.Enabled = True
      CmdGrabar.Enabled = False
      CmdEditar.Enabled = True
      
      CmdBorrar.Enabled = True
      CmdCancelar.Enabled = False
      CmdSalida.Enabled = True
    End If
    
End Sub


Private Sub CmdEditar_Click()
 On Error Resume Next
   Bandera = True
   botones
   Txtusuario.Enabled = False
   TxtTexto.Enabled = False
   TxtNombres.SetFocus
   BEditar = True
End Sub


Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
        RUsuarios.Sort = RUsuarios.Fields(ColIndex).Name
End Sub

Private Sub Form_Load()
On Error Resume Next
        Set RUsuarios = New ADODB.Recordset
        
        Call Abrir_Recordset(RUsuarios, "Select * from Usuarios")
            
        Set DbGrid1.DataSource = RUsuarios
    
                If Err <> 0 Then
                        MsgBox Err.Number & Err.Description
                End If
                
                Llena_Campos
        
        BEditar = False
        If GEditar = True Then
                DbGrid1.AllowUpdate = True
        Else
                DbGrid1.AllowUpdate = False
        End If
        
        
        
End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        CmdBorrar.Enabled = True
        If CmdGrabar.Enabled = False Then
            Llena_Campos
        End If
    Else
        CmdBorrar.Enabled = False
    End If
End Sub

Private Sub TxtPassword_Change()
    TxtTexto.Text = TxtPassword.Text
End Sub





Public Sub Llena_Campos()
On Error Resume Next
        
    If RUsuarios.RecordCount > 0 Then
        Txtusuario.Text = RUsuarios!Usuario
            If IsNull(RUsuarios!Nombres) Then
                TxtNombres.Text = ""
            Else
                TxtNombres.Text = RUsuarios!Nombres
            End If
        
        TxtPassword.Text = RUsuarios!Clave
            If IsNull(RUsuarios!Clave) Then
                TxtPassword.Text = ""
            Else
                TxtPassword.Text = RUsuarios!Clave
            End If
            
        
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ConfiguracionCalidad = "Verdadero" Then
                    Check1.Value = "1"
                Else
                    Check1.Value = "0"
                End If
        Else
                If RUsuarios!ConfiguracionCalidad = "-1" Then
                    Check1.Value = "1"
                Else
                    Check1.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Produccion = "Verdadero" Then
                    Check2.Value = "1"
                Else
                    Check2.Value = "0"
                End If
        Else
                If RUsuarios!Produccion = "-1" Then
                    Check2.Value = "1"
                Else
                    Check2.Value = "0"
                End If
        End If
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Especificaciones = "Verdadero" Then
                    Check3.Value = "1"
                Else
                    Check3.Value = "0"
                End If
        Else
                If RUsuarios!Especificaciones = "-1" Then
                    Check3.Value = "1"
                Else
                    Check3.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ReportesCalidad = "Verdadero" Then
                    Check4.Value = "1"
                Else
                    Check4.Value = "0"
                End If
        Else
                If RUsuarios!ReportesCalidad = "-1" Then
                    Check4.Value = "1"
                Else
                    Check4.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!GraficasCalidad = "Verdadero" Then
                    Check5.Value = "1"
                Else
                    Check5.Value = "0"
                End If
        Else
                If RUsuarios!GraficasCalidad = "-1" Then
                    Check5.Value = "1"
                Else
                    Check5.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ConfiguracionEficiencia = "Verdadero" Then
                    Check6.Value = "1"
                Else
                    Check6.Value = "0"
                End If
        Else
                If RUsuarios!ConfiguracionEficiencia = "-1" Then
                    Check6.Value = "1"
                Else
                    Check6.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CapturaParos = "Verdadero" Then
                    check7.Value = "1"
                Else
                    check7.Value = "0"
                End If
        Else
                If RUsuarios!CapturaParos = "-1" Then
                    check7.Value = "1"
                Else
                    check7.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ReportesEficiencia = "Verdadero" Then
                    check8.Value = "1"
                Else
                    check8.Value = "0"
                End If
        Else
                If RUsuarios!ReportesEficiencia = "-1" Then
                    check8.Value = "1"
                Else
                    check8.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!EditarEficiencia = "Verdadero" Then
                    Check9.Value = "1"
                Else
                    Check9.Value = "0"
                End If
        Else
                If RUsuarios!EditarEficiencia = "-1" Then
                    Check9.Value = "1"
                Else
                    Check9.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!BorrarEficiencia = "Verdadero" Then
                    check10.Value = "1"
                Else
                    check10.Value = "0"
                End If
        Else
                If RUsuarios!BorrarEficiencia = "-1" Then
                    check10.Value = "1"
                Else
                    check10.Value = "0"
                End If
        End If
        
        
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Usuarios = "Verdadero" Then
                    check11.Value = "1"
                Else
                    check11.Value = "0"
                End If
        Else
                If RUsuarios!Usuarios = "-1" Then
                    check11.Value = "1"
                Else
                    check11.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Editar = "Verdadero" Then
                    check12.Value = "1"
                Else
                    check12.Value = "0"
                End If
        Else
                If RUsuarios!Editar = "-1" Then
                    check12.Value = "1"
                Else
                    check12.Value = "0"
                End If
        End If
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Borrar = "Verdadero" Then
                    check13.Value = "1"
                Else
                    check13.Value = "0"
                End If
        Else
                If RUsuarios!Borrar = "-1" Then
                    check13.Value = "1"
                Else
                    check13.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!OrdenProduccion = "Verdadero" Then
                    Check14.Value = "1"
                Else
                    Check14.Value = "0"
                End If
        Else
                If RUsuarios!OrdenProduccion = "-1" Then
                    Check14.Value = "1"
                Else
                    Check14.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!InvVenRepEje = "Verdadero" Then
                    Check15.Value = "1"
                Else
                    Check15.Value = "0"
                End If
        Else
                If RUsuarios!InvVenRepEje = "-1" Then
                    Check15.Value = "1"
                Else
                    Check15.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ReportesOrdenes = "Verdadero" Then
                    Check16.Value = "1"
                Else
                    Check16.Value = "0"
                End If
        Else
                If RUsuarios!ReportesOrdenes = "-1" Then
                    Check16.Value = "1"
                Else
                    Check16.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!PedidosClientes = "Verdadero" Then
                    Check17.Value = "1"
                Else
                    Check17.Value = "0"
                End If
        Else
                If RUsuarios!PedidosClientes = "-1" Then
                    Check17.Value = "1"
                Else
                    Check17.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!PedidosProveedores = "Verdadero" Then
                    Check18.Value = "1"
                Else
                    Check18.Value = "0"
                End If
        Else
                If RUsuarios!PedidosProveedores = "-1" Then
                    Check18.Value = "1"
                Else
                    Check18.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CerrarPedidosClientes = "Verdadero" Then
                    Check19.Value = "1"
                Else
                    Check19.Value = "0"
                End If
        Else
                If RUsuarios!CerrarPedidosClientes = "-1" Then
                    Check19.Value = "1"
                Else
                    Check19.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CerrarpedidosProveedores = "Verdadero" Then
                    Check20.Value = "1"
                Else
                    Check20.Value = "0"
                End If
        Else
                If RUsuarios!CerrarpedidosProveedores = "-1" Then
                    Check20.Value = "1"
                Else
                    Check20.Value = "0"
                End If
        End If
        
        
        
        
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!EditarPedidos = "Verdadero" Then
                    Check21.Value = "1"
                Else
                    Check21.Value = "0"
                End If
        Else
                If RUsuarios!EditarPedidos = "-1" Then
                    Check21.Value = "1"
                Else
                    Check21.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!BorrarPedidos = "Verdadero" Then
                    Check22.Value = "1"
                Else
                    Check22.Value = "0"
                End If
        Else
                If RUsuarios!BorrarPedidos = "-1" Then
                    Check22.Value = "1"
                Else
                    Check22.Value = "0"
                End If
        End If
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ConfiguracionInventario = "Verdadero" Then
                    Check23.Value = "1"
                Else
                    Check23.Value = "0"
                End If
        Else
                If RUsuarios!ConfiguracionInventario = "-1" Then
                    Check23.Value = "1"
                Else
                    Check23.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Entradas = "Verdadero" Then
                    Check24.Value = "1"
                Else
                    Check24.Value = "0"
                End If
        Else
                If RUsuarios!Entradas = "-1" Then
                    Check24.Value = "1"
                Else
                    Check24.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Traslados = "Verdadero" Then
                    Check25.Value = "1"
                Else
                    Check25.Value = "0"
                End If
        Else
                If RUsuarios!Traslados = "-1" Then
                    Check25.Value = "1"
                Else
                    Check25.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Salidas = "Verdadero" Then
                    Check26.Value = "1"
                Else
                    Check26.Value = "0"
                End If
        Else
                If RUsuarios!Salidas = "-1" Then
                    Check26.Value = "1"
                Else
                    Check26.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CambiosUbicacion = "Verdadero" Then
                    Check27.Value = "1"
                Else
                    Check27.Value = "0"
                End If
        Else
                If RUsuarios!CambiosUbicacion = "-1" Then
                    Check27.Value = "1"
                Else
                    Check27.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CierreBulto = "Verdadero" Then
                    Check28.Value = "1"
                Else
                    Check28.Value = "0"
                End If
        Else
                If RUsuarios!CierreBulto = "-1" Then
                    Check28.Value = "1"
                Else
                    Check28.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!LiberacionEntradas = "Verdadero" Then
                    Check29.Value = "1"
                Else
                    Check29.Value = "0"
                End If
        Else
                If RUsuarios!LiberacionEntradas = "-1" Then
                    Check29.Value = "1"
                Else
                    Check29.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!LiberacionTraslados = "Verdadero" Then
                    Check30.Value = "1"
                Else
                    Check30.Value = "0"
                End If
        Else
                If RUsuarios!LiberacionTraslados = "-1" Then
                    Check30.Value = "1"
                Else
                    Check30.Value = "0"
                End If
        End If
        
        
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!LiberacionSalidas = "Verdadero" Then
                    Check31.Value = "1"
                Else
                    Check31.Value = "0"
                End If
        Else
                If RUsuarios!LiberacionSalidas = "-1" Then
                    Check31.Value = "1"
                Else
                    Check31.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!GraficasInventario = "Verdadero" Then
                    Check32.Value = "1"
                Else
                    Check32.Value = "0"
                End If
        Else
                If RUsuarios!GraficasInventario = "-1" Then
                    Check32.Value = "1"
                Else
                    Check32.Value = "0"
                End If
        End If
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ReportesInventario = "Verdadero" Then
                    Check33.Value = "1"
                Else
                    Check33.Value = "0"
                End If
        Else
                If RUsuarios!ReportesInventario = "-1" Then
                    Check33.Value = "1"
                Else
                    Check33.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CapturaTransito = "Verdadero" Then
                    Check34.Value = "1"
                Else
                    Check34.Value = "0"
                End If
        Else
                If RUsuarios!CapturaTransito = "-1" Then
                    Check34.Value = "1"
                Else
                    Check34.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ConsultaTransito = "Verdadero" Then
                    Check35.Value = "1"
                Else
                    Check35.Value = "0"
                End If
        Else
                If RUsuarios!ConsultaTransito = "-1" Then
                    Check35.Value = "1"
                Else
                    Check35.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!PorConEntInv = "Verdadero" Then
                    Check36.Value = "1"
                Else
                    Check36.Value = "0"
                End If
        Else
                If RUsuarios!PorConEntInv = "-1" Then
                    Check36.Value = "1"
                Else
                    Check36.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ConfiguracionEmpleados = "Verdadero" Then
                    Check37.Value = "1"
                Else
                    Check37.Value = "0"
                End If
        Else
                If RUsuarios!ConfiguracionEmpleados = "-1" Then
                    Check37.Value = "1"
                Else
                    Check37.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CapturaFaltas = "Verdadero" Then
                    Check38.Value = "1"
                Else
                    Check38.Value = "0"
                End If
        Else
                If RUsuarios!CapturaFaltas = "-1" Then
                    Check38.Value = "1"
                Else
                    Check38.Value = "0"
                End If
        End If
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CapturaCursos = "Verdadero" Then
                    Check39.Value = "1"
                Else
                    Check39.Value = "0"
                End If
        Else
                If RUsuarios!CapturaCursos = "-1" Then
                    Check39.Value = "1"
                Else
                    Check39.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CapturaAumentos = "Verdadero" Then
                    Check40.Value = "1"
                Else
                    Check40.Value = "0"
                End If
        Else
                If RUsuarios!CapturaAumentos = "-1" Then
                    Check40.Value = "1"
                Else
                    Check40.Value = "0"
                End If
        End If
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ReportesEmpleados = "Verdadero" Then
                    Check41.Value = "1"
                Else
                    Check41.Value = "0"
                End If
        Else
                If RUsuarios!ReportesEmpleados = "-1" Then
                    Check41.Value = "1"
                Else
                    Check41.Value = "0"
                End If
        End If
        
        
        
            If IsNull(RUsuarios!FechaAlta) Then
                MskFecAlt.Text = ""
            Else
                MskFecAlt.Text = RUsuarios!FechaAlta
            End If
            
            If IsNull(RUsuarios!FechaUltimoAcceso) Then
                MskFecUltAcc.Text = ""
            Else
                MskFecUltAcc.Text = RUsuarios!FechaUltimoAcceso
            End If
            
            If IsNull(RUsuarios!ContadorAccesos) Then
                TxtCon.Text = "0"
            Else
                TxtCon.Text = RUsuarios!ContadorAccesos
            End If
            
            
            
            If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!ReportesFormatos = "Verdadero" Then
                    Check42.Value = "1"
                Else
                    Check42.Value = "0"
                End If
            Else
                    If RUsuarios!ReportesFormatos = "-1" Then
                        Check42.Value = "1"
                    Else
                        Check42.Value = "0"
                    End If
            End If
            If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!Inspeccion = "Verdadero" Then
                    Check43.Value = "1"
                Else
                    Check43.Value = "0"
                End If
            Else
                    If RUsuarios!Inspeccion = "-1" Then
                        Check43.Value = "1"
                    Else
                        Check43.Value = "0"
                    End If
            End If
        
        
            If GOrigenDeDatos = "AmaproAccess" Then
                If RUsuarios!CapturaDesperdicio = "Verdadero" Then
                    Check44.Value = "1"
                Else
                    Check44.Value = "0"
                End If
            Else
                    If RUsuarios!CapturaDesperdicio = "-1" Then
                        Check44.Value = "1"
                    Else
                        Check44.Value = "0"
                    End If
            End If
        
            If GOrigenDeDatos = "AmaproAccess" Then
                    If RUsuarios!ReclamosProveedores = "Verdadero" Then
                        Check45.Value = "1"
                    Else
                        Check45.Value = "0"
                    End If
            Else
                    If RUsuarios!ReclamosProveedores = "-1" Then
                        Check45.Value = "1"
                    Else
                        Check45.Value = "0"
                    End If
            End If
        
            
            
            If Err <> 0 Then
                MsgBox Err.Number & Err.Description
            End If
            
            
    Else
                Txtusuario.Text = ""
                TxtPassword.Text = ""
                TxtTexto.Text = ""
                MskFecAlt.Text = ""
                TxtNombres.Text = ""
                MskFecUltAcc.Text = ""
                TxtCon.Text = 0
                Check1.Value = 0
                Check2.Value = 0
                Check3.Value = 0
                Check4.Value = 0
                Check5.Value = 0
                Check6.Value = 0
                check7.Value = 0
                check8.Value = 0
                Check9.Value = 0
                check10.Value = 0
                check11.Value = 0
                check12.Value = 0
                check13.Value = 0
                Check14.Value = 0
                Check15.Value = 0
                Check16.Value = 0
                Check17.Value = 0
                Check18.Value = 0
                Check19.Value = 0
                Check20.Value = 0
                Check21.Value = 0
                Check22.Value = 0
                Check23.Value = 0
                Check24.Value = 0
                Check25.Value = 0
                Check26.Value = 0
                Check27.Value = 0
                Check28.Value = 0
                Check29.Value = 0
                Check30.Value = 0
                Check31.Value = 0
                Check32.Value = 0
                Check33.Value = 0
                Check34.Value = 0
                Check35.Value = 0
                Check36.Value = 0
                Check37.Value = 0
                Check38.Value = 0
                Check39.Value = 0
                Check40.Value = 0
                Check41.Value = 0
        
    End If
    
        
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        Txtusuario.Text = ""
        TxtPassword.Text = ""
        TxtTexto.Text = ""
        MskFecAlt.Text = ""
        TxtNombres.Text = ""
        MskFecUltAcc.Text = ""
        TxtCon.Text = 0
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Check4.Value = 0
        Check5.Value = 0
        Check6.Value = 0
        check7.Value = 0
        check8.Value = 0
        Check9.Value = 0
        check10.Value = 0
        check11.Value = 0
        check12.Value = 0
        check13.Value = 0
        Check14.Value = 0
        Check15.Value = 0
        Check16.Value = 0
        Check17.Value = 0
        Check18.Value = 0
        Check19.Value = 0
        Check20.Value = 0
        Check21.Value = 0
        Check22.Value = 0
        Check23.Value = 0
        Check24.Value = 0
        Check25.Value = 0
        Check26.Value = 0
        Check27.Value = 0
        Check28.Value = 0
        Check29.Value = 0
        Check30.Value = 0
        Check31.Value = 0
        Check32.Value = 0
        Check33.Value = 0
        Check34.Value = 0
        Check35.Value = 0
        Check36.Value = 0
        Check37.Value = 0
        Check38.Value = 0
        Check39.Value = 0
        Check40.Value = 0
        Check41.Value = 0
        Check42.Value = 0
        Check43.Value = 0
        Check44.Value = 0
        Check45.Value = 0
End Sub

