VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Empleados 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha Tecnica De Empleados"
   ClientHeight    =   8115
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "Empleados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Framebuscar 
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
      Height          =   8055
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6855
         Left            =   120
         TabIndex        =   64
         Top             =   1080
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   12091
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
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   66
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   65
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox Txtbuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   63
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   7335
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   8760
         Picture         =   "Empleados.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
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
         TabIndex        =   67
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   480
      MouseIcon       =   "Empleados.frx":237C
      Picture         =   "Empleados.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   94
      ToolTipText     =   "Primer Registro"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   840
      MouseIcon       =   "Empleados.frx":2CF0
      Picture         =   "Empleados.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   93
      ToolTipText     =   "Registro Anterior"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   10680
      MouseIcon       =   "Empleados.frx":3664
      Picture         =   "Empleados.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   92
      ToolTipText     =   "Siguiente Registro"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   11040
      MouseIcon       =   "Empleados.frx":3FD8
      Picture         =   "Empleados.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   91
      ToolTipText     =   "Ultimo Registro"
      Top             =   7440
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   720
      Index           =   5
      Left            =   9120
      MouseIcon       =   "Empleados.frx":494C
      Picture         =   "Empleados.frx":4D8E
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   7320
      Width           =   1500
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   720
      Index           =   4
      Left            =   7560
      MouseIcon       =   "Empleados.frx":6E00
      Picture         =   "Empleados.frx":7242
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   7320
      Width           =   1500
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   720
      Index           =   2
      Left            =   4440
      MouseIcon       =   "Empleados.frx":7774
      Picture         =   "Empleados.frx":7BB6
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   7320
      Width           =   1500
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   720
      Index           =   1
      Left            =   2880
      MouseIcon       =   "Empleados.frx":80E8
      Picture         =   "Empleados.frx":852A
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   7320
      Width           =   1500
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   720
      Index           =   0
      Left            =   1320
      MouseIcon       =   "Empleados.frx":8A5C
      Picture         =   "Empleados.frx":8E9E
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   7320
      Width           =   1500
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   720
      Index           =   3
      Left            =   6000
      MouseIcon       =   "Empleados.frx":93D0
      Picture         =   "Empleados.frx":9812
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   7320
      Width           =   1500
   End
   Begin TabDlg.SSTab TabParos 
      Height          =   7215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12726
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "Empleados.frx":9D44
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameParos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Empleados.frx":A05E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridEmpleados"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda O Seleccion De Datos"
      TabPicture(2)   =   "Empleados.frx":A378
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusqueda"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DbGridEmpleados 
         Height          =   6375
         Left            =   -74880
         TabIndex        =   90
         Top             =   720
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   11245
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
               LCID            =   2058
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
               LCID            =   2058
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
      Begin VB.Frame FrameBusqueda 
         Height          =   4335
         Left            =   -73080
         TabIndex        =   46
         Top             =   1200
         Width           =   7935
         Begin VB.OptionButton OptBuscar 
            Caption         =   "Puesto"
            Height          =   255
            Index           =   4
            Left            =   360
            Picture         =   "Empleados.frx":A692
            TabIndex        =   51
            Top             =   2160
            Width           =   855
         End
         Begin VB.OptionButton OptBuscar 
            Caption         =   "Departamento"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   50
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton OptBuscar 
            Caption         =   "Equipo"
            Height          =   195
            Index           =   2
            Left            =   360
            TabIndex        =   49
            Top             =   1440
            Width           =   855
         End
         Begin VB.OptionButton OptBuscar 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   360
            Picture         =   "Empleados.frx":D50C
            TabIndex        =   48
            Top             =   1080
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   4920
            TabIndex        =   52
            Top             =   1920
            Width           =   2655
         End
         Begin VB.CommandButton CmdBusqueda 
            Caption         =   "Seleccionar Datos"
            Height          =   735
            Index           =   0
            Left            =   4920
            Picture         =   "Empleados.frx":10386
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   2400
            Width           =   2655
         End
         Begin VB.CommandButton CmdBusqueda 
            Caption         =   "Seleccionar Todos Los Datos"
            Height          =   735
            Index           =   1
            Left            =   4920
            Picture         =   "Empleados.frx":10910
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   3360
            Width           =   2655
         End
         Begin VB.OptionButton OptBuscar 
            Caption         =   "Codigo"
            Height          =   255
            Index           =   0
            Left            =   360
            Picture         =   "Empleados.frx":10C1A
            TabIndex        =   47
            Top             =   720
            Width           =   855
         End
         Begin VB.Label LblDesPar 
            Alignment       =   1  'Right Justify
            Caption         =   "Nombre"
            Height          =   255
            Left            =   3360
            TabIndex        =   57
            Top             =   1920
            Width           =   1455
         End
      End
      Begin VB.Frame FrameParos 
         Caption         =   "Datos del Empleado"
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
         Height          =   6375
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   11655
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   28
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   23
            Top             =   6000
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   27
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   38
            Top             =   5640
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   26
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   37
            Top             =   5280
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   25
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   36
            Top             =   4920
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   24
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   35
            Top             =   4440
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   23
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   34
            Top             =   4080
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   22
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   33
            Top             =   3720
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   21
            Left            =   4920
            MaxLength       =   20
            TabIndex        =   17
            Top             =   4200
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   20
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   32
            Top             =   3360
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   19
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   31
            Top             =   3000
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   18
            Left            =   9840
            MaxLength       =   20
            TabIndex        =   30
            Top             =   2640
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   17
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   13
            Top             =   3480
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   16
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   12
            Top             =   3120
            Width           =   6255
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Esta Afecto Al  Isr ?"
            Height          =   195
            Left            =   3480
            TabIndex        =   20
            Top             =   4920
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   15
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   10
            Top             =   2400
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   14
            Left            =   9840
            MaxLength       =   30
            TabIndex        =   24
            ToolTipText     =   "maximo 15 caracteres"
            Top             =   240
            Width           =   1695
         End
         Begin VB.ComboBox CboEstCiv 
            Height          =   315
            ItemData        =   "Empleados.frx":13A94
            Left            =   1680
            List            =   "Empleados.frx":13AA4
            TabIndex        =   16
            Text            =   "SOLTERO"
            Top             =   4200
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   13
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   11
            Top             =   2760
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   12
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   22
            Top             =   5640
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   11
            Left            =   9840
            MaxLength       =   30
            TabIndex        =   26
            ToolTipText     =   "maximo 15 caracteres"
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   10
            Left            =   9840
            MaxLength       =   30
            TabIndex        =   25
            ToolTipText     =   "maximo 15 caracteres"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   9
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   21
            ToolTipText     =   "maximo 30 caracteres"
            Top             =   5280
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   1680
            TabIndex        =   19
            Top             =   4920
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   7
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   18
            Top             =   4560
            Width           =   6255
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   6
            Left            =   1680
            MaxLength       =   30
            TabIndex        =   14
            ToolTipText     =   "maximo 30 caracteres"
            Top             =   3840
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   5
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   4
            Top             =   960
            Width           =   1695
         End
         Begin MSMask.MaskEdBox MskSue 
            Height          =   285
            Left            =   9840
            TabIndex        =   27
            Top             =   1440
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.ComboBox CboEstado 
            Height          =   315
            ItemData        =   "Empleados.frx":13ACC
            Left            =   1680
            List            =   "Empleados.frx":13AD6
            TabIndex        =   7
            Text            =   "ALTA"
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   6
            Top             =   1680
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
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
            Index           =   3
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   5
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   0
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   2
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   1
            Left            =   1680
            MaxLength       =   50
            TabIndex        =   3
            Top             =   600
            Width           =   6255
         End
         Begin MSMask.MaskEdBox MskFecNac 
            Height          =   285
            Left            =   6240
            TabIndex        =   15
            Top             =   3840
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskBon 
            Height          =   285
            Left            =   9840
            TabIndex        =   28
            Top             =   1800
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecBaj 
            Height          =   285
            Left            =   6240
            TabIndex        =   9
            Top             =   2040
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFecAlt 
            Height          =   285
            Left            =   3600
            TabIndex        =   8
            Top             =   2040
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskBonInc 
            Height          =   285
            Left            =   9840
            TabIndex        =   29
            Top             =   2160
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            BackColor       =   16777215
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   36
            Left            =   240
            TabIndex        =   108
            Top             =   6000
            Width           =   1065
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Clinica"
            Height          =   195
            Index           =   35
            Left            =   8760
            TabIndex        =   107
            Top             =   3720
            Width           =   465
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Consulta"
            Height          =   195
            Index           =   34
            Left            =   8760
            TabIndex        =   106
            Top             =   4080
            Width           =   615
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
            Height          =   195
            Index           =   33
            Left            =   8760
            TabIndex        =   105
            Top             =   4440
            Width           =   420
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Uniforme"
            Height          =   195
            Index           =   32
            Left            =   8760
            TabIndex        =   104
            Top             =   4920
            Width           =   630
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Camisa"
            Height          =   195
            Index           =   31
            Left            =   8760
            TabIndex        =   103
            Top             =   5280
            Width           =   510
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Pantalon"
            Height          =   195
            Index           =   30
            Left            =   8760
            TabIndex        =   102
            Top             =   5640
            Width           =   630
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Jefes"
            Height          =   195
            Index           =   29
            Left            =   4440
            TabIndex        =   101
            Top             =   4200
            Width           =   375
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Sangre"
            Height          =   195
            Index           =   28
            Left            =   8760
            TabIndex        =   100
            Top             =   3360
            Width           =   870
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Temperamento"
            Height          =   195
            Index           =   27
            Left            =   8760
            TabIndex        =   99
            Top             =   3000
            Width           =   1065
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Zodiaco"
            Height          =   195
            Index           =   26
            Left            =   8760
            TabIndex        =   98
            Top             =   2640
            Width           =   585
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Colonia"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   97
            Top             =   3120
            Width           =   525
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Localidad"
            Height          =   195
            Index           =   24
            Left            =   240
            TabIndex        =   96
            Top             =   3480
            Width           =   690
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Bono Incentivo"
            Height          =   195
            Index           =   23
            Left            =   8760
            TabIndex        =   95
            Top             =   2160
            Width           =   1080
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Motivo Baja"
            Height          =   195
            Index           =   22
            Left            =   240
            TabIndex        =   89
            Top             =   2400
            Width           =   840
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Baja"
            Height          =   195
            Index           =   21
            Left            =   5400
            TabIndex        =   88
            Top             =   2040
            Width           =   810
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Alta"
            Height          =   195
            Index           =   20
            Left            =   2760
            TabIndex        =   87
            Top             =   2040
            Width           =   765
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "RFC"
            Height          =   195
            Index           =   17
            Left            =   8760
            TabIndex        =   86
            Top             =   240
            Width           =   315
         End
         Begin VB.Label LblEsc 
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
            Left            =   3480
            TabIndex        =   85
            Top             =   5280
            Width           =   4455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Bonificacion"
            Height          =   195
            Index           =   16
            Left            =   8760
            TabIndex        =   84
            Top             =   1800
            Width           =   870
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "CURP"
            Height          =   195
            Index           =   15
            Left            =   8760
            TabIndex        =   83
            Top             =   600
            Width           =   450
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "IMSS"
            Height          =   195
            Index           =   14
            Left            =   8760
            TabIndex        =   82
            Top             =   960
            Width           =   390
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Correo Electronico"
            Height          =   195
            Index           =   13
            Left            =   240
            TabIndex        =   81
            Top             =   5640
            Width           =   1305
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Telefono"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   80
            Top             =   3840
            Width           =   630
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Nacimiento"
            Height          =   195
            Index           =   11
            Left            =   4920
            TabIndex        =   79
            Top             =   3840
            Width           =   1290
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Estado Civil"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   78
            Top             =   4200
            Width           =   825
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Esposa"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   77
            Top             =   4560
            Width           =   1125
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad De Hijos"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   76
            Top             =   4920
            Width           =   1275
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Escolaridad"
            Height          =   195
            Index           =   7
            Left            =   240
            TabIndex        =   75
            Top             =   5280
            Width           =   825
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Puesto"
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   74
            Top             =   960
            Width           =   495
         End
         Begin VB.Label LblPuesto 
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
            Left            =   3480
            TabIndex        =   73
            Top             =   960
            Width           =   4455
         End
         Begin VB.Label LblDep 
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
            Left            =   3480
            TabIndex        =   72
            Top             =   1680
            Width           =   4455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Direccion"
            Height          =   195
            Index           =   5
            Left            =   240
            TabIndex        =   71
            Top             =   2760
            Width           =   675
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Sueldo Base"
            Height          =   195
            Index           =   4
            Left            =   8760
            TabIndex        =   70
            Top             =   1440
            Width           =   900
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Estado"
            Height          =   195
            Index           =   3
            Left            =   240
            TabIndex        =   69
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Departamento"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   68
            Top             =   1680
            Width           =   1005
         End
         Begin VB.Label LblGrupo 
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
            Left            =   3480
            TabIndex        =   60
            Top             =   1320
            Width           =   4455
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Equipo"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   59
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   0
            Left            =   5640
            TabIndex        =   58
            Top             =   240
            Width           =   540
         End
         Begin VB.Label lblLabels 
            Caption         =   "Codigo"
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   56
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   55
            Top             =   600
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "Empleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim RBuscaGrupo As New ADODB.Recordset
Dim RBuscamaximo As New ADODB.Recordset
Dim RBuscaDepartamento As New ADODB.Recordset
Dim RBuscaPuesto As New ADODB.Recordset
Dim RBuscaEscolaridad As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim REmpleados As New ADODB.Recordset

Dim BEquipo As Boolean
Dim BDepartamento As Boolean
Dim BPuesto As Boolean
Dim BEscolaridad As Boolean

Dim BEditar As Boolean
Dim VTexto As String
Dim VEmpleado As String




Sub botones()
    If Bandera = True Then
         FrameParos.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
                 'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         DbGridEmpleados.Visible = False
         FrameBusqueda.Visible = False
    Else
         FrameParos.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         
         DbGridEmpleados.Visible = True
         FrameBusqueda.Visible = True
    End If
End Sub

Private Sub CboEstado_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
End Sub


Private Sub CboEstCiv_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
End Sub




Private Sub Chk_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
End Sub

Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next

                'AGREGAR
                If Index = 0 Then
                                            Bandera = True
                                            botones
                                            Limpia_Campos
                                            TxtTexto.Item(0).Enabled = True
                                            TxtTexto.Item(3).Text = GUsuario
                                            TxtTexto.Item(0).SetFocus
                                            CboEstado.Text = "ALTA"
                                            CboEstCiv.Text = "SOLTERO"
                                            BEditar = False
                                    
                'EDITAR
                ElseIf Index = 1 Then
                                            Bandera = True
                                            botones
                                            VEmpleado = TxtTexto.Item(0).Text
                                            TxtTexto.Item(1).SetFocus
                                            TxtTexto.Item(3).Text = GUsuario
                                            BEditar = True
                                    
                'GRABAR
                ElseIf Index = 2 Then
                                    
                                    
                                    MskFecAlt.Text = Format(MskFecAlt.Text, "dd/mm/yyyy")
                                    MskFecNac.Text = Format(MskFecNac.Text, "dd/mm/yyyy")
                                    
                                    'REVISA EL CODIGO DEL EQUIPO
                                    If TxtTexto.Item(2).Text = "" Then
                                        MsgBox "Equipo No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                                        TxtTexto.Item(2).SetFocus
                                        Exit Sub
                                    End If
                                    
                                    'REVISA EL CODIGO DEL DEPARTAMENTO
                                    If TxtTexto.Item(4).Text = "" Then
                                        MsgBox "Departamento No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                                        TxtTexto.Item(4).SetFocus
                                        Exit Sub
                                    End If
                                    
                                    'REVISA EL CODIGO DEL PUESTO
                                    If TxtTexto.Item(5).Text = "" Then
                                        MsgBox "Puesto No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                                        TxtTexto.Item(5).SetFocus
                                        Exit Sub
                                    End If
                                    'REVISA EL ESTADO DEL EMPLEADO
                                    If CboEstado.Text <> "ALTA" And CboEstado.Text <> "BAJA" Then
                                        MsgBox "Estado Del Empleado Incorrecto", vbOKOnly + vbInformation, "Informacion"
                                        CboEstado.SetFocus
                                        Exit Sub
                                    End If
                                    'REVISA EL ESTADO DEL EMPLEADO
                                    If CboEstCiv.Text <> "SOLTERO" And CboEstCiv.Text <> "CASADO" And CboEstCiv.Text <> "DIVORCIADO" And CboEstCiv.Text <> "VIUDO" Then
                                        MsgBox "Estado Civil Incorrecto", vbOKOnly + vbInformation, "Informacion"
                                        CboEstCiv.SetFocus
                                        Exit Sub
                                    End If
                                    'REVISA LA FECHA DE NACIMIENTO
                                    If Not IsDate(MskFecNac.Text) Then
                                        MsgBox "Fecha De Nacimiento Incorrecta", vbOKOnly + vbInformation, "Informacion"
                                        MskFecNac.SetFocus
                                        Exit Sub
                                    End If
                                    'REVISA SUELDO
                                    If Not IsNumeric(MskSue.Text) Then
                                        MsgBox "Sueldo Incorrecto", vbOKOnly + vbInformation, "Informacion"
                                        MskSue.SetFocus
                                        Exit Sub
                                    End If
                                    'REVISA BONIFICACION
                                    If Not IsNumeric(MskBon.Text) Then
                                        MsgBox "Bonificacion Incorrecta", vbOKOnly + vbInformation, "Informacion"
                                        MskBon.SetFocus
                                        Exit Sub
                                    End If
                                    'REVISA BONO INCENTIVO
                                    If Not IsNumeric(MskBonInc.Text) Then
                                        MsgBox "Bono Incentivo Incorrecto", vbOKOnly + vbInformation, "Informacion"
                                        MskBonInc.SetFocus
                                        Exit Sub
                                    End If
                                    'REVISA HIJOS
                                    If Not IsNumeric(TxtTexto.Item(8).Text) Then
                                        MsgBox "Cantidad De Hijos Incorrecta", vbOKOnly + vbInformation, "Informacion"
                                        TxtTexto.Item(8).SetFocus
                                        Exit Sub
                                    End If
                                    
                                    Set RBuscaGrupo = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtTexto.Item(2).Text & "'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From EmpleadosGrupos Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(2).Text) & "'")
                                        End If
                                        If RBuscaGrupo.RecordCount > 0 Then
                                        Else
                                            MsgBox "Equipo De Empleados No Existe", vbOKOnly + vbInformation, "Informacion"
                                            TxtTexto.Item(2).SetFocus
                                            Exit Sub
                                        End If
                                'BUSCA DEPARTAMENTO
                                    Set RBuscaDepartamento = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where Codigo = '" & TxtTexto.Item(4).Text & "'")
                                        Else 'ORACLE
                                                Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(4).Text) & "'")
                                        End If
                                        If RBuscaDepartamento.RecordCount > 0 Then
                                        Else
                                            MsgBox "Departamento No Existe", vbOKOnly + vbInformation, "Informacion"
                                            TxtTexto.Item(4).SetFocus
                                            Exit Sub
                                        End If
                                'BUSCA PUESTO
                                    Set RBuscaPuesto = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBuscaPuesto, "Select Descripcion From EmpleadosPuestos Where CodigoPuesto = '" & TxtTexto.Item(5).Text & "'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBuscaPuesto, "Select Descripcion From EmpleadosPuestos Where UPPER(CodigoPuesto) = '" & UCase(TxtTexto.Item(5).Text) & "'")
                                        End If
                                
                                        If RBuscaPuesto.RecordCount > 0 Then
                                        Else
                                            MsgBox "Puesto No Existe", vbOKOnly + vbInformation, "Informacion"
                                            TxtTexto.Item(5).SetFocus
                                            Exit Sub
                                        End If
                                'ESCOLARIDAD
                                    Set RBuscaEscolaridad = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBuscaEscolaridad, "Select Descripcion From EmpleadosEscolaridad Where Codigo = '" & TxtTexto.Item(9).Text & "'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBuscaEscolaridad, "Select Descripcion From EmpleadosEscolaridad Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(9).Text) & "'")
                                        End If
                                        If RBuscaEscolaridad.RecordCount > 0 Then
                                            
                                        Else
                                            MsgBox "Escolaridad No Existe", vbOKOnly + vbInformation, "Informacion"
                                            TxtTexto.Item(9).SetFocus
                                            Exit Sub
                                        End If
                            
                                    
                                    'AGREGAR
                                        If BEditar = False Then
                                                VTexto = "'" & TxtTexto.Item(0).Text & "', '" 'CODIGO
                                                VTexto = VTexto & TxtTexto.Item(1).Text & "', '" 'DESCRIPCION
                                                VTexto = VTexto & TxtTexto.Item(5).Text & "', '" 'PUESTO
                                                VTexto = VTexto & TxtTexto.Item(2).Text & "', '" 'EQUIPO
                                                VTexto = VTexto & TxtTexto.Item(4).Text & "', '" 'DEPARTAMENTO
                                                VTexto = VTexto & CboEstado.Text & "', " 'ESTADO
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "#" & Format(MskFecAlt.Text, "mm/dd/yyyy") & "#, '" 'FECHA ALTA
                                                Else 'ORACLE
                                                     VTexto = VTexto & "To_Date('" & MskFecAlt.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA ALTA
                                                End If
                                                VTexto = VTexto & MskFecBaj.Text & "', '" 'FECHA BAJA
                                                VTexto = VTexto & TxtTexto.Item(15).Text & "', " 'MOTIVO BAJA
                                                VTexto = VTexto & MskSue.Text & ", " 'SUELDO
                                                VTexto = VTexto & MskBon.Text & ", '" 'BONIFICACION
                                                VTexto = VTexto & TxtTexto.Item(13).Text & "', '" 'DIRECCION
                                                VTexto = VTexto & TxtTexto.Item(6).Text & "', " 'TELEFONO
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "#" & Format(MskFecNac.Text, "mm/dd/yyyy") & "#, '" 'FECHA NACIMIENTO
                                                Else 'ORACLE
                                                     VTexto = VTexto & "To_Date('" & MskFecNac.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA NACIMIENTO
                                                End If
                                                VTexto = VTexto & CboEstCiv.Text & "', '" 'ESTADO CIVIL
                                                VTexto = VTexto & TxtTexto.Item(7).Text & "', " 'ESPOSA
                                                VTexto = VTexto & TxtTexto.Item(8).Text & ", '" 'HIJOS
                                                VTexto = VTexto & TxtTexto.Item(9).Text & "', '" 'ESCOLARIDAD
                                                VTexto = VTexto & TxtTexto.Item(10).Text & "', '" 'CEDULA
                                                VTexto = VTexto & TxtTexto.Item(11).Text & "', '" 'IGSS
                                                VTexto = VTexto & TxtTexto.Item(14).Text & "', " 'NIT
                                                VTexto = VTexto & Chk.Value & ", '" 'ISR
                                                VTexto = VTexto & TxtTexto.Item(12).Text & "', '" 'CORREO ELECTRONICO
                                                VTexto = VTexto & TxtTexto.Item(3).Text & "', " 'USUARIO
                                                VTexto = VTexto & MskBonInc.Text & " , '" 'BONO INCENTIVO
                                                VTexto = VTexto & TxtTexto.Item(16).Text & "', '" 'COLONIA
                                                VTexto = VTexto & TxtTexto.Item(17).Text & "', '" 'LOCALIDAD
                                                VTexto = VTexto & TxtTexto.Item(18).Text & "', '" 'ZODIACO
                                                VTexto = VTexto & TxtTexto.Item(19).Text & "', '" 'TEMPERAMENTO
                                                VTexto = VTexto & TxtTexto.Item(21).Text & "', '" 'JEFES
                                                VTexto = VTexto & TxtTexto.Item(20).Text & "', '" 'TIPOSANGRE
                                                VTexto = VTexto & TxtTexto.Item(22).Text & "', '" 'CLINICA
                                                VTexto = VTexto & TxtTexto.Item(23).Text & "', '" 'CONSULTA
                                                VTexto = VTexto & TxtTexto.Item(24).Text & "', '" 'TURNO
                                                VTexto = VTexto & TxtTexto.Item(25).Text & "', '" 'UNIFORME
                                                VTexto = VTexto & TxtTexto.Item(26).Text & "', '" 'CAMISA
                                                VTexto = VTexto & TxtTexto.Item(27).Text & "', '" 'PANTALON
                                                VTexto = VTexto & TxtTexto.Item(28).Text & "'" 'OBSERVACIONES
                                                
                                                
                                                'REALIZA EL INSERT
                                                Conexion.Execute "Insert Into Empleados Values(" & VTexto & ")"
                                        'EDITAR
                                        Else
                                                VTexto = "Codigo = '" & TxtTexto.Item(0).Text & "', " 'CODIGO
                                                VTexto = VTexto & "Descripcion = '" & TxtTexto.Item(1).Text & "', " 'DESCRIPCION
                                                VTexto = VTexto & "Puesto = '" & TxtTexto.Item(5).Text & "', " 'PUESTO
                                                VTexto = VTexto & "Grupo = '" & TxtTexto.Item(2).Text & "', " 'EQUIPO
                                                VTexto = VTexto & "Departamento = '" & TxtTexto.Item(4).Text & "', " 'DEPARTAMENTO
                                                VTexto = VTexto & "Estado = '" & CboEstado.Text & "', " 'ESTADO
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "FechaAlta = #" & Format(MskFecAlt.Text, "mm/dd/yyyy") & "#, " 'FECHA ALTA
                                                Else 'ORACLE
                                                     VTexto = VTexto & "FechaAlta = To_Date('" & MskFecAlt.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA ALTA
                                                End If
                                                VTexto = VTexto & "FechaBaja = '" & MskFecBaj.Text & "', " 'FECHA BAJA
                                                VTexto = VTexto & "MotivoBaja = '" & TxtTexto.Item(15).Text & "', " 'MOTIVO BAJA
                                                VTexto = VTexto & "SueldoBase = " & MskSue.Text & ", " 'SUELDO
                                                VTexto = VTexto & "Bonificacion = " & MskBon.Text & ", " 'BONIFICACION
                                                VTexto = VTexto & "Direccion = '" & TxtTexto.Item(13).Text & "', " 'DIRECCION
                                                VTexto = VTexto & "Telefono = '" & TxtTexto.Item(6).Text & "', " 'TELEFONO
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "FechaNacimiento = #" & Format(MskFecNac.Text, "mm/dd/yyyy") & "#, " 'FECHA NACIMIENTO
                                                Else 'ORACLE
                                                     VTexto = VTexto & "FechaNacimiento = To_Date('" & MskFecNac.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA NACIMIENTO
                                                End If
                                                VTexto = VTexto & "EstadoCivil = '" & CboEstCiv.Text & "', " 'ESTADO CIVIL
                                                VTexto = VTexto & "NombreEsposa = '" & TxtTexto.Item(7).Text & "', " 'ESPOSA
                                                VTexto = VTexto & "Hijos = " & TxtTexto.Item(8).Text & ", " 'HIJOS
                                                VTexto = VTexto & "Escolaridad = '" & TxtTexto.Item(9).Text & "', " 'ESCOLARIDAD
                                                VTexto = VTexto & "Cedula = '" & TxtTexto.Item(10).Text & "', " 'CEDULA
                                                VTexto = VTexto & "Igss = '" & TxtTexto.Item(11).Text & "', " 'IGSS
                                                VTexto = VTexto & "Nit = '" & TxtTexto.Item(14).Text & "', " 'NIT
                                                VTexto = VTexto & "AfectoIsr = " & Chk.Value & ", " 'ISR
                                                VTexto = VTexto & "CorreoElectronico = '" & TxtTexto.Item(12).Text & "', " 'CORREO ELECTRONICO
                                                VTexto = VTexto & "Usuario = '" & TxtTexto.Item(3).Text & "', " 'USUARIO
                                                VTexto = VTexto & "BonoIncentivo = " & MskBonInc.Text & ", " 'BONO INCENTIVO
                                                VTexto = VTexto & "Colonia = '" & TxtTexto.Item(16).Text & "', " 'COLONIA
                                                VTexto = VTexto & "Localidad = '" & TxtTexto.Item(17).Text & "', " 'LOCALIDAD
                                                VTexto = VTexto & "Zodiaco = '" & TxtTexto.Item(18).Text & "', " 'ZODIACO
                                                VTexto = VTexto & "Temperamento = '" & TxtTexto.Item(19).Text & "', " 'TEMPERAMENTO
                                                VTexto = VTexto & "Jefes = '" & TxtTexto.Item(21).Text & "', " 'JEFES
                                                VTexto = VTexto & "TipoSanguineo = '" & TxtTexto.Item(20).Text & "', " 'TIPO SANGUINES
                                                VTexto = VTexto & "Clinica = '" & TxtTexto.Item(22).Text & "', " 'CLINICA
                                                VTexto = VTexto & "Consulta = '" & TxtTexto.Item(23).Text & "', " 'CONSULTA
                                                VTexto = VTexto & "Turno = '" & TxtTexto.Item(24).Text & "', " 'TURNO
                                                VTexto = VTexto & "Uniforme = '" & TxtTexto.Item(25).Text & "', " 'UNIFORME
                                                VTexto = VTexto & "Camisa = '" & TxtTexto.Item(26).Text & "', " 'CAMISA
                                                VTexto = VTexto & "Pantalon = '" & TxtTexto.Item(27).Text & "', " 'PANTALON
                                                VTexto = VTexto & "Observaciones = '" & TxtTexto.Item(28).Text & "'" 'OBSERVACIONES
                                                
                                                VTexto = VTexto & " Where Codigo = '" & VEmpleado & "'" 'EMPLEADO
                                                
                                                Conexion.Execute "UPDATE Empleados SET " & VTexto
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
                                                MsgBox "Codigo Empleado Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                                TxtTexto.Item(0).SetFocus
                                                Exit Sub
                                          'SI ES CUALQUIER OTRO ERROR
                                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                Exit Sub
                                            End If
                                        End If
                                    
                                    
                                       Bandera = False
                                       botones
                                       CmdBotones.Item(0).SetFocus
                                       
                                        'HABILITA LA LLAVE
                                        TxtTexto.Item(0).Enabled = True
                                                                
                                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                                        REmpleados.Requery
                                        REmpleados.MoveLast
                                        Llena_Campos
                                   
                'CANCELAR
                ElseIf Index = 3 Then
                                    
                                        Bandera = False
                                        botones
                                        Llena_Campos
                                        'HABILITA LA LLAVE
                                        TxtTexto.Item(0).Enabled = True
                                    
                'BORRAR
                ElseIf Index = 4 Then
                                       mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                                    If mensaje = vbOK Then
                                        'BORRA EL REGISTRO
                                        REmpleados.Delete
                                        
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
                                        REmpleados.Requery
                                        'MUEVE AL SIGUIENTE REGISTRO
                                        REmpleados.MoveLast
                                        'SI HAY ERRORES
                                        If Err <> 0 Then
                                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Err.Clear
                                        End If
                                        
                                        Llena_Campos
                                    End If
            
                'SALIDA
                Else
                                        Unload Me
                End If
    
    
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        REmpleados.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        REmpleados.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        REmpleados.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        REmpleados.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If REmpleados.BOF Then
        REmpleados.MoveFirst
    ElseIf REmpleados.EOF Then
        REmpleados.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBusqueda_Click(Index As Integer)
MousePointer = 11
        Set REmpleados = New ADODB.Recordset
        
        If Index = 0 Then
            If OptBuscar.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where Codigo like '%" & TxtBusqueda.Item(0).Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where UPPER(Codigo) like '%" & UCase(TxtBusqueda.Item(0).Text) & "%'")
                End If
            ElseIf OptBuscar.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where Descripcion like '%" & TxtBusqueda.Item(0).Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where UPPER(Descripcion) like '%" & UCase(TxtBusqueda.Item(0).Text) & "%'")
                End If
            ElseIf OptBuscar.Item(2).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where Grupo = '" & TxtBusqueda.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where UPPER(Grupo) = '" & UCase(TxtBusqueda.Item(0).Text) & "'")
                End If
            ElseIf OptBuscar.Item(3).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where Departamento = '" & TxtBusqueda.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where UPPER(Departamento) = '" & UCase(TxtBusqueda.Item(0).Text) & "'")
                End If
            ElseIf OptBuscar.Item(4).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where Puesto = '" & TxtBusqueda.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(REmpleados, "Select * from Empleados Where UPPER(Puesto) = '" & UCase(TxtBusqueda.Item(0).Text) & "'")
                End If
            End If
            
            Set DbGridEmpleados.DataSource = REmpleados
        End If
        
        If Index = 1 Then
            Call Abrir_Recordset(REmpleados, "Select * from Empleados")
            Set DbGridEmpleados.DataSource = REmpleados
        End If
            TabParos.Tab = 1

MousePointer = 0
End Sub


Private Sub CmdSale_Click()
    FrameBuscar.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
    If BEquipo = True Then
        TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(2).SetFocus
    ElseIf BDepartamento = True Then
        TxtTexto.Item(4).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(4).SetFocus
    ElseIf BPuesto = True Then
        TxtTexto.Item(5).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(5).SetFocus
    ElseIf BEscolaridad = True Then
        TxtTexto.Item(9).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(9).SetFocus
    End If
        FrameBuscar.Visible = False

End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                If BEquipo = True Then
                    TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0)
                    TxtTexto.Item(2).SetFocus
                ElseIf BDepartamento = True Then
                    TxtTexto.Item(4).Text = DBGridBusqueda.Columns(0)
                    TxtTexto.Item(4).SetFocus
                ElseIf BPuesto = True Then
                    TxtTexto.Item(5).Text = DBGridBusqueda.Columns(0)
                    TxtTexto.Item(5).SetFocus
                ElseIf BEscolaridad = True Then
                    TxtTexto.Item(9).Text = DBGridBusqueda.Columns(0)
                    TxtTexto.Item(9).SetFocus
                End If
                    FrameBuscar.Visible = False
        End If
End Sub


Private Sub DbGridEmpleados_HeadClick(ByVal ColIndex As Integer)
                REmpleados.Sort = REmpleados.Fields(ColIndex).Name
End Sub


Private Sub Form_Load()
                    Set REmpleados = New ADODB.Recordset
                        Call Abrir_Recordset(REmpleados, "Select * From Empleados")
                        Set DbGridEmpleados.DataSource = REmpleados
                        Llena_Campos
                    
        
    'VALIDA SI EL USUARIO PUEDE EDITAR
    If GEditar = True Then
        DbGridEmpleados.AllowUpdate = True
    Else
        DbGridEmpleados.AllowUpdate = False
    End If
    
    
End Sub


Private Sub MskBon_GotFocus()
        MskBon.SelStart = 0
        MskBon.SelLength = Len(MskBon.Text)
End Sub

Private Sub MskBon_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub



Private Sub MskBonInc_GotFocus()
        MskBonInc.SelStart = 0
        MskBonInc.SelLength = Len(MskBonInc.Text)
End Sub

Private Sub MskBonInc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
End Sub

Private Sub MskFecAlt_GotFocus()
        MskFecAlt.SelStart = 0
        MskFecAlt.SelLength = Len(MskFecAlt.Text)
End Sub

Private Sub MskFecAlt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
End Sub


Private Sub MskFecBaj_GotFocus()
        MskFecBaj.SelStart = 0
        MskFecBaj.SelLength = Len(MskFecBaj.Text)
End Sub

Private Sub MskFecBaj_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
End Sub

Private Sub MskFecNac_GotFocus()
        MskFecNac.SelStart = 0
        MskFecNac.SelLength = Len(MskFecNac.Text)
End Sub

Private Sub MskFecNac_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{Tab}"
        End If
End Sub

Private Sub MskSue_GotFocus()
        MskSue.SelStart = 0
        MskSue.SelLength = Len(MskSue.Text)
End Sub

Private Sub MskSue_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub


Private Sub OptBuscar_Click(Index As Integer)
    If Index = 0 Then
        LblDesPar.Caption = "Codigo"
    ElseIf Index = 1 Then
        LblDesPar.Caption = "Nombre"
    ElseIf Index = 2 Then
        LblDesPar.Caption = "Equipo"
    ElseIf Index = 3 Then
        LblDesPar.Caption = "Departamento"
    ElseIf Index = 4 Then
        LblDesPar.Caption = "Puesto"
    End If
        TxtBusqueda.Item(0).SetFocus
End Sub

Private Sub TabParos_Click(PreviousTab As Integer)
    If TabParos.Tab = 0 Then
        CmdBotones.Item(4).Enabled = True
        If CmdBotones.Item(2).Enabled = False Then
            Llena_Campos
        End If
    ElseIf TabParos.Tab = 1 Or TabParos.Tab = 2 Then
        CmdBotones.Item(4).Enabled = False
    ElseIf TabParos.Tab = 2 Then
        TxtBusqueda.Item(0).SetFocus
    End If
End Sub

Private Sub Txtbuscar_Change()
            
                    Set RBusqueda = New ADODB.Recordset
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                                    If BEquipo = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                                        End If
                                    ElseIf BDepartamento = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                                        End If
                                    ElseIf BPuesto = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                                        End If
                                    ElseIf BEscolaridad = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad Where Descripcion Like '%" & TxtBuscar.Text & "%'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%'")
                                        End If
                                    End If
                    'OPCION DE CODIGO
                    Else
                                If BEquipo = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos Where Codigo Like '%" & TxtBuscar.Text & "%'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                                        End If
                                    ElseIf BDepartamento = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos Where Codigo Like '%" & TxtBuscar.Text & "%'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
                                        End If
                                    ElseIf BPuesto = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos Where CodigoPuesto Like '%" & TxtBuscar.Text & "%'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos Where UPPER(CodigoPuesto) Like '%" & UCase(TxtBuscar.Text) & "%'")
                                        End If
                                    ElseIf BEscolaridad = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad Where Codigo Like '%" & TxtBuscar.Text & "%'")
                                        Else 'ORACLE
                                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%'")
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

Private Sub TxtBusqueda_GotFocus(Index As Integer)
        TxtBusqueda.Item(0).SelStart = 0
        TxtBusqueda.Item(0).SelLength = Len(TxtBusqueda.Item(0).Text)
End Sub

Private Sub TxtBusqueda_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub


Private Sub TxtTexto_Change(Index As Integer)
    'BUSCA GRUPO
    If Index = 2 Then
        Set RBuscaGrupo = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtTexto.Item(2).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From EmpleadosGrupos Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(2).Text) & "'")
            End If
            If RBuscaGrupo.RecordCount > 0 Then
                LblGrupo.Caption = RBuscaGrupo!Descripcion
            Else
                LblGrupo.Caption = ""
            End If
    'BUSCA DEPARTAMENTO
    ElseIf Index = 4 Then
        Set RBuscaDepartamento = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where Codigo = '" & TxtTexto.Item(4).Text & "'")
            Else 'ORACLE
                    Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(4).Text) & "'")
            End If
            If RBuscaDepartamento.RecordCount > 0 Then
                LblDep.Caption = RBuscaDepartamento!Descripcion
            Else
                LblDep.Caption = ""
            End If
    'BUSCA PUESTO
    ElseIf Index = 5 Then
        Set RBuscaPuesto = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaPuesto, "Select Descripcion From EmpleadosPuestos Where CodigoPuesto = '" & TxtTexto.Item(5).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaPuesto, "Select Descripcion From EmpleadosPuestos Where UPPER(CodigoPuesto) = '" & UCase(TxtTexto.Item(5).Text) & "'")
            End If
    
            If RBuscaPuesto.RecordCount > 0 Then
                LblPuesto.Caption = RBuscaPuesto!Descripcion
            Else
                LblPuesto.Caption = ""
            End If
    'ESCOLARIDAD
    ElseIf Index = 9 Then
        Set RBuscaEscolaridad = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaEscolaridad, "Select Descripcion From EmpleadosEscolaridad Where Codigo = '" & TxtTexto.Item(9).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaEscolaridad, "Select Descripcion From EmpleadosEscolaridad Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(9).Text) & "'")
            End If
            If RBuscaEscolaridad.RecordCount > 0 Then
                LblEsc.Caption = RBuscaEscolaridad!Descripcion
            Else
                LblEsc.Caption = ""
            End If
    
    End If

End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        
        Set RBusqueda = New ADODB.Recordset
        'EQUIPOS
        If Index = 2 Then
                BEquipo = True
                BDepartamento = False
                BPuesto = False
                BEscolaridad = False
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos")
        'DEPARTAMENTOS
        ElseIf Index = 4 Then
                BEquipo = False
                BDepartamento = True
                BPuesto = False
                BEscolaridad = False
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos")
        'PUESTOS
        ElseIf Index = 5 Then
                BEquipo = False
                BDepartamento = False
                BPuesto = True
                BEscolaridad = False
                Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos")
        'ESCOLARIDAD
        ElseIf Index = 9 Then
                BEquipo = False
                BDepartamento = False
                BPuesto = False
                BEscolaridad = True
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad")
        End If
                
                
        If Index = 2 Or Index = 4 Or Index = 5 Or Index = 9 Then
                Set DBGridBusqueda.DataSource = RBusqueda
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                DBGridBusqueda.Columns(1).Width = "4000"
        End If

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        TxtTexto.Item(Index).SelStart = 0
        TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index))
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                'EQUIPOS
                If Index = 2 Then
                        BEquipo = True
                        BDepartamento = False
                        BPuesto = False
                        BEscolaridad = False
                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos")
                'DEPARTAMENTOS
                ElseIf Index = 4 Then
                        BEquipo = False
                        BDepartamento = True
                        BPuesto = False
                        BEscolaridad = False
                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos")
                'PUESTOS
                ElseIf Index = 5 Then
                        BEquipo = False
                        BDepartamento = False
                        BPuesto = True
                        BEscolaridad = False
                        Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos")
                'ESCOLARIDAD
                ElseIf Index = 9 Then
                        BEquipo = False
                        BDepartamento = False
                        BPuesto = False
                        BEscolaridad = True
                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad")
                End If
                        
                        
                If Index = 2 Or Index = 4 Or Index = 5 Or Index = 9 Then
                        Set DBGridBusqueda.DataSource = RBusqueda
                        FrameBuscar.Visible = True
                        TxtBuscar.SetFocus
                        DBGridBusqueda.Columns(1).Width = "4000"
                End If
    End If
    
End Sub

Public Sub Llena_Campos()
On Error Resume Next
        'CODIGO
            If IsNull(REmpleados!Codigo) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = REmpleados!Codigo
            End If
        'Descripcion
            If IsNull(REmpleados!Descripcion) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = REmpleados!Descripcion
            End If
        'PUESTO
            If IsNull(REmpleados!Puesto) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = REmpleados!Puesto
            End If
        'GRUPO
            If IsNull(REmpleados!Grupo) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = REmpleados!Grupo
            End If
        'DEPARTAMENTO
            If IsNull(REmpleados!Departamento) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = REmpleados!Departamento
            End If
        'ESTADO
            If IsNull(REmpleados!Estado) Then
                CboEstado.Text = ""
            Else
                CboEstado.Text = REmpleados!Estado
            End If
        'FECHA ALTA
            If IsNull(REmpleados!FechaAlta) Then
                MskFecAlt.Text = ""
            Else
                MskFecAlt.Text = REmpleados!FechaAlta
            End If
        'FECHA BAJA
            If IsNull(REmpleados!FechaBaja) Then
                MskFecBaj.Text = ""
            Else
                MskFecBaj.Text = REmpleados!FechaBaja
            End If
        'MOTIVO BAJA
            If IsNull(REmpleados!MotivoBaja) Then
                TxtTexto.Item(15).Text = ""
            Else
                TxtTexto.Item(15).Text = REmpleados!MotivoBaja
            End If
        'SUELDO BASE
            If IsNull(REmpleados!SueldoBase) Then
                MskSue.Text = "0"
            Else
                MskSue.Text = REmpleados!SueldoBase
            End If
        'BONIFICACION
            If IsNull(REmpleados!Bonificacion) Then
                MskBon.Text = "0"
            Else
                MskBon.Text = REmpleados!Bonificacion
            End If
        'DIRECCION
            If IsNull(REmpleados!Direccion) Then
                TxtTexto.Item(13).Text = ""
            Else
                TxtTexto.Item(13).Text = REmpleados!Direccion
            End If
        'TELEFONO
            If IsNull(REmpleados!Telefono) Then
                TxtTexto.Item(6).Text = ""
            Else
                TxtTexto.Item(6).Text = REmpleados!Telefono
            End If
        'FECHA NACIMIENTO
            If IsNull(REmpleados!FechaNacimiento) Then
                MskFecNac.Text = ""
            Else
                MskFecNac.Text = REmpleados!FechaNacimiento
            End If
        'ESTA CIVIL
            If IsNull(REmpleados!EstadoCivil) Then
                CboEstCiv.Text = ""
            Else
                CboEstCiv.Text = REmpleados!EstadoCivil
            End If
        'NOMBRE ESPOSA
            If IsNull(REmpleados!NombreEsposa) Then
                TxtTexto.Item(7).Text = ""
            Else
                TxtTexto.Item(7).Text = REmpleados!NombreEsposa
            End If
        'HIJOS
            If IsNull(REmpleados!Hijos) Then
                TxtTexto.Item(8).Text = ""
            Else
                TxtTexto.Item(8).Text = REmpleados!Hijos
            End If
        'ESCOLARIDAD
            If IsNull(REmpleados!Escolaridad) Then
                TxtTexto.Item(9).Text = ""
            Else
                TxtTexto.Item(9).Text = REmpleados!Escolaridad
            End If
        'CEDULA
            If IsNull(REmpleados!Cedula) Then
                TxtTexto.Item(10).Text = ""
            Else
                TxtTexto.Item(10).Text = REmpleados!Cedula
            End If
        'IGSS
            If IsNull(REmpleados!Igss) Then
                TxtTexto.Item(11).Text = ""
            Else
                TxtTexto.Item(11).Text = REmpleados!Igss
            End If
        'NIT
            If IsNull(REmpleados!Nit) Then
                TxtTexto.Item(14).Text = ""
            Else
                TxtTexto.Item(14).Text = REmpleados!Nit
            End If
        'AFECTOISR
            If IsNull(REmpleados!AfectoIsr) Then
                Chk.Value = ""
            Else
                Chk.Value = REmpleados!AfectoIsr
            End If
        'CORREO ELECTRONICO
            If IsNull(REmpleados!CorreoElectronico) Then
                TxtTexto.Item(12).Text = ""
            Else
                TxtTexto.Item(12).Text = REmpleados!CorreoElectronico
            End If
        'USUARIO
            If IsNull(REmpleados!Usuario) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = REmpleados!Usuario
            End If
        'BONo INCENTIVO
            If IsNull(REmpleados!BonoIncentivo) Then
                MskBonInc.Text = "0"
            Else
                MskBonInc.Text = REmpleados!BonoIncentivo
            End If
            
            
            
            
            'COLONIA
            If IsNull(REmpleados!Colonia) Then
                TxtTexto.Item(16).Text = ""
            Else
                TxtTexto.Item(16).Text = REmpleados!Colonia
            End If
        'LOCALIDAD
            If IsNull(REmpleados!Localidad) Then
                TxtTexto.Item(17).Text = ""
            Else
                TxtTexto.Item(17).Text = REmpleados!Localidad
            End If
        'ZODIACO
            If IsNull(REmpleados!Zodiaco) Then
                TxtTexto.Item(18).Text = ""
            Else
                TxtTexto.Item(18).Text = REmpleados!Zodiaco
            End If
        'TEMPERAMENTO
            If IsNull(REmpleados!temperamento) Then
                TxtTexto.Item(19).Text = ""
            Else
                TxtTexto.Item(19).Text = REmpleados!temperamento
            End If
        'Jefes
            If IsNull(REmpleados!Jefes) Then
                TxtTexto.Item(21).Text = ""
            Else
                TxtTexto.Item(21).Text = REmpleados!Jefes
            End If
        'TIPO SANGUINEO
            If IsNull(REmpleados!TipoSanguineo) Then
                TxtTexto.Item(20).Text = ""
            Else
                TxtTexto.Item(20).Text = REmpleados!TipoSanguineo
            End If
        'CLINICA
            If IsNull(REmpleados!Clinica) Then
                TxtTexto.Item(22).Text = ""
            Else
                TxtTexto.Item(22).Text = REmpleados!Clinica
            End If
        'consulta
            If IsNull(REmpleados!Consulta) Then
                 TxtTexto.Item(23).Text = ""
            Else
                TxtTexto.Item(23).Text = REmpleados!Consulta
            End If
        'Turno
            If IsNull(REmpleados!Turno) Then
                TxtTexto.Item(24).Text = ""
            Else
                TxtTexto.Item(24).Text = REmpleados!Turno
            End If
        'Uniforme
            If IsNull(REmpleados!Uniforme) Then
                TxtTexto.Item(25).Text = ""
            Else
                TxtTexto.Item(25).Text = REmpleados!Uniforme
            End If
        'Camisa
            If IsNull(REmpleados!Camisa) Then
                TxtTexto.Item(26).Text = ""
            Else
                TxtTexto.Item(26).Text = REmpleados!Camisa
            End If
        'Pantalon
            If IsNull(REmpleados!Pantalon) Then
                TxtTexto.Item(27).Text = ""
            Else
                TxtTexto.Item(27).Text = REmpleados!Pantalon
            End If
        'observaciones
            If IsNull(REmpleados!Observaciones) Then
                TxtTexto.Item(28).Text = ""
            Else
                TxtTexto.Item(28).Text = REmpleados!Observaciones
            End If
            
        If Err <> 0 Then
            'MsgBox Err.Description
        End If

End Sub

Public Sub Limpia_Campos()
                TxtTexto.Item(0).Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(5).Text = ""
                TxtTexto.Item(2).Text = ""
                TxtTexto.Item(4).Text = ""
                CboEstado.Text = ""
                MskFecAlt.Text = ""
                MskFecBaj.Text = ""
                TxtTexto.Item(15).Text = ""
                MskSue.Text = 0
                MskBon.Text = 0
                TxtTexto.Item(13).Text = ""
                TxtTexto.Item(6).Text = ""
                MskFecNac.Text = ""
                CboEstCiv.Text = ""
                TxtTexto.Item(7).Text = ""
                TxtTexto.Item(8).Text = 0
                TxtTexto.Item(9).Text = ""
                TxtTexto.Item(10).Text = ""
                TxtTexto.Item(11).Text = ""
                TxtTexto.Item(14).Text = ""
                Chk.Value = "0"
                TxtTexto.Item(12).Text = ""
                TxtTexto.Item(3).Text = ""
                MskBonInc.Text = 0
                
                TxtTexto.Item(16).Text = ""
                TxtTexto.Item(17).Text = ""
                TxtTexto.Item(18).Text = ""
                TxtTexto.Item(19).Text = ""
                TxtTexto.Item(20).Text = ""
                TxtTexto.Item(21).Text = ""
                TxtTexto.Item(22).Text = ""
                TxtTexto.Item(23).Text = ""
                TxtTexto.Item(24).Text = ""
                TxtTexto.Item(25).Text = ""
                TxtTexto.Item(26).Text = ""
                TxtTexto.Item(27).Text = ""
                TxtTexto.Item(28).Text = ""
                
        
End Sub



