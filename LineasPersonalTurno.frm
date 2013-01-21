VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form LineasPersonalTurno 
   BackColor       =   &H00FF8080&
   Caption         =   "Personal Entrega Turno"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "LineasPersonalTurno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   8415
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
      Height          =   8295
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   33
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
         TabIndex        =   34
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   35
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   6015
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7440
         Picture         =   "LineasPersonalTurno.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   7095
         Left            =   120
         TabIndex        =   36
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   12515
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
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "LineasPersonalTurno.frx":293C
      Picture         =   "LineasPersonalTurno.frx":2D7E
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Primer Registro"
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "LineasPersonalTurno.frx":32B0
      Picture         =   "LineasPersonalTurno.frx":36F2
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Registro Anterior"
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7440
      MouseIcon       =   "LineasPersonalTurno.frx":3C24
      Picture         =   "LineasPersonalTurno.frx":4066
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Siguiente Registro"
      Top             =   7560
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7800
      MouseIcon       =   "LineasPersonalTurno.frx":4598
      Picture         =   "LineasPersonalTurno.frx":49DA
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Ultimo Registro"
      Top             =   7560
      Width           =   375
   End
   Begin TabDlg.SSTab TabPuestos 
      Height          =   7335
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12938
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   16744576
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "LineasPersonalTurno.frx":4F0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePuestos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "LineasPersonalTurno.frx":5226
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "LineasPersonalTurno.frx":5678
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "CmdBuscar(0)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "CmdBuscar(1)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "TxtBuscar"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "FrameOpciones"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame FrameOpciones 
         Caption         =   "Opciones de Busqueda"
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
         Height          =   740
         Left            =   -74760
         TabIndex        =   38
         Top             =   2160
         Width           =   2805
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Linea"
            Height          =   225
            Left            =   120
            TabIndex        =   39
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6495
         Left            =   -74880
         TabIndex        =   23
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   11456
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
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -69480
         TabIndex        =   21
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   3840
         Width           =   2085
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -69480
         Picture         =   "LineasPersonalTurno.frx":5ACA
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5280
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -69480
         Picture         =   "LineasPersonalTurno.frx":5DD4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   4320
         Width           =   2055
      End
      Begin VB.Frame FramePuestos 
         Caption         =   "Datos De El Personal"
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
         Height          =   4815
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   8175
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   23
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   9
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   4440
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   22
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   8
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   4080
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   21
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   7
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   3720
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   20
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   6
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   3360
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   19
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   5
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   2520
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   18
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   4
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   2160
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   17
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   3
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   1800
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   16
            Left            =   2280
            MaxLength       =   50
            TabIndex        =   2
            ToolTipText     =   "doble click o signo '+' para ayuda"
            Top             =   1440
            Width           =   3015
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   2280
            MaxLength       =   1
            TabIndex        =   1
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   720
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   285
            Index           =   3
            Left            =   6120
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   720
            Width           =   1935
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   720
            MaxLength       =   15
            TabIndex        =   0
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Mecanico De Linea"
            Height          =   195
            Index           =   22
            Left            =   120
            TabIndex        =   49
            Top             =   3720
            Width           =   1995
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Insp. De Aseg. Cal."
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   48
            Top             =   4080
            Width           =   1965
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Supervisor De Turno"
            Height          =   195
            Index           =   20
            Left            =   120
            TabIndex        =   47
            Top             =   4440
            Width           =   2070
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Del Operador"
            Height          =   195
            Index           =   19
            Left            =   120
            TabIndex        =   46
            Top             =   3360
            Width           =   1545
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Grupo De Manufactura Que Recibe La Linea"
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
            Index           =   26
            Left            =   120
            TabIndex        =   45
            Top             =   3000
            Width           =   3810
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Mecanico De Linea"
            Height          =   195
            Index           =   25
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   1995
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Insp. De Aseg. Cal."
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   43
            Top             =   2160
            Width           =   1965
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Supervisor De Turno"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   42
            Top             =   2520
            Width           =   2070
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Grupo De Manufactura Que Entrega La Linea"
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
            Index           =   18
            Left            =   120
            TabIndex        =   41
            Top             =   1080
            Width           =   3870
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Del Operador"
            Height          =   195
            Index           =   17
            Left            =   120
            TabIndex        =   40
            Top             =   1440
            Width           =   1545
         End
         Begin VB.Label LblPro 
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
            Height          =   285
            Left            =   5280
            TabIndex        =   31
            Top             =   360
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   30
            Top             =   720
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   2
            Left            =   5520
            TabIndex        =   29
            Top             =   720
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Grupo"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -71520
         TabIndex        =   20
         Top             =   3840
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6360
      MouseIcon       =   "LineasPersonalTurno.frx":6216
      Picture         =   "LineasPersonalTurno.frx":6658
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5280
      MouseIcon       =   "LineasPersonalTurno.frx":6B73
      Picture         =   "LineasPersonalTurno.frx":6FB5
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "LineasPersonalTurno.frx":757D
      Picture         =   "LineasPersonalTurno.frx":79BF
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   3120
      MouseIcon       =   "LineasPersonalTurno.frx":7EF6
      Picture         =   "LineasPersonalTurno.frx":8338
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2040
      MouseIcon       =   "LineasPersonalTurno.frx":8894
      Picture         =   "LineasPersonalTurno.frx":8CD6
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7440
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   960
      MouseIcon       =   "LineasPersonalTurno.frx":90AD
      Picture         =   "LineasPersonalTurno.frx":94EF
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7440
      Width           =   1000
   End
End
Attribute VB_Name = "LineasPersonalTurno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim RLineasPersonal As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset

Dim BEditar As Boolean
Dim VLlave1 As String

Dim VTexto As String


Sub botones()
    If Bandera = True Then
         FramePuestos.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         TxtTexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         
         CmdBuscar.Item(0).Visible = False
         CmdBuscar.Item(1).Visible = False
        
         
         FrameOpciones.Visible = False
         DataGrid1.Visible = False
    Else
         FramePuestos.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         
         CmdBuscar.Item(0).Visible = True
         CmdBuscar.Item(1).Visible = True

         
         FrameOpciones.Visible = True
         DataGrid1.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
            'AGREGAR
            If Index = 0 Then
                    TabPuestos.Tab = 0
                    Bandera = True
                    botones
                    Limpia_Campos
                    TxtTexto.Item(0).SetFocus
                    TxtTexto.Item(3).Text = GUsuario
                    
                    BEditar = False
            'EDITAR
            ElseIf Index = 1 Then
                    TabPuestos.Tab = 0
                    Bandera = True
                    botones
                    VLlave1 = TxtTexto.Item(0).Text
                    TxtTexto.Item(0).SetFocus
                    TxtTexto.Item(3).Text = GUsuario
                    BEditar = True
            'GRABAR
            ElseIf Index = 2 Then
                    'REVISA EL CODIGO
                    If TxtTexto.Item(0).Text = "" Then
                        MsgBox "Grupo No Puede Estar Vacia", vbOKOnly + vbInformation, "Informacion"
                        TxtTexto.Item(0).SetFocus
                        Exit Sub
                    End If
                    
                    'If TxtTexto.Item(1).Text = "" Then
                    '    MsgBox "Turno No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
                    '    TxtTexto.Item(1).SetFocus
                    '    Exit Sub
                    'End If
                                        
                    
                    If BEditar = False Then 'ESTA AGREGANDO UN REGISTRO
                            VTexto = "'" & TxtTexto.Item(0).Text & "', '"
                            VTexto = VTexto & "1" & "', '"
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', '"
                            VTexto = VTexto & TxtTexto.Item(16).Text & "', '"  'OPERADOR ENTREGA
                            VTexto = VTexto & TxtTexto.Item(17).Text & "', '"  'MECANICO ENTREGA
                            VTexto = VTexto & TxtTexto.Item(18).Text & "', '"  'INSPECTOR ENTREGA
                            VTexto = VTexto & TxtTexto.Item(19).Text & "', '"  'SUPERVISOR ENTREGA
                            VTexto = VTexto & TxtTexto.Item(20).Text & "', '"  'OPERADOR RECIBE
                            VTexto = VTexto & TxtTexto.Item(21).Text & "', '"  'MECANICO RECIBE
                            VTexto = VTexto & TxtTexto.Item(22).Text & "', '"  'INSPECTOR RECIBE
                            VTexto = VTexto & TxtTexto.Item(23).Text & "'"  'SUPERVISOR RECIBE
                         
                         
                         Conexion.Execute "Insert Into LineasPersonalTurno Values(" & VTexto & ")"
                         
                    Else 'ESTA EDITANDO UN REGISTRO
                            VTexto = "Linea = '" & TxtTexto.Item(0).Text & "', "
                            'VTexto = VTexto & "Turno = '" & TxtTexto.Item(1).Text & "', "
                            VTexto = VTexto & "Usuario = '" & TxtTexto.Item(3).Text & "', "
                            VTexto = VTexto & "OperadorEntrega = '" & TxtTexto.Item(16).Text & "', "  'OPERADOR ENTREGA
                            VTexto = VTexto & "MecanicoEntrega = '" & TxtTexto.Item(17).Text & "', "  'MECANICO ENTREGA
                            VTexto = VTexto & "InspectorEntrega = '" & TxtTexto.Item(18).Text & "', "  'INSPECTOR ENTREGA
                            VTexto = VTexto & "SupervisorEntrega = '" & TxtTexto.Item(19).Text & "', "  'SUPERVISOR ENTREGA
                            VTexto = VTexto & "OperadorRecibe = '" & TxtTexto.Item(20).Text & "', "  'OPERADOR RECIBE
                            VTexto = VTexto & "MecanicoRecibe = '" & TxtTexto.Item(21).Text & "', "  'MECANICO RECIBE
                            VTexto = VTexto & "InspectorRecibe = '" & TxtTexto.Item(22).Text & "', "  'INSPECTOR RECIBE
                            VTexto = VTexto & "SupervisorRecibe = '" & TxtTexto.Item(23).Text & "'"  'SUPERVISOR RECIBE
                    
                         Conexion.Execute "UPDATE LineasPersonalTurno SET " & VTexto & " Where Linea = '" & VLlave1 & "'"
                    End If
                                            
                      'SI ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
                        RLineasPersonal.Requery
                        RLineasPersonal.MoveLast
                        Llena_Campos
            'CANCELAR
            ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    Llena_Campos
                    
            'BORRAR
            ElseIf Index = 4 Then
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RLineasPersonal.Delete
                        
                        
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        
                        'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RLineasPersonal.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RLineasPersonal.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                        
                    End If
            'SALIDA
            ElseIf Index = 5 Then
                    Unload Me
            End If
        
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
On Error Resume Next
MousePointer = 11
    If Index = 1 Then
        RLineasPersonal.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RLineasPersonal.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RLineasPersonal.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RLineasPersonal.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RLineasPersonal.BOF Then
        RLineasPersonal.MoveFirst
    ElseIf RLineasPersonal.EOF Then
        RLineasPersonal.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        Set RLineasPersonal = New ADODB.Recordset
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                    Call Abrir_Recordset(RLineasPersonal, "Select * from LineasPersonalTurno where Linea Like '%" & TxtBuscar.Text & "%'")
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RLineasPersonal, "Select * From LineasPersonalTurno")
        End If
        Set DataGrid1.DataSource = RLineasPersonal
    
        TabPuestos.Tab = 1
End Sub


Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
                RLineasPersonal.Sort = RLineasPersonal.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If
    
End Sub


Private Sub DBGridBusqueda_DblClick()
                TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(0).SetFocus
                FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
                TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(0).SetFocus
                FrameBusqueda.Visible = False
        End If
End Sub

Private Sub Form_Load()
        Set RLineasPersonal = New ADODB.Recordset
        Call Abrir_Recordset(RLineasPersonal, "Select * From LineasPersonalTurno")
        Set DataGrid1.DataSource = RLineasPersonal
        Llena_Campos
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        RLineasPersonal.Close
        Set RLineasPersonal = Nothing
        If Err <> 0 Then
        End If
End Sub



Private Sub OptCodigo_Click()
            Lbletiqueta.Caption = "Codigo"
End Sub



Private Sub TabPuestos_Click(PreviousTab As Integer)
    If TabPuestos.Tab = 0 Then
        If CmdBotones.Item(2).Enabled = False Then
            Llena_Campos
        End If
        
        CmdBotones.Item(4).Enabled = True
    Else
        CmdBotones.Item(4).Enabled = False
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
                        'DESCRIPCION
                        If OptBusqueda.Item(0).Value = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where Descrip Like '%" & TxtBusqueda.Text & "%'")
                        'CODIGO
                        ElseIf OptBusqueda.Item(1).Value = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where Linea Like '%" & TxtBusqueda.Text & "%'")
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
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
'    If Index = 0 Then
 '       Set RBuscaLinea = New ADODB.Recordset
  '          Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(0).Text & "'")
   '             If RBuscaLinea.RecordCount > 0 Then
    '                LblPro.Caption = RBuscaLinea!Descrip
     '           Else
      '              LblPro.Caption = ""
       '         End If
    'End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        'If Index = 0 Then
        '    Set RBusqueda = New ADODB.Recordset
       '
       '     Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
       '         Set DBGridBusqueda.DataSource = RBusqueda
       '         DBGridBusqueda.Columns(1).Width = "4000"
       '         FrameBusqueda.Visible = True
       '         TxtBusqueda.Text = ""
       '         TxtBusqueda.SetFocus
       ' End If
        
End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
        TxtTexto.Item(Index).SelStart = 0
        TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        'If KeyAscii = 43 Then
            'If Index = 0 Then
            '    Set RBusqueda = New ADODB.Recordset
       '
       '         Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
       '             Set DBGridBusqueda.DataSource = RBusqueda
       '             DBGridBusqueda.Columns(1).Width = "4000"
       '             FrameBusqueda.Visible = True
       '             TxtBusqueda.Text = ""
       '             TxtBusqueda.SetFocus
       '     End If
        'End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
        If RLineasPersonal.RecordCount > 0 Then
            TxtTexto.Item(0).Text = RLineasPersonal!Linea
            TxtTexto.Item(1).Text = RLineasPersonal!Turno
            TxtTexto.Item(3).Text = RLineasPersonal!Usuario
            'OPERADOR ENTREGA
            If IsNull(RLineasPersonal!OperadorEntrega) Then
                TxtTexto.Item(16).Text = ""
            Else
                TxtTexto.Item(16).Text = RLineasPersonal!OperadorEntrega
            End If
            'MECANICO ENTREGA
            If IsNull(RLineasPersonal!MecanicoEntrega) Then
                TxtTexto.Item(17).Text = ""
            Else
                TxtTexto.Item(17).Text = RLineasPersonal!MecanicoEntrega
            End If
            'INSPECTOR ENTREGA
            If IsNull(RLineasPersonal!InspectorEntrega) Then
                TxtTexto.Item(18).Text = ""
            Else
                TxtTexto.Item(18).Text = RLineasPersonal!InspectorEntrega
            End If
            'SUPERVISOR ENTREGA
            If IsNull(RLineasPersonal!SupervisorEntrega) Then
                TxtTexto.Item(19).Text = ""
            Else
                TxtTexto.Item(19).Text = RLineasPersonal!SupervisorEntrega
            End If
            
            'OPERADOR RECIBE
            If IsNull(RLineasPersonal!OperadorRecibe) Then
                TxtTexto.Item(20).Text = ""
            Else
                TxtTexto.Item(20).Text = RLineasPersonal!OperadorRecibe
            End If
            'MECANICO RECIBE
            If IsNull(RLineasPersonal!MecanicoRecibe) Then
                TxtTexto.Item(21).Text = ""
            Else
                TxtTexto.Item(21).Text = RLineasPersonal!MecanicoRecibe
            End If
            'INSPECTOR RECIBE
            If IsNull(RLineasPersonal!InspectorRecibe) Then
                TxtTexto.Item(22).Text = ""
            Else
                TxtTexto.Item(22).Text = RLineasPersonal!InspectorRecibe
            End If
            'SUPERVISOR RECIBE
            If IsNull(RLineasPersonal!SupervisorRecibe) Then
                TxtTexto.Item(23).Text = ""
            Else
                TxtTexto.Item(23).Text = RLineasPersonal!SupervisorRecibe
            End If
        Else
            TxtTexto.Item(0).Text = ""
            TxtTexto.Item(1).Text = ""
            TxtTexto.Item(2).Text = ""
            TxtTexto.Item(3).Text = ""
        End If
        
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
            TxtTexto.Item(0).Text = ""
            TxtTexto.Item(1).Text = ""
            TxtTexto.Item(3).Text = ""
        'OPERADOR ENTREGA
            TxtTexto.Item(16).Text = ""
            'MECANICO ENTREGA
            TxtTexto.Item(17).Text = ""
            'INSPECTOR ENTREGA
            TxtTexto.Item(18).Text = ""
            'SUPERVISOR ENTREGA
            TxtTexto.Item(19).Text = ""
            'OPERADOR RECIBE
            TxtTexto.Item(20).Text = ""
            'MECANICO RECIBE
            TxtTexto.Item(21).Text = ""
            'INSPECTOR RECIBE
            TxtTexto.Item(22).Text = ""
            'SUPERVISOR RECIBE
            TxtTexto.Item(23).Text = ""
End Sub
