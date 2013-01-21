VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CapturaProduccionDefectos 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Defectos De Produccion"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   Icon            =   "CapturaProduccionDefectos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   8625
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
      Height          =   6615
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DataGridBusqueda 
         Height          =   5415
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "click en encabezado de columna para indexar"
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   9551
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   16
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
         TabIndex        =   17
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   18
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "CapturaProduccionDefectos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7800
      MouseIcon       =   "CapturaProduccionDefectos.frx":237C
      Picture         =   "CapturaProduccionDefectos.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Ultimo Registro"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7440
      MouseIcon       =   "CapturaProduccionDefectos.frx":2CF0
      Picture         =   "CapturaProduccionDefectos.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Siguiente Registro"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "CapturaProduccionDefectos.frx":3664
      Picture         =   "CapturaProduccionDefectos.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Registro Anterior"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "CapturaProduccionDefectos.frx":3FD8
      Picture         =   "CapturaProduccionDefectos.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Primer Registro"
      Top             =   5880
      Width           =   375
   End
   Begin TabDlg.SSTab TabDefectos 
      Height          =   5535
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   255
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CapturaProduccionDefectos.frx":494C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameDefectos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaProduccionDefectos.frx":4C66
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CapturaProduccionDefectos.frx":50B8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameOpciones"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   38
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8281
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
         ColumnCount     =   6
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
         BeginProperty Column03 
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
         BeginProperty Column04 
            DataField       =   "Defecto"
            Caption         =   "Defecto"
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
            DataField       =   "Cantidad"
            Caption         =   "Cantidad"
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
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   494.929
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   959.811
            EndProperty
         EndProperty
      End
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
         Height          =   4695
         Left            =   -74880
         TabIndex        =   24
         Top             =   720
         Width           =   8085
         Begin VB.TextBox TxtBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   5760
            TabIndex        =   43
            ToolTipText     =   "Digite los datos para hacer la busqueda"
            Top             =   2160
            Width           =   2085
         End
         Begin VB.CommandButton CmdBotones 
            Caption         =   "Seleccion o Busqueda"
            Height          =   855
            Index           =   6
            Left            =   5760
            Picture         =   "CapturaProduccionDefectos.frx":550A
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   2640
            Width           =   2055
         End
         Begin VB.CommandButton CmdBotones 
            Caption         =   "Seleccionar Todos"
            Height          =   855
            Index           =   7
            Left            =   5760
            Picture         =   "CapturaProduccionDefectos.frx":594C
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   3600
            Width           =   2055
         End
         Begin VB.OptionButton OptFechas 
            Caption         =   "Fechas"
            Height          =   225
            Left            =   120
            TabIndex        =   12
            ToolTipText     =   " "
            Top             =   720
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptFechasYLinea 
            Caption         =   "Fechas Y Linea"
            Height          =   195
            Left            =   1080
            TabIndex        =   13
            ToolTipText     =   " "
            Top             =   720
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DTPFecFin 
            Height          =   255
            Left            =   6360
            TabIndex        =   39
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   51249155
            CurrentDate     =   37522
         End
         Begin MSComCtl2.DTPicker DTPFecIni 
            Height          =   255
            Left            =   4320
            TabIndex        =   40
            Top             =   1200
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   51249155
            CurrentDate     =   37522
         End
         Begin VB.Label Lbletiqueta 
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
            Left            =   3720
            TabIndex        =   46
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label Label2 
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
            Index           =   5
            Left            =   5760
            TabIndex        =   45
            Top             =   1200
            Width           =   510
         End
         Begin VB.Label Label2 
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
            Index           =   6
            Left            =   3720
            TabIndex        =   44
            Top             =   1200
            Width           =   555
         End
      End
      Begin VB.Frame FrameDefectos 
         Caption         =   "Datos De Defectos Por Tarima"
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
         Height          =   2535
         Left            =   120
         TabIndex        =   21
         Top             =   1680
         Width           =   8235
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   1
            Left            =   1200
            MaxLength       =   2
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   5
            ToolTipText     =   " "
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   4
            Left            =   1200
            MaxLength       =   4
            TabIndex        =   4
            ToolTipText     =   " "
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   3
            Left            =   1200
            TabIndex        =   3
            ToolTipText     =   " "
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   2
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   2
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo"
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
            Left            =   2880
            TabIndex        =   33
            Top             =   2160
            Width           =   390
         End
         Begin VB.Label LblTipo 
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
            TabIndex        =   32
            Top             =   2160
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   31
            Top             =   2160
            Width           =   630
         End
         Begin VB.Label LblDefecto 
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
            Left            =   2880
            TabIndex        =   30
            Top             =   1800
            Width           =   5175
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Defecto"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   570
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tarima"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Ficha Tecnica"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   1095
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
            Left            =   2880
            TabIndex        =   26
            Top             =   720
            Width           =   5175
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
            Left            =   2880
            TabIndex        =   25
            Top             =   1080
            Width           =   5175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label2 
            Caption         =   "Linea"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6360
      MouseIcon       =   "CapturaProduccionDefectos.frx":5C56
      Picture         =   "CapturaProduccionDefectos.frx":6098
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5280
      MouseIcon       =   "CapturaProduccionDefectos.frx":65B3
      Picture         =   "CapturaProduccionDefectos.frx":69F5
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "CapturaProduccionDefectos.frx":6FBD
      Picture         =   "CapturaProduccionDefectos.frx":73FF
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   3120
      MouseIcon       =   "CapturaProduccionDefectos.frx":7936
      Picture         =   "CapturaProduccionDefectos.frx":7D78
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2040
      MouseIcon       =   "CapturaProduccionDefectos.frx":82D4
      Picture         =   "CapturaProduccionDefectos.frx":8716
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5760
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   960
      MouseIcon       =   "CapturaProduccionDefectos.frx":8AED
      Picture         =   "CapturaProduccionDefectos.frx":8F2F
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1000
   End
End
Attribute VB_Name = "CapturaProduccionDefectos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim VTipo As String
Dim VCantidad As Long

Dim BFichaTecnica As Boolean
Dim BLinea As Boolean
Dim BDefecto As Boolean
Dim BEditar As Boolean

Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaFichaTecnica2 As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaDefectos As New ADODB.Recordset
Dim RBuscaDefectosMen As New ADODB.Recordset
Dim RBuscaDefectosMay As New ADODB.Recordset
Dim RBuscaDefectosCri As New ADODB.Recordset
Dim RBuscaAtributos As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RDefectos As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VSumaCriticos As Long
Dim VSumaMayores As Long
Dim VSumaMenores As Long
Dim VTexto As String
Dim VMensaje As String



Sub botones()
    If Bandera = True Then
         FrameDefectos.Enabled = True
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

         FrameOpciones.Visible = False
         DataGrid1.Visible = False
    Else
         FrameDefectos.Enabled = False
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

         FrameOpciones.Visible = True
         DataGrid1.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
    
    
        'AGREGAR
        If Index = 0 Then
                Bandera = True
                botones
                Limpia_Campos
                
                'HABILITA LA LLAVE
                TxtTexto.Item(0).Enabled = True
                TxtTexto.Item(1).Enabled = True
                TxtTexto.Item(2).Enabled = True
                TxtTexto.Item(3).Enabled = True
                TxtTexto.Item(4).Enabled = True
                                
                BEditar = False
                 
                TxtTexto.Item(0).Text = VPFecha
                TxtTexto.Item(1).Text = VPLinea
                TxtTexto.Item(2).Text = VPFicha
                TxtTexto.Item(3).Text = VPTarima
                
                TxtTexto.Item(4).SetFocus
                
        'EDITAR
        ElseIf Index = 1 Then
        
                Bandera = True
                botones
                'DESABILITA LA LLAVE
                TxtTexto.Item(0).Enabled = False
                TxtTexto.Item(1).Enabled = False
                TxtTexto.Item(2).Enabled = False
                TxtTexto.Item(3).Enabled = False
                TxtTexto.Item(4).Enabled = False
                                
                TxtTexto.Item(4).SetFocus
                BEditar = True
        'GRABAR
        ElseIf Index = 2 Then
                        
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                        Else
                             TxtTexto.Item(0).Text = Format(TxtTexto.Item(0).Text, "dd/mm/yyyy")
                        End If
                        
                        'REVISA LA FECHA
                        If Not IsDate(TxtTexto.Item(0).Text) Then
                            MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                        End If
                                                
                        'REVISA Tarima
                        If Not IsNumeric(TxtTexto.Item(3).Text) Then
                            MsgBox "Tarima Debe Ser Numerica", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(3).SetFocus
                            Exit Sub
                        End If
                        
                        'REVISA Cantidad
                        If Not IsNumeric(TxtTexto.Item(5).Text) Then
                            MsgBox "Cantidad debe Ser Numerica", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(5).SetFocus
                            Exit Sub
                        End If
                        
                    'AGREGAR
                    If BEditar = False Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "Values(#" & Format(TxtTexto.Item(0).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = "Values(To_Date('" & Format(TxtTexto.Item(0).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            
                            VTexto = VTexto & "'" & TxtTexto.Item(1).Text & "', '" 'LINEA
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(3).Text & ", '" 'TARIMA
                            VTexto = VTexto & TxtTexto.Item(4).Text & "', " 'DEFECTO
                            VTexto = VTexto & TxtTexto.Item(5).Text & ")" 'CANTIDAD
                            
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into ProduccionConDefectos " & VTexto
                    'EDITAR
                    Else
                            VTexto = "Cantidad = " & TxtTexto.Item(5).Text & " " 'CANTIDAD
                            
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "Where Fec_Prd = #" & Format(TxtTexto.Item(0).Text, "mm/dd/yyyy") & "# And "
                            Else 'ORACLE
                                VTexto = VTexto & "Where Fec_Prd = TO_DATE('" & Format(TxtTexto.Item(0).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy') And "
                            End If
                                VTexto = VTexto & "Linea = '" & TxtTexto.Item(1) & "' And Esp_Tec = '" & TxtTexto.Item(2) & "' And "
                                VTexto = VTexto & "Tarima = " & TxtTexto.Item(3) & " And Defecto = '" & TxtTexto.Item(4) & "'"
                        
                            Conexion.Execute "UPDATE ProduccionConDefectos SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Fecha, Linea, FichaTecnica, Trima y Defecto Ya Existe Para Esta Tarima", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Fecha, Linea, FichaTecnica, Trima y Defecto Ya Existe Para Esta Tarima", vbOKOnly + vbInformation, "Informacion"
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
                        TxtTexto.Item(1).Enabled = True
                        TxtTexto.Item(2).Enabled = True
                        TxtTexto.Item(3).Enabled = True
                        TxtTexto.Item(4).Enabled = True
                                                
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RDefectos.Requery
                        RDefectos.MoveLast
                        Llena_Campos

        'CANCELAR
        ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    Llena_Campos
                    'HABILITA LA LLAVE
                    TxtTexto.Item(0).Enabled = True
                    TxtTexto.Item(1).Enabled = True
                    TxtTexto.Item(2).Enabled = True
                    TxtTexto.Item(3).Enabled = True
                    TxtTexto.Item(4).Enabled = True
                    
        ElseIf Index = 4 Then ' BORRAR
        
                On Error Resume Next
            VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If VMensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RDefectos.Delete
                        
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
                        RDefectos.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RDefectos.MoveNext
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
                    Set RDefectos = New ADODB.Recordset
                    If OptFechas.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RDefectos, "Select * From ProduccionConDefectos Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#")
                        Else 'ORACLE
                            Call Abrir_Recordset(RDefectos, "Select * From ProduccionConDefectos Where Fecha >= TO_DATE('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= TO_DATE('" & DTPFecFin.Value & "', 'dd/mm/yyyy')")
                        End If
                    ElseIf OptFechasYLinea.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RDefectos, "Select * From ProduccionConDefectos Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea Like '" & TxtBuscar.Text & "%'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RDefectos, "Select * From ProduccionConDefectos Where Fecha >= TO_DATE('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fecha <= TO_DATE('" & DTPFecFin.Value & "', 'dd/mm/yyyy') And UPPER(Linea) Like '" & UCase(TxtBuscar.Text) & "%'")
                        End If
                    End If
                    Set DataGrid1.DataSource = RDefectos
                    TabDefectos.Tab = 1
        ElseIf Index = 7 Then 'ACTUALIZAR
                    Call Abrir_Recordset(RDefectos, "Select * From ProduccionConDefectos")
                    Set DataGrid1.DataSource = RDefectos
                    TabDefectos.Tab = 1
        End If
    
    

End Sub

Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RDefectos.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RDefectos.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RDefectos.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RDefectos.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RDefectos.BOF Then
        RDefectos.MoveFirst
    ElseIf RDefectos.EOF Then
        RDefectos.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub


Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub


Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
        RDefectos.Sort = RDefectos.Fields(ColIndex).Name
End Sub


Private Sub DataGridBusqueda_DblClick()
            If BFichaTecnica = True Then
                TxtTexto.Item(2).Text = DataGridBusqueda.Columns(0).Text
                TxtTexto.Item(2).SetFocus
            ElseIf BLinea = True Then
                TxtTexto.Item(1).Text = DataGridBusqueda.Columns(0).Text
                TxtTexto.Item(1).SetFocus
            ElseIf BDefecto = True Then
                TxtTexto.Item(4).Text = DataGridBusqueda.Columns(0).Text
                TxtTexto.Item(4).SetFocus
            End If
                FrameBusqueda.Visible = False
End Sub

Private Sub DataGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DataGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                If BFichaTecnica = True Then
                    TxtTexto.Item(2).Text = DataGridBusqueda.Columns(0).Text
                    TxtTexto.Item(2).SetFocus
                ElseIf BLinea = True Then
                    TxtTexto.Item(1).Text = DataGridBusqueda.Columns(0).Text
                    TxtTexto.Item(1).SetFocus
                ElseIf BDefecto = True Then
                    TxtTexto.Item(4).Text = DataGridBusqueda.Columns(0).Text
                    TxtTexto.Item(4).SetFocus
                End If
                    FrameBusqueda.Visible = False
            End If
                    
End Sub



Private Sub Form_Load()
        Set RDefectos = New ADODB.Recordset
        Call Abrir_Recordset(RDefectos, "Select * From ProduccionConDefectos order by Fec_Prd")
        Set DataGrid1.DataSource = RDefectos
        Llena_Campos
        
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date
                
End Sub

Private Sub Form_Unload(Cancel As Integer)
        'BUSCA LOS DEFECTOS TIPO CRITICO OSEA '2'
        Set RBuscaDefectosCri = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaDefectosCri, "Select Sum(PD.Cantidad) From ProduccionConDefectos PD, Defectos D Where PD.Esp_Tec = '" & VPFicha & "' And PD.Fec_Prd = #" & Format(VPFecha, "mm/dd/yyyy") & "# And PD.Linea = '" & VPLinea & "' And PD.Tarima = " & VPTarima & " And D.Tipo = '2' And D.Defecto = PD.Defecto")
            Else 'ORACLE
                    Call Abrir_Recordset(RBuscaDefectosCri, "Select Sum(PD.Cantidad) From ProduccionConDefectos PD, Defectos D Where PD.Esp_Tec = '" & VPFicha & "' And PD.Fec_Prd = TO_DATE('" & VPFecha & "', 'dd/mm/yyyy')" & " And PD.Linea = '" & VPLinea & "' And PD.Tarima = " & VPTarima & " And D.Tipo = '2' And D.Defecto = PD.Defecto")
            End If
            If RBuscaDefectosCri.RecordCount > 0 Then
                If Not IsNull(RBuscaDefectosCri(0)) Then
                    VSumaCriticos = RBuscaDefectosCri(0)
                Else
                    VSumaCriticos = 0
                End If
            Else
                VSumaCriticos = 0
            End If
        
        'BUSCA LOS DEFECTOS TIPO CRITICO OSEA '1'
        Set RBuscaDefectosMay = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaDefectosMay, "Select Sum(PD.Cantidad) From ProduccionConDefectos PD, Defectos D Where PD.Esp_Tec = '" & VPFicha & "' And PD.Fec_Prd = #" & Format(VPFecha, "mm/dd/yyyy") & "# And PD.Linea = '" & VPLinea & "' And PD.Tarima = " & VPTarima & " And D.Tipo = '1' And D.Defecto = PD.Defecto")
            Else 'ORACLE
                    Call Abrir_Recordset(RBuscaDefectosMay, "Select Sum(PD.Cantidad) From ProduccionConDefectos PD, Defectos D Where PD.Esp_Tec = '" & VPFicha & "' And PD.Fec_Prd = TO_DATE('" & VPFecha & "', 'dd/mm/yyyy')" & " And PD.Linea = '" & VPLinea & "' And PD.Tarima = " & VPTarima & " And D.Tipo = '1' And D.Defecto = PD.Defecto")
            End If
        
            If RBuscaDefectosMay.RecordCount > 0 Then
                If Not IsNull(RBuscaDefectosMay(0)) Then
                    VSumaMayores = RBuscaDefectosMay(0)
                Else
                    VSumaMayores = 0
                End If
            Else
                VSumaMayores = 0
            End If
                        
        'BUSCA LOS DEFECTOS TIPO CRITICO OSEA '0'
        Set RBuscaDefectosMen = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaDefectosMen, "Select Sum(PD.Cantidad) From ProduccionConDefectos PD, Defectos D Where PD.Esp_Tec = '" & VPFicha & "' And PD.Fec_Prd = #" & Format(VPFecha, "mm/dd/yyyy") & "# And PD.Linea = '" & VPLinea & "' And PD.Tarima = " & VPTarima & " And D.Tipo = '0' And D.Defecto = PD.Defecto")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaDefectosMen, "Select Sum(PD.Cantidad) From ProduccionConDefectos PD, Defectos D Where PD.Esp_Tec = '" & VPFicha & "' And PD.Fec_Prd = TO_DATE('" & VPFecha & "', 'dd/mm/yyyy')" & " And PD.Linea = '" & VPLinea & "' And PD.Tarima = " & VPTarima & " And D.Tipo = '0' And D.Defecto = PD.Defecto")
            End If
        If RBuscaDefectosMen.RecordCount > 0 Then
                If Not IsNull(RBuscaDefectosMen(0)) Then
                    VSumaMenores = RBuscaDefectos(0)
                Else
                    VSumaMenores = 0
                End If
            Else
                VSumaMenores = 0
            End If
                        
                        
        'BUSCA EL ATRIBUTO DE LA FICHA TECNICA
        Set RBuscaFichaTecnica2 = New ADODB.Recordset
        Call Abrir_Recordset(RBuscaFichaTecnica2, "Select Atributos From FichaTecnica where Esp_Tec = '" & VPFicha & "'")
            If RBuscaFichaTecnica2.RecordCount > 0 Then
                'BUSCA LA CANTIDAD DE TIPOS DE DEFECTO QUE PUEDE ACEPTAR
                Set RBuscaAtributos = New ADODB.Recordset
                Call Abrir_Recordset(RBuscaAtributos, "Select Criticos, Mayores, Menores From Atributos Where Codigo = '" & RBuscaFichaTecnica2(0) & "'")
                    If RBuscaAtributos.RecordCount > 0 Then
                            If VSumaCriticos > RBuscaAtributos(0) Then
                                    VPCalidad = "R"
                            ElseIf VSumaMayores > RBuscaAtributos(1) Then
                                    VPCalidad = "R"
                            ElseIf VSumaMenores > RBuscaAtributos(2) Then
                                    VPCalidad = " R"
                            Else
                                    VPCalidad = "A"
                            End If
                                    
                    End If
            Else
                    LblFichaTecnica.Caption = ""
            End If
            
            'BUSCA LA TARIMA Y ACTUALIZA LA CALIDAD
            If GOrigenDeDatos = "AmaproAccess" Then
                Conexion.Execute ("Update Produccion Set Calidad = '" & VPCalidad & "' Where Linea = '" & VPLinea & "' and Esp_tec = '" & VPFicha & "' and Fec_prd = #" & Format(VPFecha, "mm/dd/yyyy") & "# and Tarima = " & VPTarima)
            Else
                Conexion.Execute ("Update Produccion Set Calidad = '" & VPCalidad & "' Where Linea = '" & VPLinea & "' and Esp_tec = '" & VPFicha & "' and Fec_prd = To_Date('" & VPFecha & "', 'dd/mm/yyyy')" & " and Tarima = " & VPTarima)
            End If
                                   
                        
End Sub

Private Sub OptFechas_Click()
        Lbletiqueta.Caption = ""
End Sub

Private Sub OptFechasYlinea_Click()
        Lbletiqueta.Caption = "Linea"
End Sub

Private Sub TabDefectos_Click(PreviousTab As Integer)
        If TabDefectos.Tab = 0 Then
            CmdBotones.Item(4).Enabled = True
            If CmdBotones.Item(2).Enabled = False Then
                Llena_Campos
            End If
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
    'LINEA
    If BLinea = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where Descrip Like '%" & TxtBusqueda.Text & "%'")
                Else ' ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where Linea Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
    'FICHA TECNICA
    ElseIf BFichaTecnica = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Descrip Like '%" & TxtBusqueda.Text & "%' Order By Esp_Tec")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%' Order By Esp_Tec")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '%" & TxtBusqueda.Text & "%' Order By Esp_Tec")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%' Order By Esp_Tec")
                End If
            End If
    'DEFECTO
    ElseIf BDefecto = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos where Descrip Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos where Defecto Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos where UPPER(Defecto) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
    
    End If
            
            Set DataGridBusqueda.DataSource = RBusqueda
            DataGridBusqueda.Columns(1).Width = "4000"

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
        If Index = 2 Then
        'BUSCA LA DESCRIPCION DE FICHA TECNICA
            Set RBuscaFichaTecnica = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where Esp_Tec = '" & TxtTexto.Item(2).Text & "'")
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblFichaTecnica.Caption = ""
                End If
        'BUSCA LA DESCRIPCION DE LINEA
        ElseIf Index = 1 Then
            Set RBuscaLinea = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
                
        'BUSCA LA DESCRIPCION DE DEFECTOS
        ElseIf Index = 4 Then
            Set RBuscaDefectos = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaDefectos, "Select Descrip, Tipo From Defectos Where Defecto = '" & TxtTexto.Item(4).Text & "'")
                If RBuscaDefectos.RecordCount > 0 Then
                    LblDefecto.Caption = RBuscaDefectos!Descrip
                    LblTipo.Caption = RBuscaDefectos!Tipo
                Else
                    LblDefecto.Caption = ""
                    LblTipo.Caption = ""
                End If
        End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        If (Index = 0 Or Index = 2 Or Index = 4) Then
            Set RBusqueda = New ADODB.Recordset
        End If
        
        'SI ELIGE FICHA TECNICA
        If Index = 0 Then
            BFichaTecnica = True
            BLinea = False
            BDefecto = False
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size From FichaTecnica Where Activa = -1")
        'LINEAS
        ElseIf Index = 2 Then
            BFichaTecnica = False
            BLinea = True
            BDefecto = False
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
        'DEFECTOS
        ElseIf Index = 4 Then
            BFichaTecnica = False
            BLinea = False
            BDefecto = True
            Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos")
        End If
        
        If (Index = 0 Or Index = 2 Or Index = 4) Then
            Set DataGridBusqueda.DataSource = RBusqueda
            DataGridBusqueda.Columns(1).Width = "4000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
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
        If (Index = 0 Or Index = 2 Or Index = 4) Then
            Set RBusqueda = New ADODB.Recordset
        End If
        
        'SI ELIGE FICHA TECNICA
        If Index = 0 Then
            BFichaTecnica = True
            BLinea = False
            BDefecto = False
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size From FichaTecnica Where Activa = -1")
        'LINEAS
        ElseIf Index = 2 Then
            BFichaTecnica = False
            BLinea = True
            BDefecto = False
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
        'DEFECTOS
        ElseIf Index = 4 Then
            BFichaTecnica = False
            BLinea = False
            BDefecto = True
            Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos")
        End If
        
        If (Index = 0 Or Index = 2 Or Index = 4) Then
            Set DataGridBusqueda.DataSource = RBusqueda
            DataGridBusqueda.Columns(1).Width = "4000"
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
        End If
      
    End If
    
End Sub
Public Sub Llena_Campos()
On Error Resume Next
        'FECHA
            If IsNull(RDefectos!fec_prd) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = RDefectos!fec_prd
            End If
        'LINEA
            If IsNull(RDefectos!Linea) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = RDefectos!Linea
            End If
        'FICHA
            If IsNull(RDefectos!Esp_Tec) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RDefectos!Esp_Tec
            End If
        'TARIMA
            If IsNull(RDefectos!Tarima) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = RDefectos!Tarima
            End If
        'DEFECTO
            If IsNull(RDefectos!Defecto) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = RDefectos!Defecto
            End If
        'CANTIDAD
            If IsNull(RDefectos!Cantidad) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = RDefectos!Cantidad
            End If
            
        
        If Err <> 0 Then
            
        End If

End Sub

Public Sub Limpia_Campos()
        
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(3).Text = ""
        TxtTexto.Item(4).Text = ""
        TxtTexto.Item(5).Text = ""
        TxtTexto.Item(8).Text = ""
        
End Sub

