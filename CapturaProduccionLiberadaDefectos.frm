VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CapturaProduccionLiberadaDefectos 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Defectos De Tarimas Liberadas"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8985
   Icon            =   "CapturaProduccionLiberadaDefectos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   8985
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
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   8775
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   5415
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9551
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   22
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
         TabIndex        =   23
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4095
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7800
         Picture         =   "CapturaProduccionLiberadaDefectos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":237C
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   57
      ToolTipText     =   "Primer Registro"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":2CF0
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Registro Anterior"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7920
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":3664
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Siguiente Registro"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   8280
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":3FD8
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Ultimo Registro"
      Top             =   5760
      Width           =   375
   End
   Begin TabDlg.SSTab TabDefectos 
      Height          =   5415
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   9551
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   255
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CapturaProduccionLiberadaDefectos.frx":494C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameDefectos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaProduccionLiberadaDefectos.frx":4C66
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridDefectos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CapturaProduccionLiberadaDefectos.frx":50B8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPFecFin"
      Tab(2).Control(1)=   "DTPFecIni"
      Tab(2).Control(2)=   "CmdBuscar(1)"
      Tab(2).Control(3)=   "CmdBuscar(0)"
      Tab(2).Control(4)=   "TxtBuscar"
      Tab(2).Control(5)=   "FrameOpciones"
      Tab(2).Control(6)=   "Label2(6)"
      Tab(2).Control(7)=   "Label2(5)"
      Tab(2).Control(8)=   "Lbletiqueta"
      Tab(2).ControlCount=   9
      Begin MSDataGridLib.DataGrid DbGridDefectos 
         Height          =   4575
         Left            =   -74880
         TabIndex        =   53
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   720
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   10
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
            DataField       =   "Fec_PrdL"
            Caption         =   "Fecha Liberada"
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
            DataField       =   "LineaL"
            Caption         =   "Linea Liberada"
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
            DataField       =   "Esp_TecL"
            Caption         =   "Ficha Liberada"
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
            DataField       =   "TarimaL"
            Caption         =   "Tarima Liberada"
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
         BeginProperty Column09 
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
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1260.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1005.165
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPFecFin 
         Height          =   255
         Left            =   -68160
         TabIndex        =   39
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51314691
         CurrentDate     =   37522
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   255
         Left            =   -70200
         TabIndex        =   38
         Top             =   2520
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51314691
         CurrentDate     =   37522
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "CapturaProduccionLiberadaDefectos.frx":550A
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "CapturaProduccionLiberadaDefectos.frx":5814
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3480
         Width           =   2055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   17
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   3120
         Width           =   2085
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
         Height          =   1815
         Left            =   -74880
         TabIndex        =   29
         Top             =   960
         Width           =   3045
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Linea Revisada"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   52
            ToolTipText     =   " "
            Top             =   1440
            Width           =   2295
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Revisada"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Liberada"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Linea Liberada"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   " "
            Top             =   720
            Width           =   2175
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
         Height          =   4455
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   8475
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   3
            ToolTipText     =   " "
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   2
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   2
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   1
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   4
            ToolTipText     =   " "
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   9
            Left            =   1320
            TabIndex        =   9
            ToolTipText     =   " "
            Top             =   4080
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   8
            Left            =   1320
            MaxLength       =   4
            TabIndex        =   8
            ToolTipText     =   " "
            Top             =   3720
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   7
            Left            =   1320
            TabIndex        =   7
            ToolTipText     =   " "
            Top             =   3240
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   5
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   5
            ToolTipText     =   " "
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   6
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   6
            ToolTipText     =   " "
            Top             =   2880
            Width           =   1575
         End
         Begin VB.Label LblFichaTecnica 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   50
            Top             =   1080
            Width           =   5175
         End
         Begin VB.Label LblLinea 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   49
            Top             =   720
            Width           =   5175
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Tarima"
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   48
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Ficha Tecnica"
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   47
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Linea"
            Height          =   255
            Index           =   10
            Left            =   240
            TabIndex        =   46
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
            BackColor       =   &H008080FF&
            Caption         =   "Fecha "
            Height          =   255
            Index           =   9
            Left            =   240
            TabIndex        =   45
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   44
            Top             =   2880
            Width           =   1020
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
            Left            =   3000
            TabIndex        =   43
            Top             =   4080
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
            Left            =   3480
            TabIndex        =   42
            Top             =   4080
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
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
            Left            =   240
            TabIndex        =   37
            Top             =   4080
            Width           =   765
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
            Left            =   3000
            TabIndex        =   36
            Top             =   3720
            Width           =   5175
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Defecto"
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
            Left            =   240
            TabIndex        =   35
            Top             =   3720
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H0080C0FF&
            Caption         =   "Tarima"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   34
            Top             =   3240
            Width           =   480
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Linea"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   33
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label LblLinea2 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   32
            Top             =   2520
            Width           =   5175
         End
         Begin VB.Label LblFichaTecnica2 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
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
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   3000
            TabIndex        =   31
            Top             =   2880
            Width           =   5175
         End
         Begin VB.Label Label2 
            BackColor       =   &H0080C0FF&
            Caption         =   "Fecha "
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   28
            Top             =   2160
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H008080FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   6  'Inside Solid
            Height          =   1575
            Index           =   0
            Left            =   120
            Top             =   240
            Width           =   8175
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H0080C0FF&
            BackStyle       =   1  'Opaque
            BorderStyle     =   6  'Inside Solid
            Height          =   1575
            Index           =   1
            Left            =   120
            Top             =   2040
            Width           =   8175
         End
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
         Left            =   -70800
         TabIndex        =   41
         Top             =   2520
         Width           =   555
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
         Left            =   -68760
         TabIndex        =   40
         Top             =   2520
         Width           =   510
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
         Left            =   -70800
         TabIndex        =   30
         Top             =   3120
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6720
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":5C56
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":6098
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   1065
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5280
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":65B3
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":69F5
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1300
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   3840
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":6FBD
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":73FF
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1300
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2400
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":7936
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":7D78
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   1300
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   960
      MouseIcon       =   "CapturaProduccionLiberadaDefectos.frx":82D4
      Picture         =   "CapturaProduccionLiberadaDefectos.frx":8716
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1300
   End
End
Attribute VB_Name = "CapturaProduccionLiberadaDefectos"
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
Dim BFichaTecnica2 As Boolean
Dim BLinea2 As Boolean
Dim BDefecto As Boolean

Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaFichaTecnica2 As New ADODB.Recordset
Dim RBuscaLinea2 As New ADODB.Recordset
Dim RBuscaDefectos As New ADODB.Recordset
Dim RBuscaAtributos As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RBuscaTarimaLiberada As New ADODB.Recordset
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
         DbGridDefectos.Visible = False
    Else
         FrameDefectos.Enabled = False
         CmdBotones.Item(0).Enabled = True
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
         DbGridDefectos.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next

            If Index = 0 Then
                        Bandera = True
                        botones
                        Limpia_Campos
                        
                    
                        'ASIGNA LA TARIMA NUEVAQUE SE ESTA LIBERANDO
                        TxtTexto.Item(0).Text = VPDFecha
                        TxtTexto.Item(1).Text = VPDLinea
                        TxtTexto.Item(2).Text = VPDFicha
                        TxtTexto.Item(3).Text = VPDTarima
                        
                        'ASIGNA LA TARIMA QUE SE LIBERO PERO QUE ES RECHAZADA
                        TxtTexto.Item(4).Text = VPLDFecha
                        TxtTexto.Item(5).Text = VPLDLinea
                        TxtTexto.Item(6).Text = VPLDFicha
                        TxtTexto.Item(7).Text = VPLDTarima
                        
                        TxtTexto.Item(8).SetFocus
            'GRABAR
            ElseIf Index = 2 Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                    Else
                        TxtTexto.Item(0).Text = Format(TxtTexto.Item(0).Text, "dd/mm/yyyy")
                    End If
                    
                    'FECHA TARIMA LIBERADA
                     If Not IsDate(TxtTexto.Item(0).Text) Then
                            MsgBox "Fecha De Tarima Liberada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(1).SetFocus
                            Exit Sub
                     End If
                     'NUMERO DE TARIMA LIBERADA
                     If Not IsNumeric(TxtTexto.Item(3).Text) Then
                            MsgBox "Numero De Tarima Liberada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(3).SetFocus
                            Exit Sub
                     End If
                   
                     'FECHA TARIMA REVISADA
                     If Not IsDate(TxtTexto.Item(4).Text) Then
                            MsgBox "Fecha De Tarima Revisada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(4).SetFocus
                            Exit Sub
                     End If
                     'FECHA TARIMA REVISADA
                     If Not IsNumeric(TxtTexto.Item(7).Text) Then
                            MsgBox "Numero De Tarima Revisada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(7).SetFocus
                            Exit Sub
                     End If
                     
                     'CANTIDAD DE DEFECTOS
                     If Not IsNumeric(TxtTexto.Item(9).Text) Then
                            MsgBox "Cantidad De Defectos Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(9).SetFocus
                            Exit Sub
                     End If
                     
                                 
            
                    'BUSCA SI EXISTE LA TARIMA LIBERADA
                     Set RBuscaTarima = New ADODB.Recordset
                         If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaTarima, "Select * from produccionLiberada Where Linea = '" & TxtTexto.Item(1) & "' and Esp_tec = '" & TxtTexto.Item(2).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(0).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(3).Text)
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaTarima, "Select * from produccionLiberada Where Linea = '" & TxtTexto.Item(1) & "' and Esp_tec = '" & TxtTexto.Item(2).Text & "' and Fec_prd = To_Date('" & TxtTexto.Item(0).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(3).Text)
                        End If
                         If RBuscaTarima.RecordCount > 0 Then
                         Else
                            MsgBox "La Tarima Liberada No Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                        End If
                   
                     'BUSCA SI EXISTE LA TARIMA LIBERADA
                     Set RBuscaTarimaLiberada = New ADODB.Recordset
                         If GOrigenDeDatos = "AmaproAccess" Then
                              Call Abrir_Recordset(RBuscaTarimaLiberada, "Select * from produccion Where Linea = '" & TxtTexto.Item(5) & "' and Esp_tec = '" & TxtTexto.Item(6).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(4).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(7).Text)
                         Else
                              Call Abrir_Recordset(RBuscaTarimaLiberada, "Select * from produccion Where Linea = '" & TxtTexto.Item(5) & "' and Esp_tec = '" & TxtTexto.Item(6).Text & "' and Fec_prd = To_Date('" & TxtTexto.Item(4).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(7).Text)
                         End If
                         If RBuscaTarimaLiberada.RecordCount > 0 Then
                         Else
                            MsgBox "La Tarima Revisada No Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(4).SetFocus
                            Exit Sub
                        End If
                     
                   
                    'AGREGAR
                    'If BEditar = False Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "Values(#" & Format(TxtTexto.Item(0).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = "Values(To_Date('" & Format(TxtTexto.Item(0).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy'), " 'FECHA
                            End If
                            
                            VTexto = VTexto & "'" & TxtTexto.Item(1).Text & "', '" 'LINEA
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(3).Text & ", " 'TARIMA
                            
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "#" & Format(TxtTexto.Item(4).Text, "mm/dd/yyyy") & "#, "  'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "To_Date('" & TxtTexto.Item(4).Text & "', 'dd/mm/yyyy'), "  'FECHA
                            End If
                            
                            VTexto = VTexto & "'" & TxtTexto.Item(5).Text & "', '" 'LINEA
                            VTexto = VTexto & TxtTexto.Item(6).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(7).Text & ", '" 'TARIMA
                            
                            VTexto = VTexto & TxtTexto.Item(8).Text & "', " 'DEFECTO
                            VTexto = VTexto & TxtTexto.Item(9).Text & ")" 'CANTIDAD
                            
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into ProduccionLiberadaConDefectos " & VTexto
                    'EDITAR
                    'Else
                            'VTexto = "Cantidad = " & Txttexto.Item(5).Text & " " 'CANTIDAD
                            
                            'If GOrigenDeDatos = "AmaproAccess" Then
                            '    VTexto = VTexto & "Where Fec_Prd = #" & Format(Txttexto.Item(0).Text, "mm/dd/yyyy") & "# And "
                            'Else 'ORACLE
                            '    VTexto = VTexto & "Where Fec_Prd = TO_DATE('" & Txttexto.Item(0).Text & "', 'dd/mm/yyyy') And "
                            'End If
                            '    VTexto = VTexto & "Linea = '" & Txttexto.Item(1) & "' And Esp_Tec = '" & Txttexto.Item(2) & "' And "
                            '    VTexto = VTexto & "Tarima = " & Txttexto.Item(3) & " And Defecto = '" & Txttexto.Item(4) & "'"
                        
                            'Conexion.Execute "UPDATE ProduccionLiberadaConDefectos SET " & VTexto
                    'End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Fecha, Linea, FichaTecnica, Tarima y Defecto Ya Existe Para Esta Tarima", vbOKOnly + vbInformation, "Informacion"
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
                    
                    
            'BORRAR
            ElseIf Index = 4 Then
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
                    
            'SALIDA
            ElseIf Index = 5 Then
                    Unload Me
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

Private Sub CmdBuscar_Click(Index As Integer)
    Set RDefectos = New ADODB.Recordset
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptOpcion.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDefectos, "Select * from ProduccionLiberadaConDefectos where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order by Fec_prd")
                Else 'ORACLE
                    Call Abrir_Recordset(RDefectos, "Select * from ProduccionLiberadaConDefectos where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " Order by Fec_prd")
                End If
            ElseIf OptOpcion.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDefectos, "Select * from ProduccionLiberadaConDefectos where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBuscar.Text & "' Order by Fec_prd")
                Else 'ORACLE
                    Call Abrir_Recordset(RDefectos, "Select * from ProduccionLiberadaConDefectos where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy') And UPPER(Linea) = '" & UCase(TxtBuscar.Text) & "' Order by Fec_prd")
                End If
            ElseIf OptOpcion.Item(2).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDefectos, "Select * from ProduccionLiberadaConDefectos where Fec_PrdL >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order by Fec_prd")
                Else 'ORACLE
                    Call Abrir_Recordset(RDefectos, "Select * from ProduccionLiberadaConDefectos where Fec_PrdL >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " Order by Fec_prd")
                End If
            ElseIf OptOpcion.Item(3).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDefectos, "Select * from ProduccionLiberadaConDefectos where Fec_PrdL >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And LineaL = '" & TxtBuscar.Text & "' Order by Fec_prd")
                Else 'ORACLE
                    Call Abrir_Recordset(RDefectos, "Select * from ProduccionLiberadaConDefectos where Fec_PrdL >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy') And UPPER(LineaL) = '" & UCase(TxtBuscar.Text) & "' Order by Fec_prd")
                End If
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RDefectos, "Select * From ProduccionLiberadaConDefectos")
        End If
                
                Set DbGridDefectos.DataSource = RDefectos
        
        
    
        TabDefectos.Tab = 1
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub


Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub


Private Sub Dbgriddefectos_HeadClick(ByVal ColIndex As Integer)
            RDefectos.Sort = RDefectos.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_DblClick()
            If BLinea = True Then
                TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(1).SetFocus
            ElseIf BFichaTecnica = True Then
                TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(2).SetFocus
            ElseIf BLinea2 = True Then
                TxtTexto.Item(5).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(5).SetFocus
            ElseIf BFichaTecnica2 = True Then
                TxtTexto.Item(6).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(6).SetFocus
            ElseIf BDefecto = True Then
                TxtTexto.Item(8).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(8).SetFocus
            End If
                FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                    If BLinea = True Then
                        TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0).Text
                        TxtTexto.Item(1).SetFocus
                    ElseIf BFichaTecnica = True Then
                        TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
                        TxtTexto.Item(2).SetFocus
                    ElseIf BLinea2 = True Then
                        TxtTexto.Item(5).Text = DBGridBusqueda.Columns(0).Text
                        TxtTexto.Item(5).SetFocus
                    ElseIf BFichaTecnica2 = True Then
                        TxtTexto.Item(6).Text = DBGridBusqueda.Columns(0).Text
                        TxtTexto.Item(6).SetFocus
                    ElseIf BDefecto = True Then
                        TxtTexto.Item(8).Text = DBGridBusqueda.Columns(0).Text
                        TxtTexto.Item(8).SetFocus
                    End If
                        FrameBusqueda.Visible = False
            End If
                    
End Sub






Private Sub Form_Load()
        Set RDefectos = New ADODB.Recordset
        
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RDefectos, "Select * From ProduccionLiberadaConDefectos Where Fec_Prd = #" & Format(Date, "mm/dd/yyyy") & "#")
            Else 'ORACLE
                Call Abrir_Recordset(RDefectos, "Select * From ProduccionLiberadaConDefectos Where Fec_Prd = To_Date('" & Date & "', 'dd/mm/yyyy')")
            End If
        
            DtpFecIni.Value = Date
            DTPFecFin.Value = Date
            
            Set DbGridDefectos.DataSource = RDefectos
            Llena_Campos
End Sub



Private Sub OptOpcion_Click(Index As Integer)
        If (Index = 1 Or Index = 3) Then
            Lbletiqueta.Caption = "Linea"
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
        Else
            Lbletiqueta.Caption = ""
            TxtBuscar.Visible = False
        End If
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
    If (BLinea = True Or BLinea2 = True) Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where Descrip Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
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
    ElseIf (BFichaTecnica = True Or BFichaTecnica2 = True) Then
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
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%' Order By Esp_Tec")
                    End If
            End If
            
    'DEFECTO
    ElseIf BDefecto = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos where Descrip Like '%" & TxtBusqueda.Text & "%' Order By Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%' Order By Descrip")
                    End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos Where Defecto Like '%" & TxtBusqueda.Text & "%' Order By Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos Where UPPER(Defecto) Like '%" & UCase(TxtBusqueda.Text) & "%' Order By Descrip")
                    End If
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
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtTexto_Change(Index As Integer)
        'BUSCA LA DESCRIPCION DE LINEA
        If Index = 1 Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(1).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(1).Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        'BUSCA LA DESCRIPCION DE FICHA TECNICA
        ElseIf Index = 2 Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where Esp_Tec = '" & TxtTexto.Item(2).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(2).Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblFichaTecnica.Caption = ""
                End If
        'BUSCA LA DESCRIPCION DE LINEA
        ElseIf Index = 5 Then
            Set RBuscaLinea2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea2, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(5).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaLinea2, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(5).Text) & "'")
                End If
                If RBuscaLinea2.RecordCount > 0 Then
                    LblLinea2.Caption = RBuscaLinea2!Descrip
                Else
                    LblLinea2.Caption = ""
                End If
                
        'BUSCA LA DESCRIPCION DE FICHA TECNICA
        ElseIf Index = 6 Then
            Set RBuscaFichaTecnica2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica2, "Select Descrip From FichaTecnica where Esp_Tec = '" & TxtTexto.Item(6).Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica2, "Select Descrip From FichaTecnica where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(6).Text) & "'")
                End If
                If RBuscaFichaTecnica2.RecordCount > 0 Then
                    LblFichaTecnica2.Caption = RBuscaFichaTecnica2!Descrip
                Else
                    LblFichaTecnica2.Caption = ""
                End If
        
        'BUSCA LA DESCRIPCION DE DEFECTOS
        ElseIf Index = 8 Then
            Set RBuscaDefectos = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaDefectos, "Select Descrip, Tipo From Defectos Where Defecto = '" & TxtTexto.Item(8).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaDefectos, "Select Descrip, Tipo From Defectos Where UPPER(Defecto) = '" & UCase(TxtTexto.Item(8).Text) & "'")
                End If
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
                If (Index = 1 Or Index = 2 Or Index = 5 Or Index = 6 Or Index = 8) Then
                    Set RBusqueda = New ADODB.Recordset
                End If
                'LINEAS
                If Index = 1 Then
                    BFichaTecnica = False
                    BLinea = True
                    BFichaTecnica2 = False
                    BLinea2 = False
                    BDefecto = False
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
                'SI ELIGE FICHA TECNICA
                ElseIf Index = 2 Then
                    BFichaTecnica = True
                    BLinea = False
                    BFichaTecnica2 = False
                    BLinea2 = False
                    BDefecto = False
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                'LINEAS
                ElseIf Index = 5 Then
                    BFichaTecnica = False
                    BLinea = False
                    BFichaTecnica2 = False
                    BLinea2 = True
                    BDefecto = False
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
                'SI ELIGE FICHA TECNICA
                ElseIf Index = 6 Then
                    BFichaTecnica = False
                    BLinea = False
                    BFichaTecnica2 = True
                    BLinea2 = False
                    BDefecto = False
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                'DEFECTOS
                ElseIf Index = 8 Then
                    BFichaTecnica = False
                    BLinea = False
                    BFichaTecnica2 = False
                    BLinea2 = False
                    BDefecto = True
                    Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos")
                End If
                
                If (Index = 1 Or Index = 2 Or Index = 5 Or Index = 6 Or Index = 8) Then
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
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
                        If (Index = 1 Or Index = 2 Or Index = 5 Or Index = 6 Or Index = 8) Then
                            Set RBusqueda = New ADODB.Recordset
                        End If
                        'LINEAS
                        If Index = 1 Then
                            BFichaTecnica = False
                            BLinea = True
                            BFichaTecnica2 = False
                            BLinea2 = False
                            BDefecto = False
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
                        'SI ELIGE FICHA TECNICA
                        ElseIf Index = 2 Then
                            BFichaTecnica = True
                            BLinea = False
                            BFichaTecnica2 = False
                            BLinea2 = False
                            BDefecto = False
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                        'LINEAS
                        ElseIf Index = 5 Then
                            BFichaTecnica = False
                            BLinea = False
                            BFichaTecnica2 = False
                            BLinea2 = True
                            BDefecto = False
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
                        'SI ELIGE FICHA TECNICA
                        ElseIf Index = 6 Then
                            BFichaTecnica = False
                            BLinea = False
                            BFichaTecnica2 = True
                            BLinea2 = False
                            BDefecto = False
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                        'DEFECTOS
                        ElseIf Index = 8 Then
                            BFichaTecnica = False
                            BLinea = False
                            BFichaTecnica2 = False
                            BLinea2 = False
                            BDefecto = True
                            Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip From Defectos")
                        End If
                        
                        If (Index = 1 Or Index = 2 Or Index = 5 Or Index = 6 Or Index = 8) Then
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"
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
        'FECHA
            If IsNull(RDefectos!fec_prdL) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = RDefectos!fec_prdL
            End If
        'LINEA
            If IsNull(RDefectos!LineaL) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = RDefectos!LineaL
            End If
        'FICHA
            If IsNull(RDefectos!Esp_TecL) Then
                TxtTexto.Item(6).Text = ""
            Else
                TxtTexto.Item(6).Text = RDefectos!Esp_TecL
            End If
        'TARIMA
            If IsNull(RDefectos!TarimaL) Then
                TxtTexto.Item(7).Text = ""
            Else
                TxtTexto.Item(7).Text = RDefectos!TarimaL
            End If
        
        'DEFECTO
            If IsNull(RDefectos!Defecto) Then
                TxtTexto.Item(8).Text = ""
            Else
                TxtTexto.Item(8).Text = RDefectos!Defecto
            End If
        'CANTIDAD
            If IsNull(RDefectos!Cantidad) Then
                TxtTexto.Item(9).Text = ""
            Else
                TxtTexto.Item(9).Text = RDefectos!Cantidad
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
        TxtTexto.Item(6).Text = ""
        TxtTexto.Item(7).Text = ""
        TxtTexto.Item(8).Text = ""
        TxtTexto.Item(9).Text = ""
        
        
End Sub


