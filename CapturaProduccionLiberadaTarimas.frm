VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CapturaProduccionLiberadaTarimas 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Tarimas Revisadas Para Produccion Liberada"
   ClientHeight    =   7965
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8685
   Icon            =   "CapturaProduccionLiberadaTarimas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7965
   ScaleWidth      =   8685
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
      Height          =   7815
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6495
         Left            =   120
         TabIndex        =   31
         Top             =   1080
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
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   28
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
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   30
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   7440
         Picture         =   "CapturaProduccionLiberadaTarimas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   8040
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":237C
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   67
      ToolTipText     =   "Ultimo Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7680
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":2CF0
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Siguiente Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   600
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":3664
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Registro Anterior"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   240
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":3FD8
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Primer Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Defectos Tarima"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Index           =   6
      Left            =   5400
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":494C
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":4D8E
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7080
      Width           =   1000
   End
   Begin TabDlg.SSTab TabDefectos 
      Height          =   6855
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12091
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   255
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CapturaProduccionLiberadaTarimas.frx":52C0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameDefectos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaProduccionLiberadaTarimas.frx":55DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CapturaProduccionLiberadaTarimas.frx":5A2C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "Label2(5)"
      Tab(2).Control(2)=   "Label2(6)"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "TxtBuscar"
      Tab(2).Control(5)=   "CmdBuscar(0)"
      Tab(2).Control(6)=   "CmdBuscar(1)"
      Tab(2).Control(7)=   "DTPFecIni"
      Tab(2).Control(8)=   "DTPFecFin"
      Tab(2).ControlCount=   9
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   63
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   720
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   10610
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
         ColumnCount     =   15
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
         BeginProperty Column08 
            DataField       =   "CalidadL"
            Caption         =   "Calidad"
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
            DataField       =   "Revisados"
            Caption         =   "Revisados"
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
         BeginProperty Column10 
            DataField       =   "NoConforme"
            Caption         =   "No Conforme"
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
         BeginProperty Column11 
            DataField       =   "Liberados"
            Caption         =   "Liberados"
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
         BeginProperty Column12 
            DataField       =   "EnTarima"
            Caption         =   "En Tarima"
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
         BeginProperty Column13 
            DataField       =   "Minutos"
            Caption         =   "Minutos"
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
         BeginProperty Column14 
            DataField       =   "Empleado"
            Caption         =   "Empleado"
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
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   585.071
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   329.953
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1035.213
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column14 
               ColumnWidth     =   945.071
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPFecFin 
         Height          =   255
         Left            =   -68160
         TabIndex        =   45
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51838979
         CurrentDate     =   37522
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   255
         Left            =   -70200
         TabIndex        =   44
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51838979
         CurrentDate     =   37522
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "CapturaProduccionLiberadaTarimas.frx":5E7E
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5280
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "CapturaProduccionLiberadaTarimas.frx":6188
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   4320
         Width           =   2055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   23
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   3840
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
         TabIndex        =   36
         Top             =   960
         Width           =   2925
         Begin VB.OptionButton OptOpciones 
            Caption         =   "Fechas Y Linea Revisado"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   61
            Top             =   1440
            Width           =   2295
         End
         Begin VB.OptionButton OptOpciones 
            Caption         =   "Fechas Revisado"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   60
            Top             =   1080
            Width           =   1695
         End
         Begin VB.OptionButton OptOpciones 
            Caption         =   "Fechas Liberado"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptOpciones 
            Caption         =   "Fechas Y Linea Liberado"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   " "
            Top             =   720
            Width           =   2175
         End
      End
      Begin VB.Frame FrameDefectos 
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
         Height          =   6015
         Left            =   120
         TabIndex        =   33
         Top             =   720
         Width           =   8235
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   14
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   12
            ToolTipText     =   " "
            Top             =   5520
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   13
            Left            =   6480
            MaxLength       =   50
            TabIndex        =   14
            ToolTipText     =   " "
            Top             =   5520
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   12
            Left            =   6480
            TabIndex        =   13
            ToolTipText     =   " "
            Top             =   5160
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Index           =   11
            Left            =   1440
            TabIndex        =   11
            ToolTipText     =   " "
            Top             =   5160
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
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
            ForeColor       =   &H000000FF&
            Height          =   285
            Index           =   10
            Left            =   1440
            TabIndex        =   10
            ToolTipText     =   " "
            Top             =   4680
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
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
            Index           =   9
            Left            =   1440
            TabIndex        =   9
            ToolTipText     =   " "
            Top             =   4320
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   8
            Left            =   1440
            MaxLength       =   1
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
            Left            =   1440
            TabIndex        =   7
            ToolTipText     =   " "
            Top             =   3360
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   6
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   6
            ToolTipText     =   " "
            Top             =   3000
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   0
            Left            =   1440
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   5
            Left            =   1440
            MaxLength       =   2
            TabIndex        =   5
            ToolTipText     =   " "
            Top             =   2640
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   4
            Left            =   1440
            TabIndex        =   4
            ToolTipText     =   " "
            Top             =   2280
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   3
            Left            =   1440
            TabIndex        =   3
            ToolTipText     =   " "
            Top             =   1440
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   1
            Left            =   1440
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
            Index           =   2
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   2
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "En Tarima"
            Height          =   195
            Index           =   17
            Left            =   240
            TabIndex        =   62
            Top             =   5520
            Width           =   720
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   3000
            Picture         =   "CapturaProduccionLiberadaTarimas.frx":65CA
            Top             =   4560
            Width           =   480
         End
         Begin VB.Line Line1 
            X1              =   1320
            X2              =   3240
            Y1              =   5040
            Y2              =   5040
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tarima Liberada"
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
            Index           =   16
            Left            =   6480
            TabIndex        =   59
            Top             =   120
            Width           =   1380
         End
         Begin VB.Shape Shape2 
            Height          =   1575
            Left            =   120
            Top             =   240
            Width           =   8055
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   " Tarima Revisada "
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
            Index           =   15
            Left            =   6480
            TabIndex        =   58
            Top             =   2040
            Width           =   1560
         End
         Begin VB.Shape Shape1 
            Height          =   1935
            Left            =   120
            Top             =   2160
            Width           =   8055
         End
         Begin VB.Label LblFichaTecnica2 
            Appearance      =   0  'Flat
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
            Left            =   3120
            TabIndex        =   57
            Top             =   3000
            Width           =   4935
         End
         Begin VB.Label LblLinea2 
            Appearance      =   0  'Flat
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
            Left            =   3120
            TabIndex        =   56
            Top             =   2640
            Width           =   4935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Empleado Que Reviso"
            Height          =   195
            Index           =   14
            Left            =   4080
            TabIndex        =   55
            Top             =   5520
            Width           =   1590
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Minutos En Revision"
            Height          =   195
            Index           =   13
            Left            =   4080
            TabIndex        =   54
            Top             =   5160
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Liberados"
            Height          =   195
            Index           =   12
            Left            =   240
            TabIndex        =   53
            Top             =   5160
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No Conforme"
            Height          =   195
            Index           =   11
            Left            =   240
            TabIndex        =   52
            Top             =   4680
            Width           =   930
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Revisados"
            Height          =   195
            Index           =   10
            Left            =   240
            TabIndex        =   51
            Top             =   4320
            Width           =   750
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Calidad"
            Height          =   195
            Index           =   9
            Left            =   240
            TabIndex        =   50
            Top             =   3720
            Width           =   525
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Height          =   195
            Index           =   8
            Left            =   240
            TabIndex        =   49
            Top             =   3000
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tarima"
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
            Left            =   240
            TabIndex        =   48
            Top             =   3360
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Index           =   4
            Left            =   240
            TabIndex        =   43
            Top             =   2640
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            TabIndex        =   42
            Top             =   2280
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tarima"
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
            Left            =   240
            TabIndex        =   41
            Top             =   1440
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
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
            Left            =   240
            TabIndex        =   40
            Top             =   720
            Width           =   480
         End
         Begin VB.Label LblLinea 
            Appearance      =   0  'Flat
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
            Left            =   3120
            TabIndex        =   39
            Top             =   720
            Width           =   4935
         End
         Begin VB.Label LblFichaTecnica 
            Appearance      =   0  'Flat
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
            Left            =   3120
            TabIndex        =   38
            Top             =   1080
            Width           =   4935
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
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
            Height          =   195
            Left            =   240
            TabIndex        =   35
            Top             =   1080
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
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
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   600
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
         TabIndex        =   47
         Top             =   2640
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
         TabIndex        =   46
         Top             =   2640
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
         Left            =   -70920
         TabIndex        =   37
         Top             =   3840
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6480
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":6A0C
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":6E4E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7080
      Width           =   1125
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   4320
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":7369
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":77AB
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   3240
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":7D73
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":81B5
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2160
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":86EC
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":8B2E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   1080
      MouseIcon       =   "CapturaProduccionLiberadaTarimas.frx":908A
      Picture         =   "CapturaProduccionLiberadaTarimas.frx":94CC
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7080
      Width           =   1000
   End
End
Attribute VB_Name = "CapturaProduccionLiberadaTarimas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim VCantidad As Long

Dim BFichaTecnica As Boolean
Dim BLinea As Boolean
Dim BFichaTecnicaL As Boolean
Dim BLineaL As Boolean
Dim VTexto As String
Dim VMensaje As String

Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RBuscaTarima2 As New ADODB.Recordset
Dim RBuscaTarimaLiberada As New ADODB.Recordset
Dim RBuscaTarimaLiberada2 As New ADODB.Recordset
Dim RTarimas As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset


Sub botones()
    If Bandera = True Then
         FrameDefectos.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         CmdBotones.Item(6).Enabled = False
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
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         CmdBotones.Item(6).Enabled = True
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
        
            If Index = 0 Then
                    Bandera = True
                    botones
                    Limpia_Campos
                    
                    
                    TxtTexto.Item(0).Text = VPLFecha
                    TxtTexto.Item(1).Text = VPLLinea
                    TxtTexto.Item(2).Text = VPLFicha
                    TxtTexto.Item(3).Text = VPLTarima
                    
                    TxtTexto.Item(6).Text = VPLFicha
                    
                    TxtTexto.Item(4).SetFocus
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
                   
                   
                     If GOrigenDeDatos = "AmaproAccess" Then
                     Else
                         TxtTexto.Item(4).Text = Format(TxtTexto.Item(4).Text, "dd/mm/yyyy")
                     End If
                     'FECHA TARIMA REVISADA
                     If Not IsDate(TxtTexto.Item(4).Text) Then
                            MsgBox "Fecha De Tarima Revisada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(4).SetFocus
                            Exit Sub
                     End If
                     
                     If Not IsNumeric(TxtTexto.Item(7).Text) Then
                            MsgBox "Numero De Tarima Revisada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(7).SetFocus
                            Exit Sub
                     End If
                     
                     'REVISADOS
                     If Not IsNumeric(TxtTexto.Item(9).Text) Then
                            MsgBox "Cantidad Revisada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(9).SetFocus
                            Exit Sub
                     End If
                   
                     'NO CONFORME
                     If Not IsNumeric(TxtTexto.Item(10).Text) Then
                            MsgBox "Cantidad No Conforme Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(10).SetFocus
                            Exit Sub
                     End If
                     
                     'LIBERADOS
                     If Not IsNumeric(TxtTexto.Item(11).Text) Then
                            MsgBox "Cantidad Liberada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(11).SetFocus
                            Exit Sub
                     End If
                     
                     'EN TARIMA
                     If Not IsNumeric(TxtTexto.Item(14).Text) Then
                            MsgBox "Cantidad En Tarima Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(14).SetFocus
                            Exit Sub
                     End If
                   
                     'MINUTOS
                     If Not IsNumeric(TxtTexto.Item(12).Text) Then
                            MsgBox "Cantidad Minutos Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(12).SetFocus
                            Exit Sub
                     End If
                     
                     'BUSCA SI EXISTE LA TARIMA LIBERADA
                     Set RBuscaTarima = New ADODB.Recordset
                         If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaTarima, "Select * from produccionLiberada Where Linea = '" & TxtTexto.Item(1) & "' and Esp_tec = '" & TxtTexto.Item(2).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(0).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(3).Text)
                         Else 'ORACLE
                            Call Abrir_Recordset(RBuscaTarima, "Select * from produccionLiberada Where UPPER(Linea) = '" & UCase(TxtTexto.Item(1)) & "' and UPPER(Esp_tec) = '" & UCase(TxtTexto.Item(2).Text) & "' and Fec_prd = TO_Date('" & TxtTexto.Item(0).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(3).Text)
                         End If
                         If RBuscaTarima.RecordCount > 0 Then
                         Else
                            MsgBox "La Tarima Liberada No Existe, En Produccion Liberada", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                        End If
                   
                     'BUSCA SI EXISTE LA TARIMA LIBERADA EN PRODUCCION
                     Set RBuscaTarimaLiberada = New ADODB.Recordset
                         If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaTarimaLiberada, "Select * from produccion Where Linea = '" & TxtTexto.Item(5) & "' and Esp_tec = '" & TxtTexto.Item(6).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(4).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(7).Text)
                         Else 'ORACLE
                            Call Abrir_Recordset(RBuscaTarimaLiberada, "Select * from produccion Where UPPER(Linea) = '" & UCase(TxtTexto.Item(5)) & "' and UPPER(Esp_tec) = '" & UCase(TxtTexto.Item(6).Text) & "' and Fec_prd = To_Date('" & TxtTexto.Item(4).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(7).Text)
                         End If
                     'BUSCA SI EXISTE LA TARIMA LIBERADA EN PRODUCCION
                     Set RBuscaTarimaLiberada2 = New ADODB.Recordset
                         If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaTarimaLiberada2, "Select * from produccionLiberada Where Linea = '" & TxtTexto.Item(5) & "' and Esp_tec = '" & TxtTexto.Item(6).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(4).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(7).Text)
                         Else 'ORACLE
                                Call Abrir_Recordset(RBuscaTarimaLiberada2, "Select * from produccionLiberada Where UPPER(Linea) = '" & UCase(TxtTexto.Item(5)) & "' and UPPER(Esp_tec) = '" & UCase(TxtTexto.Item(6).Text) & "' and Fec_prd = To_Date('" & TxtTexto.Item(4).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(7).Text)
                         End If
                                
                     
                         If (RBuscaTarimaLiberada.RecordCount > 0 Or RBuscaTarimaLiberada2.RecordCount > 0) Then
                         Else
                            MsgBox "La Tarima Revisada No Existe, Ni En Produccion Ni En Produccion Liberada", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(4).SetFocus
                            Exit Sub
                        End If
                        
                         
                        'AGREGAR
                    'If BEditar = False Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "Values(#" & Format(TxtTexto.Item(0).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = "Values(To_Date('" & Format(TxtTexto.Item(0).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            
                            VTexto = VTexto & "'" & TxtTexto.Item(1).Text & "', '" 'LINEA
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(3).Text & ", " 'TARIMA
                            
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "#" & Format(TxtTexto.Item(4).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "To_Date('" & Format(TxtTexto.Item(4).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            
                            VTexto = VTexto & "'" & TxtTexto.Item(5).Text & "', '" 'LINEA
                            VTexto = VTexto & TxtTexto.Item(6).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(7).Text & ", '" 'TARIMA
                            
                            VTexto = VTexto & TxtTexto.Item(8).Text & "', " 'CALIDAD
                            VTexto = VTexto & TxtTexto.Item(9).Text & "," 'REVISADOS
                            VTexto = VTexto & TxtTexto.Item(10).Text & "," 'NO CONFORME
                            VTexto = VTexto & TxtTexto.Item(11).Text & "," 'LIBERADOS
                            VTexto = VTexto & TxtTexto.Item(14).Text & "," 'EN TARIMA
                            VTexto = VTexto & TxtTexto.Item(12).Text & ", '" 'MINUTOS
                            VTexto = VTexto & TxtTexto.Item(13).Text & "')" 'EMPLEADO
                            
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into ProduccionLiberadaConTarimas " & VTexto
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
                                                
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RTarimas.Requery
                        RTarimas.MoveLast
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
                        RTarimas.Delete
                        
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
                        RTarimas.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RTarimas.MoveNext
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
            'SALIDA
            ElseIf Index = 6 Then
                    VPDFecha = TxtTexto.Item(0).Text
                    VPDLinea = TxtTexto.Item(1).Text
                    VPDFicha = TxtTexto.Item(2).Text
                    VPDTarima = TxtTexto.Item(3).Text
                                        
                    VPLDFecha = TxtTexto.Item(4).Text
                    VPLDLinea = TxtTexto.Item(5).Text
                    VPLDFicha = TxtTexto.Item(6).Text
                    VPLDTarima = TxtTexto.Item(7).Text
                    
                    CapturaProduccionLiberadaDefectos.Show 1
            End If
        
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RTarimas.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RTarimas.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RTarimas.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RTarimas.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RTarimas.BOF Then
        RTarimas.MoveFirst
    ElseIf RTarimas.EOF Then
        RTarimas.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0


End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        Set RTarimas = New ADODB.Recordset
        
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptOpciones.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTarimas, "Select * from ProduccionLiberadaConTarimas where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order by Fec_prd")
                Else 'ORACLE
                    Call Abrir_Recordset(RTarimas, "Select * from ProduccionLiberadaConTarimas where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " Order by Fec_prd")
                End If
            ElseIf OptOpciones.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTarimas, "Select * from ProduccionLiberadaConTarimas where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBuscar.Text & "' Order by Fec_prd")
                Else 'ORACLE
                    Call Abrir_Recordset(RTarimas, "Select * from ProduccionLiberadaConTarimas where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And Linea = '" & TxtBuscar.Text & "' Order by Fec_prd")
                End If
            ElseIf OptOpciones.Item(2).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTarimas, "Select * from ProduccionLiberadaConTarimas where Fec_PrdL >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_PrdL <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order by Fec_prd")
                Else 'ORACLE
                    Call Abrir_Recordset(RTarimas, "Select * from ProduccionLiberadaConTarimas where Fec_PrdL >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_PrdL <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " Order by Fec_prd")
                End If
            ElseIf OptOpciones.Item(3).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTarimas, "Select * from ProduccionLiberadaConTarimas where Fec_PrdL >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_PrdL <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And LineaL = '" & TxtBuscar.Text & "' Order by Fec_prd")
                Else 'ORACLE
                    Call Abrir_Recordset(RTarimas, "Select * from ProduccionLiberadaConTarimas where Fec_PrdL >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_PrdL <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And LineaL = '" & TxtBuscar.Text & "' Order by Fec_prd")
                End If
            End If
                
                Set DataGrid1.DataSource = RTarimas
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RTarimas, "Select * From ProduccionLiberadaConTarimas")
                Set DataGrid1.DataSource = RTarimas
        End If
    
        TabDefectos.Tab = 1
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub



Private Sub DBGridBusqueda_DblClick()
            If BFichaTecnica = True Then
                TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(0).SetFocus
            ElseIf BLinea = True Then
                TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(2).SetFocus
            End If
                FrameBusqueda.Visible = False
End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                If BFichaTecnica = True Then
                    TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
                    TxtTexto.Item(0).SetFocus
                ElseIf BLinea = True Then
                    TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
                    TxtTexto.Item(2).SetFocus
                End If
                    FrameBusqueda.Visible = False
            End If
                    
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
        RTarimas.Sort = RTarimas.Fields(ColIndex).Name
End Sub


Private Sub Form_Load()
On Error Resume Next
        Set RTarimas = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RTarimas, "Select * From ProduccionLiberadaConTarimas Where Fec_Prd = #" & Format(Date, "mm/dd/yyyy") & "#")
            Else 'ORACLE
                Call Abrir_Recordset(RTarimas, "Select * From ProduccionLiberadaConTarimas Where Fec_Prd = To_Date('" & Date & "', 'dd/mm/yyyy')")
            End If
        
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date
        
        Set DataGrid1.DataSource = RTarimas
        RTarimas.MoveLast
        If Err <> 0 Then
        End If
        Llena_Campos

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
            Set RTarimas = New ADODB.Recordset
    'LINEA
    If (BLinea = True Or BLineaL = True) Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTarimas, "Select Linea, Descrip From Lineas where Descrip Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RTarimas, "Select Linea, Descrip From Lineas where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
                
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTarimas, "Select Linea, Descrip From Lineas where Linea Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RTarimas, "Select Linea, Descrip From Lineas where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            End If
    'FICHA TECNICA
    ElseIf (BFichaTecnica = True Or BFichaTecnicaL = True) Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                    If OptBusqueda.Item(0).Value = True Then
                            Call Abrir_Recordset(RTarimas, "Select Esp_Tec, Descrip From FichaTecnica where Descrip Like '%" & TxtBusqueda.Text & "%' Order By Esp_Tec")
                    Else 'ORACLE
                            Call Abrir_Recordset(RTarimas, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%' Order By Esp_Tec")
                    End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If OptBusqueda.Item(0).Value = True Then
                            Call Abrir_Recordset(RTarimas, "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '%" & TxtBusqueda.Text & "%' Order By Esp_Tec")
                    Else 'ORACLE
                            Call Abrir_Recordset(RTarimas, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%' Order By Esp_Tec")
                    End If
            End If
    End If
            
            Set DBGridBusqueda.DataSource = RTarimas
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
        If Index = 2 Then
        'BUSCA LA DESCRIPCION DE FICHA TECNICA
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where Esp_Tec = '" & TxtTexto.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                End If
                        If RBuscaFichaTecnica.RecordCount > 0 Then
                            LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
                        Else
                            LblFichaTecnica.Caption = ""
                        End If
        'BUSCA LA DESCRIPCION DE LINEA
        ElseIf Index = 1 Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(2).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(2).Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        ElseIf Index = 6 Then
        'BUSCA LA DESCRIPCION DE FICHA TECNICA
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where Esp_Tec = '" & TxtTexto.Item(6).Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(6).Text) & "'")
                End If
                
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblFichaTecnica2.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblFichaTecnica2.Caption = ""
                End If
        'BUSCA LA DESCRIPCION DE LINEA
        ElseIf Index = 5 Then
            Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(5).Text & "'")
            Else
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(5).Text) & "'")
            End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea2.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea2.Caption = ""
                End If
        End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        Set RBusqueda = New ADODB.Recordset
        'SI ELIGE FICHA TECNICA
        If Index = 2 Then
            BFichaTecnica = True
            BLinea = False
            BFichaTecnicaL = False
            BLineaL = False
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
        'LINEAS
        ElseIf Index = 1 Then
            BFichaTecnica = False
            BLinea = True
            BFichaTecnicaL = False
            BLineaL = False
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
        'SI ELIGE FICHA TECNICA
        ElseIf Index = 6 Then
            BFichaTecnicaL = True
            BLineaL = False
            BFichaTecnica = False
            BLinea = False
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
        'LINEAS
        ElseIf Index = 5 Then
            BFichaTecnicaL = False
            BLineaL = True
            BFichaTecnica = False
            BLinea = False
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
        End If
        
        If (Index = 1 Or Index = 2 Or Index = 5 Or Index = 6) Then
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
            Set RBusqueda = New ADODB.Recordset
            'SI ELIGE FICHA TECNICA
            If Index = 2 Then
                BFichaTecnica = True
                BLinea = False
                BFichaTecnicaL = False
                BLineaL = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
            'LINEAS
            ElseIf Index = 1 Then
                BFichaTecnica = False
                BLinea = True
                BFichaTecnicaL = False
                BLineaL = False
                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
            'SI ELIGE FICHA TECNICA
            ElseIf Index = 6 Then
                BFichaTecnicaL = True
                BLineaL = False
                BFichaTecnica = False
                BLinea = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
            'LINEAS
            ElseIf Index = 5 Then
                BFichaTecnicaL = False
                BLineaL = True
                BFichaTecnica = False
                BLinea = False
                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
            End If
            
            If (Index = 1 Or Index = 2 Or Index = 5 Or Index = 6) Then
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
            End If
    End If
    
End Sub

Private Sub Txttexto_LostFocus(Index As Integer)
            If Index = 0 Then
                    TxtTexto.Item(0).Text = Format(TxtTexto.Item(0).Text, "dd/mm/yyyy")
            ElseIf Index = 4 Then
                    TxtTexto.Item(4).Text = Format(TxtTexto.Item(4).Text, "dd/mm/yyyy")
            ElseIf Index = 7 Then
            'BUSCA LA CALIDAD DE LA TARIMA
                     Set RBuscaTarima2 = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaTarima2, "Select Calidad from produccion Where Linea = '" & TxtTexto.Item(5) & "' and Esp_tec = '" & UCase(TxtTexto.Item(6).Text) & "' and Fec_prd = #" & Format(TxtTexto.Item(4).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(7).Text)
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaTarima2, "Select Calidad from produccion Where UPPER(Linea) = '" & UCase(TxtTexto.Item(5)) & "' and UPPER(Esp_tec) = '" & UCase(TxtTexto.Item(6).Text) & "' and Fec_prd = TO_DATE('" & TxtTexto.Item(4).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(7).Text)
                        End If
                         If RBuscaTarima2.RecordCount > 0 Then
                            TxtTexto.Item(8).Text = RBuscaTarima2!Calidad
                         Else
                            TxtTexto.Item(8).Text = ""
                        End If
            End If
                   
End Sub

Public Sub Llena_Campos()
On Error Resume Next
        'FECHA
            If IsNull(RTarimas!fec_prd) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = RTarimas!fec_prd
            End If
        'LINEA
            If IsNull(RTarimas!Linea) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = RTarimas!Linea
            End If
        'FICHA
            If IsNull(RTarimas!Esp_Tec) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RTarimas!Esp_Tec
            End If
        'TARIMA
            If IsNull(RTarimas!Tarima) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = RTarimas!Tarima
            End If
        'FECHA LIBERADA
            If IsNull(RTarimas!fec_prdL) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = RTarimas!fec_prdL
            End If
        'LINEA
            If IsNull(RTarimas!LineaL) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = RTarimas!LineaL
            End If
        'FICHA
            If IsNull(RTarimas!Esp_TecL) Then
                TxtTexto.Item(6).Text = ""
            Else
                TxtTexto.Item(6).Text = RTarimas!Esp_TecL
            End If
        'TARIMA
            If IsNull(RTarimas!TarimaL) Then
                TxtTexto.Item(7).Text = ""
            Else
                TxtTexto.Item(7).Text = RTarimas!TarimaL
            End If
        'CALIDAD
            If IsNull(RTarimas!CalidadL) Then
                TxtTexto.Item(8).Text = ""
            Else
                TxtTexto.Item(8).Text = RTarimas!CalidadL
            End If
        
        'REVISADOS
            If IsNull(RTarimas!Revisados) Then
                TxtTexto.Item(9).Text = ""
            Else
                TxtTexto.Item(9).Text = RTarimas!Revisados
            End If
        'NOCONFORME
            If IsNull(RTarimas!NoConforme) Then
                TxtTexto.Item(10).Text = ""
            Else
                TxtTexto.Item(10).Text = RTarimas!NoConforme
            End If
        'LIBERADOS
            If IsNull(RTarimas!Liberados) Then
                TxtTexto.Item(11).Text = ""
            Else
                TxtTexto.Item(11).Text = RTarimas!Liberados
            End If
        'EN TARIMA
            If IsNull(RTarimas!EnTarima) Then
                TxtTexto.Item(14).Text = ""
            Else
                TxtTexto.Item(14).Text = RTarimas!EnTarima
            End If
        'MINUTOS
            If IsNull(RTarimas!Minutos) Then
                TxtTexto.Item(12).Text = ""
            Else
                TxtTexto.Item(12).Text = RTarimas!Minutos
            End If
        'Empleado
            If IsNull(RTarimas!Empleado) Then
                TxtTexto.Item(13).Text = ""
            Else
                TxtTexto.Item(13).Text = RTarimas!Empleado
            End If
        
        
        If Err <> 0 Then
            
        End If

End Sub

Public Sub Limpia_Campos()
        
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(3).Text = 0
        TxtTexto.Item(4).Text = ""
        TxtTexto.Item(5).Text = ""
        TxtTexto.Item(6).Text = ""
        TxtTexto.Item(7).Text = 0
        TxtTexto.Item(8).Text = ""
        TxtTexto.Item(9).Text = 0
        TxtTexto.Item(10).Text = 0
        TxtTexto.Item(11).Text = 0
        TxtTexto.Item(12).Text = 0
        TxtTexto.Item(13).Text = ""
        TxtTexto.Item(14).Text = 0
        
End Sub



