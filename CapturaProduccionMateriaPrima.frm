VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CapturaProduccionMateriaPrima 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Materia Prima De Produccion"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   Icon            =   "CapturaProduccionMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
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
      Height          =   6495
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   5175
         Left            =   120
         TabIndex        =   23
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   9128
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
         TabIndex        =   20
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
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7440
         Picture         =   "CapturaProduccionMateriaPrima.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   240
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":237C
      Picture         =   "CapturaProduccionMateriaPrima.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Primer Registro"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   600
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":2CF0
      Picture         =   "CapturaProduccionMateriaPrima.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Registro Anterior"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":3664
      Picture         =   "CapturaProduccionMateriaPrima.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Siguiente Registro"
      Top             =   5880
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":3FD8
      Picture         =   "CapturaProduccionMateriaPrima.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Ultimo Registro"
      Top             =   5880
      Width           =   375
   End
   Begin TabDlg.SSTab TabMateriaPrima 
      Height          =   5535
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   255
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CapturaProduccionMateriaPrima.frx":494C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameMateriaPrima"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaProduccionMateriaPrima.frx":4C66
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridMateriaPrima"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CapturaProduccionMateriaPrima.frx":50B8
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
      Begin MSDataGridLib.DataGrid DbGridMateriaPrima 
         Height          =   4695
         Left            =   -74880
         TabIndex        =   42
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8281
         _Version        =   393216
         AllowUpdate     =   -1  'True
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
            Caption         =   "Ficha  Tecnica"
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
            DataField       =   "FechaProduccion"
            Caption         =   "Fecha Entrada"
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
            DataField       =   "LineaProduccion"
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
         BeginProperty Column06 
            DataField       =   "CodigoMateriaPrima"
            Caption         =   "Materia Prima"
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
            DataField       =   "Bulto"
            Caption         =   "Bulto"
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
               Locked          =   -1  'True
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   615.118
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1214.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1379.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   480.189
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPFecFin 
         Height          =   255
         Left            =   -68040
         TabIndex        =   38
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17039363
         CurrentDate     =   37522
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   255
         Left            =   -70200
         TabIndex        =   37
         Top             =   2160
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17039363
         CurrentDate     =   37522
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "CapturaProduccionMateriaPrima.frx":550A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "CapturaProduccionMateriaPrima.frx":5814
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3240
         Width           =   2055
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   15
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   2760
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
         Height          =   740
         Left            =   -74880
         TabIndex        =   28
         Top             =   960
         Width           =   4212
         Begin VB.OptionButton OptFechasYCodigo 
            Caption         =   "Fechas Y Codigo"
            Height          =   192
            Left            =   2520
            TabIndex        =   41
            Top             =   360
            Width           =   1572
         End
         Begin VB.OptionButton OptFechas 
            Caption         =   "Fechas"
            Height          =   225
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptFechasYLinea 
            Caption         =   "Fechas Y Linea"
            Height          =   195
            Left            =   1080
            TabIndex        =   14
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame FrameMateriaPrima 
         Caption         =   "Datos De Materia Prima Por Tarima"
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
         Height          =   3615
         Left            =   120
         TabIndex        =   25
         Top             =   1080
         Width           =   8235
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   7
            Left            =   1200
            TabIndex        =   5
            ToolTipText     =   " "
            Top             =   2520
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   6
            Left            =   1200
            TabIndex        =   4
            ToolTipText     =   " "
            Top             =   2160
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   1
            Left            =   1200
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   5
            Left            =   1200
            TabIndex        =   7
            ToolTipText     =   " "
            Top             =   3240
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   4
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   6
            ToolTipText     =   " "
            Top             =   2880
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
            MaxLength       =   2
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   1575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Height          =   285
            Index           =   0
            Left            =   1200
            MaxLength       =   15
            TabIndex        =   2
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Line Line1 
            X1              =   0
            X2              =   8160
            Y1              =   1920
            Y2              =   1920
         End
         Begin VB.Label Label2 
            Caption         =   "Linea"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   48
            Top             =   2520
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Entrada"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   47
            Top             =   2160
            Width           =   1050
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Bulto"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   3240
            Width           =   360
         End
         Begin VB.Label LblMateriaPrima 
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
            TabIndex        =   35
            Top             =   2880
            Width           =   5175
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Materia Prima"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   2880
            Width           =   960
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tarima"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Linea"
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   975
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
            TabIndex        =   31
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
            TabIndex        =   30
            Top             =   1080
            Width           =   5175
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   1080
            Width           =   1020
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha "
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   975
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
         TabIndex        =   40
         Top             =   2160
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
         Left            =   -68640
         TabIndex        =   39
         Top             =   2160
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
         TabIndex        =   29
         Top             =   2760
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6360
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":5C56
      Picture         =   "CapturaProduccionMateriaPrima.frx":6098
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5760
      Width           =   1080
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5040
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":65B3
      Picture         =   "CapturaProduccionMateriaPrima.frx":69F5
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5760
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   3720
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":6FBD
      Picture         =   "CapturaProduccionMateriaPrima.frx":73FF
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5760
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2400
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":7936
      Picture         =   "CapturaProduccionMateriaPrima.frx":7D78
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5760
      Width           =   1200
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   1080
      MouseIcon       =   "CapturaProduccionMateriaPrima.frx":82D4
      Picture         =   "CapturaProduccionMateriaPrima.frx":8716
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5760
      Width           =   1200
   End
End
Attribute VB_Name = "CapturaProduccionMateriaPrima"
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
Dim VTexto As String
Dim VMensaje As String

Dim BFichaTecnica As Boolean
Dim BLinea As Boolean
Dim BMateriaPrima As Boolean

Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaMateriaPrima As New ADODB.Recordset
Dim RBuscaAtributos As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RMateriaPrima As New ADODB.Recordset

Dim VSumaCriticos As Long
Dim VSumaMayores As Long
Dim VSumaMenores As Long




Sub botones()
    If Bandera = True Then
         FrameMateriaPrima.Enabled = True
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
         DbGridMateriaPrima.Visible = False
    Else
         FrameMateriaPrima.Enabled = False
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
         DbGridMateriaPrima.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        
            If Index = 0 Then
                    Bandera = True
                    botones
                    Limpia_Campos
                    
                    TxtTexto.Item(0).Text = VPFicha
                    TxtTexto.Item(1).Text = VPFecha
                    TxtTexto.Item(2).Text = VPLinea
                    TxtTexto.Item(3).Text = VPTarima
                    
                    TxtTexto.Item(4).SetFocus
            'GRABAR
            ElseIf Index = 2 Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                    Else
                        TxtTexto.Item(1).Text = Format(TxtTexto.Item(1).Text, "dd/mm/yyyy")
                    End If
            
                     If Not IsDate(TxtTexto.Item(1).Text) Then
                        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        TxtTexto.Item(1).SetFocus
                        Exit Sub
                    End If
                   
                     If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "Values(#" & Format(TxtTexto.Item(1).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = "Values(To_Date('" & Format(TxtTexto.Item(1).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            
                            VTexto = VTexto & "'" & TxtTexto.Item(2).Text & "', '" 'LINEA
                            VTexto = VTexto & TxtTexto.Item(0).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(3).Text & ", '" 'TARIMA
                            
                            
                            VTexto = VTexto & TxtTexto.Item(4).Text & "', '" 'CODIGO MATERIA PRIMA
                            VTexto = VTexto & TxtTexto.Item(5).Text & "')" 'BULTO
                            
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into ProduccionConMateriaPrima " & VTexto
                     
                     
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
                        ElseIf Err = -2147467259 Then
                            MsgBox "Fecha, Linea, FichaTecnica, Tarima No Existe, En Produccion", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                        ElseIf Err <> -2147217873 And Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
                                                
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RMateriaPrima.Requery
                        RMateriaPrima.MoveLast
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
                        RMateriaPrima.Delete
                        
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
                        RMateriaPrima.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RMateriaPrima.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
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
        RMateriaPrima.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RMateriaPrima.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RMateriaPrima.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RMateriaPrima.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RMateriaPrima.BOF Then
        RMateriaPrima.MoveFirst
    ElseIf RMateriaPrima.EOF Then
        RMateriaPrima.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        Set RMateriaPrima = New ADODB.Recordset
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptFechas.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RMateriaPrima, "Select * from ProduccionConMateriaPrima where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order by Fec_prd, Linea, Esp_Tec, Tarima")
                Else 'ORACLE
                    Call Abrir_Recordset(RMateriaPrima, "Select * from ProduccionConMateriaPrima where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " Order by Fec_prd, Linea, Esp_Tec, Tarima")
                End If
            ElseIf OptFechasYLinea.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RMateriaPrima, "Select * from ProduccionConMateriaPrima where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBuscar.Text & "' Order by Fec_prd, Linea, Esp_Tec, Tarima")
                Else 'ORACLE
                    Call Abrir_Recordset(RMateriaPrima, "Select * from ProduccionConMateriaPrima where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And Upper(Linea) = '" & UCase(TxtBuscar.Text) & "' Order by Fec_prd, Linea, Esp_Tec, Tarima")
                End If
            ElseIf OptFechasYCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RMateriaPrima, "Select * from ProduccionConMateriaPrima where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Esp_Tec = '" & TxtBuscar.Text & "' Order by Fec_prd, Linea, Esp_Tec, Tarima")
                Else 'ORACLE
                    Call Abrir_Recordset(RMateriaPrima, "Select * from ProduccionConMateriaPrima where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And Upper(Esp_Tec) = '" & UCase(TxtBuscar.Text) & "' Order by Fec_prd, Linea, Esp_Tec, Tarima")
                End If
            End If
                
                Set DbGridMateriaPrima.DataSource = RMateriaPrima
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RMateriaPrima, "Select * From ProduccionConMateriaPrima")
                Set DbGridMateriaPrima.DataSource = RMateriaPrima
        End If
    
        TabMateriaPrima.Tab = 1
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
            ElseIf BMateriaPrima = True Then
                TxtTexto.Item(4).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(4).SetFocus
            End If
                FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                If BFichaTecnica = True Then
                    TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0).Text
                    TxtTexto.Item(0).SetFocus
                ElseIf BLinea = True Then
                    TxtTexto.Item(2).Text = DBGridBusqueda.Columns(0).Text
                    TxtTexto.Item(2).SetFocus
                ElseIf BMateriaPrima = True Then
                    TxtTexto.Item(4).Text = DBGridBusqueda.Columns(0).Text
                    TxtTexto.Item(4).SetFocus
                End If
                    FrameBusqueda.Visible = False
            End If
                    
End Sub

Private Sub DbGridMateriaPrima_DblClick()
        Llena_Campos
End Sub

Private Sub DBGridMateriaPrima_HeadClick(ByVal ColIndex As Integer)
            RMateriaPrima.Sort = RMateriaPrima.Fields(ColIndex).Name
End Sub


Private Sub DbGridMateriaPrima_KeyDown(KeyCode As Integer, Shift As Integer)
        Llena_Campos
End Sub

Private Sub DbGridMateriaPrima_KeyUp(KeyCode As Integer, Shift As Integer)
        Llena_Campos
End Sub

Private Sub Form_Load()
        On Error Resume Next
        Set RMateriaPrima = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RMateriaPrima, "Select * From ProduccionConMateriaPrima Where Fec_Prd = #" & Format(Date, "mm/dd/yyyy") & "#")
            Else 'ORACLE
                Call Abrir_Recordset(RMateriaPrima, "Select * From ProduccionConMateriaPrima Where Fec_Prd = To_Date('" & Date & "', 'dd/mm/yyyy')")
            End If
        
        Set DbGridMateriaPrima.DataSource = RMateriaPrima
        RMateriaPrima.MoveLast
        If Err <> 0 Then
        End If
        Llena_Campos
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date
End Sub


Private Sub OptFechas_Click()
        Lbletiqueta.Caption = ""
End Sub

Private Sub OptFechasYCodigo_Click()
        Lbletiqueta.Caption = "Codigo"
End Sub

Private Sub OptFechasYlinea_Click()
        Lbletiqueta.Caption = "Linea"
End Sub

Private Sub TabMateriaPrima_Click(PreviousTab As Integer)
        If TabMateriaPrima.Tab = 0 Then
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
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Esp_Tec Like '%" & TxtBusqueda.Text & "%' Order By Esp_Tec")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%' Order By Esp_Tec")
                End If
                
            End If
    'DEFECTO
    ElseIf BMateriaPrima = True Then
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Esp_Tec Like '%" & TxtBusqueda.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%'")
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
        If Index = 0 Then
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
        ElseIf Index = 2 Then
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
                
        'BUSCA LA DESCRIPCION DE FichaTecnica
        ElseIf Index = 4 Then
            Set RBuscaMateriaPrima = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(4).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(4).Text) & "'")
                End If
                If RBuscaMateriaPrima.RecordCount > 0 Then
                    LblMateriaPrima.Caption = RBuscaMateriaPrima!Descrip
                Else
                    LblMateriaPrima.Caption = ""
                End If
        End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        Set RBusqueda = New ADODB.Recordset
        
        'SI ELIGE FICHA TECNICA
        If Index = 0 Then
            BFichaTecnica = True
            BLinea = False
            BMateriaPrima = False
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
        'LINEAS
        ElseIf Index = 2 Then
            BFichaTecnica = False
            BLinea = True
            BMateriaPrima = False
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
        'FichaTecnica
        ElseIf Index = 4 Then
            BFichaTecnica = False
            BLinea = False
            BMateriaPrima = True
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
        End If
        
        If (Index = 0 Or Index = 2 Or Index = 4) Then
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
                If Index = 0 Then
                    BFichaTecnica = True
                    BLinea = False
                    BMateriaPrima = False
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                'LINEAS
                ElseIf Index = 2 Then
                    BFichaTecnica = False
                    BLinea = True
                    BMateriaPrima = False
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
                'FichaTecnica
                ElseIf Index = 4 Then
                    BFichaTecnica = False
                    BLinea = False
                    BMateriaPrima = True
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                End If
                
                If (Index = 0 Or Index = 2 Or Index = 4) Then
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
            If IsNull(RMateriaPrima!fec_prd) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = RMateriaPrima!fec_prd
            End If
        'LINEA
            If IsNull(RMateriaPrima!Linea) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RMateriaPrima!Linea
            End If
        'FICHA
            If IsNull(RMateriaPrima!Esp_Tec) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = RMateriaPrima!Esp_Tec
            End If
        'TARIMA
            If IsNull(RMateriaPrima!Tarima) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = RMateriaPrima!Tarima
            End If
        'CODIGO MATERIA PRIMA
            If IsNull(RMateriaPrima!Esp_Tec) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = RMateriaPrima!Esp_Tec
            End If
        'BULTO
            If IsNull(RMateriaPrima!Bulto) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = RMateriaPrima!Bulto
            End If
        'FECHA ENTRADA
            If IsNull(RMateriaPrima!FechaProduccion) Then
                TxtTexto.Item(6).Text = ""
            Else
                TxtTexto.Item(6).Text = RMateriaPrima!FechaProduccion
            End If
        'LINEA
            If IsNull(RMateriaPrima!LineaProduccion) Then
                TxtTexto.Item(7).Text = ""
            Else
                TxtTexto.Item(7).Text = RMateriaPrima!LineaProduccion
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
        TxtTexto.Item(5).Text = 0
        TxtTexto.Item(6).Text = ""
        TxtTexto.Item(7).Text = ""
        
End Sub




