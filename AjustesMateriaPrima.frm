VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AjustesMateriaPrima 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajustes De Materia Prima"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8475
   Icon            =   "AjustesMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   8475
   StartUpPosition =   2  'CenterScreen
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
      Height          =   5655
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DbGridBuscar 
         Height          =   4335
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   7646
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
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7320
         Picture         =   "AjustesMateriaPrima.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Txtbusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   26
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   5775
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   24
         Top             =   360
         Width           =   1455
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
         TabIndex        =   29
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "AjustesMateriaPrima.frx":293C
      Picture         =   "AjustesMateriaPrima.frx":2D7E
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Primer Registro"
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "AjustesMateriaPrima.frx":32B0
      Picture         =   "AjustesMateriaPrima.frx":36F2
      Style           =   1  'Graphical
      TabIndex        =   46
      ToolTipText     =   "Registro Anterior"
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "AjustesMateriaPrima.frx":3C24
      Picture         =   "AjustesMateriaPrima.frx":4066
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Siguiente Registro"
      Top             =   4920
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "AjustesMateriaPrima.frx":4598
      Picture         =   "AjustesMateriaPrima.frx":49DA
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Ultimo Registro"
      Top             =   4920
      Width           =   375
   End
   Begin TabDlg.SSTab TabPuestos 
      Height          =   4695
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "AjustesMateriaPrima.frx":4F0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePuestos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "AjustesMateriaPrima.frx":5226
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGrid"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "AjustesMateriaPrima.frx":5678
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "Label1(1)"
      Tab(2).Control(2)=   "Label1(2)"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "CmdBuscar(0)"
      Tab(2).Control(5)=   "CmdBuscar(1)"
      Tab(2).Control(6)=   "TxtBuscar"
      Tab(2).Control(7)=   "DTPFecIni"
      Tab(2).Control(8)=   "DTPFecFin"
      Tab(2).ControlCount=   9
      Begin MSDataGridLib.DataGrid DbGrid 
         Height          =   3855
         Left            =   -74880
         TabIndex        =   42
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6800
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
            DataField       =   "FechaOperacion"
            Caption         =   "FechaOperacion"
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
            DataField       =   "Fecha"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "Efecto"
            Caption         =   "Efecto"
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
            DataField       =   "CodigoMateriaPrima"
            Caption         =   "CodigoMateriaPrima"
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
            DataField       =   "NumeroIngreso"
            Caption         =   "NumeroIngreso"
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
         BeginProperty Column07 
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
         BeginProperty Column08 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   225.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1065.26
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPFecFin 
         Height          =   255
         Left            =   -68760
         TabIndex        =   37
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62128131
         CurrentDate     =   38127
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   255
         Left            =   -68760
         TabIndex        =   36
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   62128131
         CurrentDate     =   38127
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   39
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   2280
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "AjustesMateriaPrima.frx":5ACA
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "AjustesMateriaPrima.frx":5DD4
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2640
         Width           =   2055
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
         Height          =   1215
         Left            =   -74880
         TabIndex        =   19
         Top             =   960
         Width           =   2445
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Fechas"
            Height          =   225
            Left            =   120
            TabIndex        =   13
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptDescripcion 
            Caption         =   "Fechas y Codigo"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            ToolTipText     =   " "
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame FramePuestos 
         Caption         =   "Datos Del Ajuste"
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
         Height          =   3735
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   8175
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   7
            Top             =   3000
            Width           =   6495
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   5
            Top             =   2280
            Width           =   1455
         End
         Begin VB.ComboBox CboEfecto 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "AjustesMateriaPrima.frx":6216
            Left            =   1560
            List            =   "AjustesMateriaPrima.frx":6220
            TabIndex        =   3
            Text            =   "+"
            Top             =   1440
            Width           =   1455
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   0
            Left            =   1560
            TabIndex        =   0
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Index           =   2
            Left            =   1560
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   3360
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   4
            Top             =   1920
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Documento"
            Height          =   285
            Index           =   0
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   2
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   1080
            Width           =   1455
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   1
            Left            =   1560
            TabIndex        =   1
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   2
            Left            =   1560
            TabIndex        =   6
            Top             =   2640
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   35
            Top             =   3000
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Suma o Resta"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   33
            Top             =   1920
            Width           =   495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "# Bulto"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   31
            Top             =   2640
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   1080
            Width           =   825
         End
         Begin VB.Label LblCod 
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
            Left            =   3120
            TabIndex        =   22
            Top             =   1920
            Width           =   4935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   21
            Top             =   3360
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operacion"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   17
            Top             =   720
            Width           =   450
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
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
         Left            =   -69960
         TabIndex        =   40
         Top             =   1800
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
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
         Left            =   -69960
         TabIndex        =   38
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
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
         Left            =   -69720
         TabIndex        =   20
         Top             =   2280
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   5880
      MouseIcon       =   "AjustesMateriaPrima.frx":622A
      Picture         =   "AjustesMateriaPrima.frx":666C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   1620
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4320
      MouseIcon       =   "AjustesMateriaPrima.frx":86DE
      Picture         =   "AjustesMateriaPrima.frx":8B20
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   1500
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2760
      MouseIcon       =   "AjustesMateriaPrima.frx":9052
      Picture         =   "AjustesMateriaPrima.frx":9494
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   1500
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   1080
      MouseIcon       =   "AjustesMateriaPrima.frx":99C6
      Picture         =   "AjustesMateriaPrima.frx":9E08
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   1620
   End
End
Attribute VB_Name = "AjustesMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim RBuscaCodigo As New ADODB.Recordset
Dim RBuscaNumeroIngreso As New ADODB.Recordset
Dim RMateriaPrima As New ADODB.Recordset
Dim RAjustesMateriaPrima As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VCodigo As String
Dim VBulto As Long
Dim VCantidad As Single
Dim VTipo As String

Dim VValores As String
Dim VCampos As String




Sub botones()
    If Bandera = True Then
         FramePuestos.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
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
         DBGrid.Visible = False
         
         CmdBuscar.Item(0).Visible = False
         CmdBuscar.Item(1).Visible = False
    Else
         FramePuestos.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(5).Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         FrameOpciones.Visible = True
         DBGrid.Visible = True
         
         CmdBuscar.Item(0).Visible = True
         CmdBuscar.Item(1).Visible = True
    End If
End Sub



Private Sub CboEfecto_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        
                If Index = 0 Then
                    'AGREGA UN REGISTRO
                    Limpia_Campos
                    Bandera = True
                    botones
                    Msk.Item(0).Text = Date
                    Msk.Item(1).SetFocus
                    TxtTexto.Item(2).Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                     'REVISA LA FECHA
                     If Not IsDate(Msk.Item(1).Text) Then
                        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(1).SetFocus
                        Exit Sub
                    End If
                    
                    'EFECTO
                    If CboEfecto.Text <> "+" And CboEfecto.Text <> "-" Then
                        MsgBox "Suma O Resta Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        CboEfecto.SetFocus
                        Exit Sub
                    End If
                    
                    'NUMERO DE BULTO
                    If Not IsNumeric(TxtTexto.Item(3).Text) Then
                        MsgBox "Numero De Ingreso Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        TxtTexto.Item(3).SetFocus
                        Exit Sub
                    End If
                     
                    'CANTIDAD INCORRECTA
                    If Not IsNumeric(Msk.Item(2).Text) Then
                        MsgBox "Cantidad Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(2).SetFocus
                        Exit Sub
                    End If
                     
                    Set RBuscaNumeroIngreso = New ADODB.Recordset
                    'BUSCA EL NUMERO DE INGRESO
                    Call Abrir_Recordset(RBuscaNumeroIngreso, "Select * From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtTexto.Item(3).Text & " And Codigo = '" & TxtTexto.Item(1).Text & "'")
                    'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
                    If RBuscaNumeroIngreso.RecordCount > 0 Then
                    Else
                        MsgBox "El Numero De Bulto Para Esta Materia Prima, No Exite", vbOKOnly + vbExclamation, "Informacion"
                        Exit Sub
                    End If
                    
                     'GUARDA VARIABLES
                     VCodigo = TxtTexto.Item(1).Text
                     VBulto = TxtTexto.Item(3).Text
                     VCantidad = Msk.Item(2).Text
                     VTipo = CboEfecto.Text
                    
                     'GRABA EL REGISTRO
                     VCampos = "FechaOperacion, Fecha, Documento, Efecto, CodigoMateriaPrima, NumeroIngreso, Cantidad, Observaciones, Usuario"
                                                
                        If GOrigenDeDatos = "AmaproAccess" Then
                             VValores = "#" & Format(Msk.Item(0).Text, "mm/dd/yyyy") & "#," 'FECHA Operacion
                             VValores = VValores & "#" & Format(Msk.Item(1).Text, "mm/dd/yyyy") & "#," 'FECHA
                        Else 'ORACLE
                             VValores = "To_Date('" & Format(Msk.Item(0).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                             VValores = VValores & "To_Date('" & Format(Msk.Item(1).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                        End If
                        VValores = VValores & TxtTexto.Item(0).Text & "," 'DOCUMENTO
                        VValores = VValores & "'" & CboEfecto.Text & "'," 'EFECTO
                        VValores = VValores & "'" & TxtTexto.Item(1).Text & "'," 'CODIGO MATERIA PRIMA
                        VValores = VValores & TxtTexto.Item(3).Text & "," 'NUMERO INGRESO
                        VValores = VValores & Msk.Item(2).Text & "," 'CANTIDAD
                        VValores = VValores & "'" & TxtTexto.Item(4).Text & "'," 'OBSERVACIONES
                        VValores = VValores & "'" & TxtTexto.Item(2).Text & "'" 'USUARIO
                        
                        'INICIA UNA TRANSACCION
                       'SI ESTA GRABANDO UN REGISTRO NUEVO
                        Conexion.BeginTrans
                                Conexion.Execute "Insert Into AjustesMateriaPrima (" & VCampos & ") Values(" & VValores & ")"
                        
                                 'SI ES CUALQUIER OTRO ERROR
                                If Err <> 0 Then
                                   MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                   Conexion.RollbackTrans
                                   Exit Sub
                                End If
                     
                                'RESTAR EL SALDO
                                If VTipo = "-" Then
                                       Conexion.Execute "Update DetalleEntradasMateriaPrima set saldodisponibilidad = (saldodisponibilidad - " & VCantidad & ") where codigo = '" & VCodigo & "' And NumeroIngreso = " & VBulto
                                   If Err <> 0 Then
                                       MsgBox "No Pudo Rebajar El Saldo, No Se Puede Grabar " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                       Conexion.RollbackTrans
                                       Exit Sub
                                   End If
                                'AUMENTAR EL SALDO
                                ElseIf VTipo = "+" Then
                                       Conexion.Execute "update DetalleEntradasMateriaPrima set saldodisponibilidad = (saldodisponibilidad + " & VCantidad & ") where codigo = '" & VCodigo & "' And NumeroIngreso = " & VBulto
                                   If Err <> 0 Then
                                       MsgBox "No Pudo Aumentar El Saldo, No Se Puede Grabar " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                                       Conexion.RollbackTrans
                                       Exit Sub
                                   End If
                                End If
                                   
                        'FINALIZA LA TRANSACCIOON
                        Conexion.CommitTrans
                        'VUELVE A LLENAR EL RECORDSET
                        RAjustesMateriaPrima.Requery
                    
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
            'CANCELAR
            ElseIf Index = 3 Then
                    
                    Bandera = False
                    botones
                    
                    Llena_Campos
            'SALIDA
            ElseIf Index = 5 Then
                    Unload Me
            End If
        
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RAjustesMateriaPrima.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RAjustesMateriaPrima.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RAjustesMateriaPrima.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RAjustesMateriaPrima.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RAjustesMateriaPrima.BOF Then
        RAjustesMateriaPrima.MoveFirst
    ElseIf RAjustesMateriaPrima.EOF Then
        RAjustesMateriaPrima.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        'INICIALIZAMOS EL RECORDSET
        Set RAjustesMateriaPrima = New ADODB.Recordset
    
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RAjustesMateriaPrima, "Select * from AjustesMateriaPrima where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#")
                Else 'ORACLE
                    Call Abrir_Recordset(RAjustesMateriaPrima, "Select * from AjustesMateriaPrima where Fecha >= to_date('" & DtpFecIni.Value & "', 'dd/mm/yyyy') And Fecha <= to_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy')")
                End If
            ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RAjustesMateriaPrima, "Select * from AjustesMateriaPrima where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And CodigoMateriaPrima Like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RAjustesMateriaPrima, "Select * from AjustesMateriaPrima where Fecha >= to_date('" & DtpFecIni.Value & "', 'dd/mm/yyyy') And Fecha <= to_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy') And UPPER(CodigoMateriaPrima) Like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RAjustesMateriaPrima, "Select * From AjustesMateriaPrima")
        End If
        
            'LLENA EL GRID
            Set DBGrid.DataSource = RAjustesMateriaPrima
    
        TabPuestos.Tab = 1
End Sub


Private Sub CmdSale_Click()
        Framebuscar.Visible = False
End Sub


Private Sub DBGrid_HeadClick(ByVal ColIndex As Integer)
On Error Resume Next
        If RAjustesMateriaPrima.RecordCount > 0 Then
            RAjustesMateriaPrima.Sort = RAjustesMateriaPrima.Fields(ColIndex).Name
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            End If
        End If
End Sub

Private Sub DBGridBuscar_DblClick()
        TxtTexto.Item(1).Text = DBGridBuscar.Columns(0)
        TxtTexto.Item(1).SetFocus
        Framebuscar.Visible = False

End Sub

Private Sub DbGridBuscar_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        TxtTexto.Item(1).Text = DBGridBuscar.Columns(0)
        TxtTexto.Item(1).SetFocus
        Framebuscar.Visible = False
End Sub

Private Sub Form_Load()
        'INICIALIZA EL RECORDSET
        Set RAjustesMateriaPrima = New ADODB.Recordset
        Call Abrir_Recordset(RAjustesMateriaPrima, "Select * From AjustesMateriaPrima")
        'LLENA EL GRID CON EL RECORDSET
        Set DBGrid.DataSource = RAjustesMateriaPrima
        Llena_Campos
                
        DtpFecIni.Value = Date
        DtpFecFin.Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    'CERRAMOS LOS RECORDSET
        RBuscaCodigo.Close
        RBuscaNumeroIngreso.Close
        RMateriaPrima.Close
        RAjustesMateriaPrima.Close
        RBusqueda.Close
        
        Set RBuscaCodigo = Nothing
        Set RBuscaNumeroIngreso = Nothing
        Set RMateriaPrima = Nothing
        Set RAjustesMateriaPrima = Nothing
        Set RBusqueda = Nothing
        
        'SI NO SE A ABIERTO ALGUN RECORDSET DA ERROR
        If Err <> 0 Then
        End If

End Sub

Private Sub Msk_GotFocus(Index As Integer)
        Msk.Item(Index).SelStart = 0
        Msk.Item(Index).SelLength = Len(Msk.Item(Index))
End Sub

Private Sub Msk_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub OptCodigo_Click()
        Lbletiqueta.Caption = ""
        TxtBuscar.Visible = False
End Sub

Private Sub OptDescripcion_Click()
        Lbletiqueta.Visible = True
        TxtBuscar.Visible = True
        TxtBuscar.SetFocus
End Sub

Private Sub TabPuestos_Click(PreviousTab As Integer)
        If TabPuestos.Tab = 0 Then
            If CmdBotones.Item(2).Enabled = False Then
                Llena_Campos
            End If
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

Private Sub Txtbusqueda_Change()
On Error Resume Next
                    Set RBusqueda = New ADODB.Recordset
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBusqueda, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                    'OPCION DE CODIGO
                    Else
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima Like '%" & TxtBusqueda.Text & "%'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBusqueda, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where UPPER(CodigoMateriaPrima) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                    End If
                            'LLENAMOS EL GRID CON EL RECORDSET
                            Set DBGridBuscar.DataSource = RMateriaPrima
                            DBGridBuscar.Refresh
                            DBGridBuscar.Columns(1).Width = "4000"
                        
                    If Err <> 0 Then
                        MsgBox Err.Description
                    End If
                            

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
        If Index = 1 Then
            'INICIALIZAMOS EL RECORDSET
            Set RBuscaCodigo = New ADODB.Recordset
            'ABRIMOS EL RECORDSET
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaCodigo, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtTexto.Item(1).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaCodigo, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima Where Upper(CodigoMateriaPrima) = '" & UCase(TxtTexto.Item(1).Text) & "'")
            End If
                If RBuscaCodigo.RecordCount > 0 Then
                    LblCod.Caption = RBuscaCodigo!Descripcion
                Else
                    LblCod.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
            If Index = 1 Then
                'INICIALIZAMOS EL RECORDSET
                Set RMateriaPrima = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RMateriaPrima, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima")
            End If
            
            If Index = 1 Then
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBuscar.DataSource = RMateriaPrima
                DBGridBuscar.Columns(1).Width = "4000"
                Framebuscar.Visible = True
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
            If Index = 1 Then
            'INICIALIZAMOS EL RECORDSET
                Set RMateriaPrima = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RMateriaPrima, "Select CodigoMateriaPrima, Descripcion From CorrelativosMateriaPrima")
            End If
            
            If Index = 1 Then
                'INICIALIZAMOS EL GRID CON EL ORIGEN DE DATOS DEL RECORDSET
                Set DBGridBuscar.DataSource = RMateriaPrima
                DBGridBuscar.Columns(1).Width = "4000"
                Framebuscar.Visible = True
                TxtBusqueda.SetFocus
            End If
        End If
End Sub

Private Sub Txttexto_LostFocus(Index As Integer)
        If Index = 3 Then
                If IsNumeric(TxtTexto.Item(3).Text) Then
                    'INICIALIZAMOS EL RECORDSET
                    Set RBuscaNumeroIngreso = New ADODB.Recordset
                    'BUSCA EL NUMERO DE INGRESO
                    Call Abrir_Recordset(RBuscaNumeroIngreso, "Select SaldoDisponibilidad From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtTexto.Item(3).Text & " And Codigo = '" & TxtTexto.Item(1).Text & "'")
                    'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
                    If RBuscaNumeroIngreso.RecordCount > 0 Then
                            Msk.Item(2).Text = RBuscaNumeroIngreso!SaldoDisponibilidad
                    Else
                            Msk.Item(2).Text = "0"
                    End If
                End If
        End If
                    
End Sub

Public Sub Llena_Campos()
On Error Resume Next
            If RAjustesMateriaPrima.RecordCount > 0 Then
                Msk.Item(0).Text = RAjustesMateriaPrima!FechaOperacion
                Msk.Item(1).Text = RAjustesMateriaPrima!fecha
                TxtTexto.Item(0).Text = RAjustesMateriaPrima!Documento
                CboEfecto.Text = RAjustesMateriaPrima!Efecto
                TxtTexto.Item(1).Text = RAjustesMateriaPrima!CodigoMateriaPrima
                TxtTexto.Item(3).Text = RAjustesMateriaPrima!NumeroIngreso
                Msk.Item(2).Text = RAjustesMateriaPrima!Cantidad
                TxtTexto.Item(4).Text = RAjustesMateriaPrima!Observaciones
                TxtTexto.Item(2).Text = RAjustesMateriaPrima!usuario
                If Err <> 0 Then
                    'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                End If
            Else
                Msk.Item(0).Text = ""
                Msk.Item(1).Text = ""
                TxtTexto.Item(0).Text = ""
                CboEfecto.Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(3).Text = ""
                Msk.Item(2).Text = ""
                TxtTexto.Item(4).Text = ""
                TxtTexto.Item(2).Text = ""
            End If

End Sub

Public Sub Limpia_Campos()
                Msk.Item(0).Text = ""
                Msk.Item(1).Text = ""
                TxtTexto.Item(0).Text = ""
                CboEfecto.Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(3).Text = ""
                Msk.Item(2).Text = ""
                TxtTexto.Item(4).Text = ""
                TxtTexto.Item(2).Text = ""
End Sub

