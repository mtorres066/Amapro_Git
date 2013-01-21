VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form AjustesProductoTerminado 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ajustes De Producto Terminado"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "AjustesMateriaProductoTerminado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   8415
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
      Height          =   6495
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DBGridBuscar 
         Height          =   5175
         Left            =   120
         TabIndex        =   47
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
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   7320
         Picture         =   "AjustesMateriaProductoTerminado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Txtbusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   31
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   5775
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   30
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   29
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
         TabIndex        =   33
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "AjustesMateriaProductoTerminado.frx":293C
      Picture         =   "AjustesMateriaProductoTerminado.frx":2D7E
      Style           =   1  'Graphical
      TabIndex        =   52
      ToolTipText     =   "Ultimo Registro"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "AjustesMateriaProductoTerminado.frx":32B0
      Picture         =   "AjustesMateriaProductoTerminado.frx":36F2
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Siguiente Registro"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "AjustesMateriaProductoTerminado.frx":3C24
      Picture         =   "AjustesMateriaProductoTerminado.frx":4066
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Registro Anterior"
      Top             =   5760
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "AjustesMateriaProductoTerminado.frx":4598
      Picture         =   "AjustesMateriaProductoTerminado.frx":49DA
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Primer Registro"
      Top             =   5760
      Width           =   375
   End
   Begin TabDlg.SSTab TabPuestos 
      Height          =   5535
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9763
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "AjustesMateriaProductoTerminado.frx":4F0C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePuestos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "AjustesMateriaProductoTerminado.frx":5226
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGrid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "AjustesMateriaProductoTerminado.frx":5678
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DTPFecFin"
      Tab(2).Control(1)=   "DTPFecIni"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "CmdBuscar(1)"
      Tab(2).Control(4)=   "CmdBuscar(0)"
      Tab(2).Control(5)=   "FrameOpciones"
      Tab(2).Control(6)=   "Label1(2)"
      Tab(2).Control(7)=   "Label1(1)"
      Tab(2).Control(8)=   "Lbletiqueta"
      Tab(2).ControlCount=   9
      Begin MSDataGridLib.DataGrid DbGrid 
         Bindings        =   "AjustesMateriaProductoTerminado.frx":5ACA
         Height          =   4695
         Left            =   -74880
         TabIndex        =   48
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
         ColumnCount     =   11
         BeginProperty Column00 
            DataField       =   "FECHAOPERACION"
            Caption         =   "FECHAOPERACION"
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
            DataField       =   "FECHA"
            Caption         =   "FECHA"
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
            DataField       =   "DOCUMENTO"
            Caption         =   "DOCUMENTO"
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
            DataField       =   "EFECTO"
            Caption         =   "EFECTO"
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
            DataField       =   "FECHAPRODUCCION"
            Caption         =   "FECHAPRODUCCION"
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
            DataField       =   "LINEA"
            Caption         =   "LINEA"
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
            DataField       =   "FICHATECNICA"
            Caption         =   "FICHATECNICA"
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
            DataField       =   "TARIMA"
            Caption         =   "TARIMA"
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
            DataField       =   "CANTIDAD"
            Caption         =   "CANTIDAD"
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
            DataField       =   "OBSERVACIONES"
            Caption         =   "OBSERVACIONES"
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
            DataField       =   "USUARIO"
            Caption         =   "USUARIO"
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
               ColumnWidth     =   1140.095
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   239.811
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   345.26
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1275.024
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   390.047
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   1140.095
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTPFecFin 
         Height          =   255
         Left            =   -68760
         TabIndex        =   41
         Top             =   1800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20971523
         CurrentDate     =   38127
      End
      Begin MSComCtl2.DTPicker DTPFecIni 
         Height          =   255
         Left            =   -68760
         TabIndex        =   40
         Top             =   1440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   20971523
         CurrentDate     =   38127
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -68760
         TabIndex        =   17
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   2280
         Width           =   2085
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68760
         Picture         =   "AjustesMateriaProductoTerminado.frx":5ADD
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3600
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68760
         Picture         =   "AjustesMateriaProductoTerminado.frx":5DE7
         Style           =   1  'Graphical
         TabIndex        =   18
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
         TabIndex        =   24
         Top             =   960
         Width           =   2445
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Fechas"
            Height          =   225
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptDescripcion 
            Caption         =   "Fechas y Ficha Tecnica"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            ToolTipText     =   " "
            Top             =   720
            Width           =   2175
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
         Height          =   4455
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   8175
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   1560
            TabIndex        =   7
            Top             =   3000
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   9
            Top             =   3720
            Width           =   6495
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   6
            Top             =   2640
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
            ItemData        =   "AjustesMateriaProductoTerminado.frx":6229
            Left            =   1560
            List            =   "AjustesMateriaProductoTerminado.frx":6233
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
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1560
            MaxLength       =   2
            TabIndex        =   5
            Top             =   2280
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            TabIndex        =   8
            Top             =   3360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   3
            Left            =   1560
            TabIndex        =   4
            Top             =   1920
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label LblLin 
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
            TabIndex        =   46
            Top             =   2280
            Width           =   4935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tarima"
            Height          =   195
            Index           =   9
            Left            =   120
            TabIndex        =   45
            Top             =   3000
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   44
            Top             =   2640
            Width           =   1020
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   39
            Top             =   3720
            Width           =   1065
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Suma o Resta"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   1005
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Produccion"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   37
            Top             =   1920
            Width           =   1305
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Linea"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   36
            Top             =   2280
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Actual"
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
            Left            =   120
            TabIndex        =   35
            Top             =   3360
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   34
            Top             =   1080
            Width           =   825
         End
         Begin VB.Label LblFic 
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
            TabIndex        =   27
            Top             =   2640
            Width           =   4935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   4080
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Operacion"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   22
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   25
         Top             =   2280
         Width           =   855
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   5760
      MouseIcon       =   "AjustesMateriaProductoTerminado.frx":623D
      Picture         =   "AjustesMateriaProductoTerminado.frx":667F
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5640
      Width           =   1600
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "AjustesMateriaProductoTerminado.frx":86F1
      Picture         =   "AjustesMateriaProductoTerminado.frx":8B33
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5640
      Width           =   1485
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   2640
      MouseIcon       =   "AjustesMateriaProductoTerminado.frx":9065
      Picture         =   "AjustesMateriaProductoTerminado.frx":94A7
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5640
      Width           =   1485
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   1080
      MouseIcon       =   "AjustesMateriaProductoTerminado.frx":99D9
      Picture         =   "AjustesMateriaProductoTerminado.frx":9E1B
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   1485
   End
End
Attribute VB_Name = "AjustesProductoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RMateriaPrima As New ADODB.Recordset
Dim RAjustesProductoTerminado As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VFichaTecnica As String
Dim VFecha As Date
Dim VTarima As Long
Dim VLinea As String
Dim VCantidad As Single
Dim VTipo As String

Dim BLinea As Boolean
Dim BFicha As Boolean

Dim VCampos As String
Dim VValores As String



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
         FrameOpciones.Visible = False
         DbGrid.Visible = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         
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
         FrameOpciones.Visible = True
         DbGrid.Visible = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
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
                    Limpia_Campos
                    Bandera = True
                    botones
                    Msk.Item(0).Text = Date
                    Msk.Item(1).Text = Date
                    Msk.Item(1).SetFocus
                    TxtTexto.Item(2).Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                     If GOrigenDeDatos = "AmaproAccess" Then
                     Else
                        Msk.Item(1).Text = Format(Msk.Item(1).Text, "dd/mm/yyyy")
                     End If
            
                     'REVISA LA FECHA
                     If Not IsDate(Msk.Item(1).Text) Then
                        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(1).SetFocus
                        Exit Sub
                    End If
                    
                    'REVISA LA FECHA DE TARIMA
                    If Not IsDate(Msk.Item(3).Text) Then
                        MsgBox "Fecha De Tarima Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(3).SetFocus
                        Exit Sub
                    End If
                    
                    'EFECTO
                    If CboEfecto.Text <> "+" And CboEfecto.Text <> "-" Then
                        MsgBox "Suma O Resta Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        CboEfecto.SetFocus
                        Exit Sub
                    End If
                    
                    'NUMERO DE TARIMA
                    If Not IsNumeric(TxtTexto.Item(6).Text) Then
                        MsgBox "Numero De Tarima Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        TxtTexto.Item(6).SetFocus
                        Exit Sub
                    End If
                     
                    'CANTIDAD INCORRECTA
                    If Not IsNumeric(Msk.Item(2).Text) Then
                        MsgBox "Cantidad Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(2).SetFocus
                        Exit Sub
                    End If
                     
                    'BUSCA TARIMA
                    Set RBuscaTarima = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaTarima, "Select * From DetalleEntradasInventario Where FechaProduccion = #" & Format(Msk.Item(3).Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTexto.Item(6).Text & " And Linea = '" & TxtTexto.Item(1).Text & "' And FichaTecnica = '" & TxtTexto.Item(3).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaTarima, "Select * From DetalleEntradasInventario Where FechaProduccion = TO_DATE('" & Msk.Item(3).Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTexto.Item(6).Text & " And Linea = '" & TxtTexto.Item(1).Text & "' And FichaTecnica = '" & TxtTexto.Item(3).Text & "'")
                    End If
                        If RBuscaTarima.RecordCount > 0 Then
                        Else
                            MsgBox "Tarima No Existe En Inventario", vbOKOnly + vbInformation, "Informacion"
                            Msk.Item(3).SetFocus
                            Exit Sub
                        End If
                    
                     'GUARDA VARIABLES
                     
                     VFichaTecnica = TxtTexto.Item(3).Text
                     VLinea = TxtTexto.Item(1).Text
                     VTarima = TxtTexto.Item(6).Text
                     VFecha = Msk.Item(3).Text
                     
                     VCantidad = Msk.Item(2).Text
                     VTipo = CboEfecto.Text
                    
                     'GRABA EL REGISTRO
                      VCampos = "FechaOperacion, Fecha, Documento, Efecto, FechaProduccion, Linea, FichaTecnica, Tarima, Cantidad, Observaciones, Usuario"
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                             VValores = "#" & Format(Msk.Item(0).Text, "mm/dd/yyyy") & "#," 'FECHA Operacion
                             VValores = VValores & "#" & Format(Msk.Item(1).Text, "mm/dd/yyyy") & "#," 'FECHA
                        Else 'ORACLE
                             VValores = "To_Date('" & Format(Msk.Item(0).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                             VValores = VValores & "To_Date('" & Format(Msk.Item(1).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                        End If
                        VValores = VValores & "'" & TxtTexto.Item(0).Text & "'," 'DOCUMENTO
                        VValores = VValores & "'" & CboEfecto.Text & "'," 'EFECTO
                        If GOrigenDeDatos = "AmaproAccess" Then
                             VValores = VValores & " #" & Format(Msk.Item(3).Text, "mm/dd/yyyy") & "#," 'FECHA PRODUCCION
                        Else 'ORACLE
                             VValores = VValores & " To_Date('" & Msk.Item(3).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA PRODUCCION
                        End If
                        VValores = VValores & "'" & TxtTexto.Item(1).Text & "'," 'LINEA
                        VValores = VValores & "'" & TxtTexto.Item(3).Text & "'," 'FICHA TECNICA
                        VValores = VValores & TxtTexto.Item(6).Text & "," 'TARIMA
                        VValores = VValores & Msk.Item(2).Text & "," 'CANTIDAD
                        VValores = VValores & "'" & TxtTexto.Item(4).Text & "'," 'OBSERVACIONES
                        VValores = VValores & "'" & TxtTexto.Item(2).Text & "'" 'USUARIO
                        
                        'INICIA UNA TRANSACCION
                       'SI ESTA GRABANDO UN REGISTRO NUEVO
                        Conexion.BeginTrans
                                Conexion.Execute "Insert Into AjustesProductoTerminado (" & VCampos & ") Values(" & VValores & ")"
                    
                    
                                   'SI ES CUALQUIER OTRO ERROR
                                    If Err <> 0 Then
                                       Conexion.RollbackTrans
                                       MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                       Exit Sub
                                    End If
                     
                                    'DISMINUYE EL SALDO
                                    If VTipo = "-" Then
                                               If GOrigenDeDatos = "AmaproAccess" Then
                                                    Conexion.Execute "update DetalleEntradasInventario set Saldo = Saldo - " & VCantidad & " where FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FechaProduccion = #" & Format(VFecha, "mm/dd/yyyy") & "#"
                                               Else 'ORACLE
                                                    Conexion.Execute "update DetalleEntradasInventario set Saldo = Saldo - " & VCantidad & " where FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FechaProduccion = TO_DATE('" & VFecha & "', 'dd/mm/yyyy')"
                                               End If
                                                   If Err <> 0 Then
                                                       Conexion.RollbackTrans
                                                       MsgBox "No Pudo Rebajar El Saldo " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                       Exit Sub
                                                   End If
                                    'AUMENTA EL SALDO
                                    ElseIf VTipo = "+" Then
                                               If GOrigenDeDatos = "AmaproAccess" Then
                                                    Conexion.Execute "update DetalleEntradasInventario set Saldo = Saldo + " & VCantidad & " where FichaTecnica = '" & VFichaTecnica & "' And Tarima = " & VTarima & " And Linea = '" & VLinea & "' And FechaProduccion = #" & Format(VFecha, "mm/dd/yyyy") & "#"
                                               Else 'ORACLE
                                                    Conexion.Execute "update DetalleEntradasInventario set Saldo = Saldo + " & VCantidad & " where UPPER(FichaTecnica) = '" & UCase(VFichaTecnica) & "' And Tarima = " & VTarima & " And UPPER(Linea) = '" & UCase(VLinea) & "' And FechaProduccion = TO_DATE('" & VFecha & "', 'dd/mm/yyyy')"
                                               End If
                                                   If Err <> 0 Then
                                                       Conexion.RollbackTrans
                                                       MsgBox "No Pudo Aumentar El Saldo " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                       Exit Sub
                                                   End If
                                    End If
                     
                         'FINALIZA LA TRANSACCION
                         Conexion.CommitTrans
                         
                         RAjustesProductoTerminado.Requery
                     
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
        RAjustesProductoTerminado.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RAjustesProductoTerminado.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RAjustesProductoTerminado.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RAjustesProductoTerminado.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RAjustesProductoTerminado.BOF Then
        RAjustesProductoTerminado.MoveFirst
    ElseIf RAjustesProductoTerminado.EOF Then
        RAjustesProductoTerminado.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0


End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        
        Set RAjustesProductoTerminado = New ADODB.Recordset
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RAjustesProductoTerminado, "Select * from AjustesProductoTerminado where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "#")
                Else
                    Call Abrir_Recordset(RAjustesProductoTerminado, "Select * from AjustesProductoTerminado where Fecha >= to_date('" & DtpFecIni.Value & "', 'dd/mm/yyyy') And Fecha <= to_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')")
                End If
            ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RAjustesProductoTerminado, "Select * from AjustesProductoTerminado where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And " & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And FichaTecnica Like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RAjustesProductoTerminado, "Select * from AjustesProductoTerminado where Fecha >= to_date('" & DtpFecIni.Value & "', 'dd/mm/yyyy') And Fecha <= to_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy') And UPPER(FichaTecnica) Like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RAjustesProductoTerminado, "Select * From AjustesProductoTerminado")
        End If
        
        'LLENA EL GRID
            Set DbGrid.DataSource = RAjustesProductoTerminado
    
        TabPuestos.Tab = 1
End Sub


Private Sub CmdSale_Click()
    FrameBuscar.Visible = False
End Sub

Private Sub DbGrid_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
        If RAjustesProductoTerminado.RecordCount > 0 Then
            RAjustesProductoTerminado.Sort = RAjustesProductoTerminado.Fields(ColIndex).Name
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            End If
        End If
End Sub






Private Sub DBGridBuscar_DblClick()
    If BLinea = True Then
        TxtTexto.Item(1).Text = DbGridBuscar.Columns(0)
        TxtTexto.Item(1).SetFocus
    ElseIf BFicha = True Then
        TxtTexto.Item(3).Text = DbGridBuscar.Columns(0)
        TxtTexto.Item(3).SetFocus
    End If
        FrameBuscar.Visible = False

End Sub

Private Sub DbGridBuscar_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If BLinea = True Then
            TxtTexto.Item(1).Text = DbGridBuscar.Columns(0)
            TxtTexto.Item(1).SetFocus
        ElseIf BFicha = True Then
            TxtTexto.Item(3).Text = DbGridBuscar.Columns(0)
            TxtTexto.Item(3).SetFocus
        End If
            FrameBuscar.Visible = False
End Sub

Private Sub Form_Load()
                
        Set RAjustesProductoTerminado = New ADODB.Recordset
        Call Abrir_Recordset(RAjustesProductoTerminado, "Select * From AjustesProductoTerminado")
        Set DbGrid.DataSource = RAjustesProductoTerminado
        Llena_Campos
                
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
        'CIERRA TODOS LOS RECORDSET Y CONTROLA SI HAY ALGUN ERROR PORQUE PUEDA SER QUE NO HABRAN ALGUN RECORDSET
        'Y A LA HORA DE CERRARLO GENERA ERROR
        RBuscaFichaTecnica.Close
        RBuscaTarima.Close
        RBuscaLinea.Close
        RMateriaPrima.Close
        RAjustesProductoTerminado.Close
        RBusqueda.Close
        
        Set RBuscaFichaTecnica = Nothing
        Set RBuscaTarima = Nothing
        Set RBuscaLinea = Nothing
        Set RMateriaPrima = Nothing
        Set RAjustesProductoTerminado = Nothing
        Set RBusqueda = Nothing
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
        Lbletiqueta.Caption = "Codigo"
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
                    'INICIALIZAMOS EL RECORDSET
                    Set RBusqueda = New ADODB.Recordset
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                        If BLinea = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where Descrip Like '*" & TxtBusqueda.Text & "*'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where UPPER(Descrip) Like '*" & UCase(TxtBusqueda.Text) & "*'")
                            End If
                        ElseIf BFicha = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from CorrelativosMateriaPrima Where Descripcion Like '*" & TxtBusqueda.Text & "*'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from CorrelativosMateriaPrima Where UPPER(Descripcion) Like '*" & UCase(TxtBusqueda.Text) & "*'")
                            End If
                        End If
                    'OPCION DE CODIGO
                    Else
                        If BLinea = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where Linea Like '*" & TxtBusqueda.Text & "*'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where UPPER(Linea) Like '*" & UCase(TxtBusqueda.Text) & "*'")
                            End If
                        ElseIf BFicha = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Esp_Tec Like '*" & TxtBusqueda.Text & "*'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Esp_Tec) Like '*" & UCase(TxtBusqueda.Text) & "*'")
                            End If
                        End If
                    End If
                            'LLENAMOS EL GRID CON EL RECORDSET
                            Set DbGridBuscar.DataSource = RMateriaPrima
                            DbGridBuscar.Refresh
                            DbGridBuscar.Columns(1).Width = "4000"
                        
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
        'LINEA
        If Index = 1 Then
            Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(1).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(1).Text) & "'")
            End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblLin.Caption = RBuscaLinea!Descrip
                Else
                    LblLin.Caption = ""
                End If
        'FICHA TECNICA
        ElseIf Index = 3 Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip, Esp_Tec From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(3).Text & "'")
            Else
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip, Esp_Tec From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(3).Text) & "'")
            End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblFic.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblFic.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
            If Index = 1 Then
                BLinea = True
                BFicha = False
                'INICIALIZAMOS EL RECORDSET
                Set RMateriaPrima = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RMateriaPrima, "Select Linea, Descrip from Lineas")
            ElseIf Index = 3 Then
                BLinea = False
                BFicha = True
                'INICIALIZAMOS EL RECORDSET
                Set RMateriaPrima = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RMateriaPrima, "Select Esp_Tec, Descrip from FichaTecnica")
            End If
            
            If Index = 1 Or Index = 3 Then
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DbGridBuscar.DataSource = RMateriaPrima
                DbGridBuscar.Refresh
                DbGridBuscar.Columns(1).Width = "4000"
                FrameBuscar.Visible = True
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
                BLinea = True
                BFicha = False
                'INICIALIZAMOS EL RECORDSET
                Set RMateriaPrima = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RMateriaPrima, "Select Linea, Descrip from Lineas")
            ElseIf Index = 3 Then
                BLinea = False
                BFicha = True
                'INICIALIZAMOS EL RECORDSET
                Set RMateriaPrima = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RMateriaPrima, "Select Esp_Tec, Descrip from FichaTecnica")
            End If
            
            If Index = 1 Or Index = 3 Then
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DbGridBuscar.DataSource = RMateriaPrima
                DbGridBuscar.Refresh
                DbGridBuscar.Columns(1).Width = "4000"
                FrameBuscar.Visible = True
                TxtBusqueda.SetFocus
            End If
        End If
End Sub

Private Sub Txttexto_LostFocus(Index As Integer)
        If Index = 6 Then
            If IsNumeric(TxtTexto.Item(6).Text) Then
                    'BUSCA TARIMA
                    Set RBuscaTarima = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaTarima, "Select Saldo From DetalleEntradasInventario Where FechaProduccion = #" & Format(Msk.Item(3).Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTexto.Item(6).Text & " And Linea = '" & TxtTexto.Item(1).Text & "' And FichaTecnica = '" & TxtTexto.Item(3).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaTarima, "Select Saldo From DetalleEntradasInventario Where FechaProduccion = TO_DATE('" & Msk.Item(3).Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTexto.Item(6).Text & " And Linea = '" & TxtTexto.Item(1).Text & "' And FichaTecnica = '" & TxtTexto.Item(3).Text & "'")
                    End If
                        If RBuscaTarima.RecordCount > 0 Then
                            Msk.Item(2).Text = RBuscaTarima!Saldo
                        Else
                            Msk.Item(2).Text = "0"
                        End If
            End If
        End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
            If RAjustesProductoTerminado.RecordCount > 0 Then
                Msk.Item(0).Text = RAjustesProductoTerminado!FechaOperacion
                Msk.Item(1).Text = RAjustesProductoTerminado!fecha
                If Not IsNull(RAjustesProductoTerminado!Documento) Then
                    TxtTexto.Item(0).Text = RAjustesProductoTerminado!Documento
                Else
                    TxtTexto.Item(0).Text = ""
                End If
                CboEfecto.Text = RAjustesProductoTerminado!Efecto
                Msk.Item(3).Text = RAjustesProductoTerminado!Fechaproduccion
                TxtTexto.Item(1).Text = RAjustesProductoTerminado!Linea
                TxtTexto.Item(3).Text = RAjustesProductoTerminado!FichaTecnica
                TxtTexto.Item(6).Text = RAjustesProductoTerminado!Tarima
                Msk.Item(2).Text = RAjustesProductoTerminado!Cantidad
                If Not IsNull(RAjustesProductoTerminado!Observaciones) Then
                    TxtTexto.Item(4).Text = RAjustesProductoTerminado!Observaciones
                Else
                    TxtTexto.Item(4).Text = ""
                End If
                TxtTexto.Item(2).Text = RAjustesProductoTerminado!Usuario
                
                If Err <> 0 Then
                    'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                End If
            Else
                Msk.Item(0).Text = ""
                Msk.Item(1).Text = ""
                TxtTexto.Item(0).Text = ""
                CboEfecto.Text = ""
                Msk.Item(3).Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(3).Text = ""
                TxtTexto.Item(6).Text = ""
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
                Msk.Item(3).Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(3).Text = ""
                TxtTexto.Item(6).Text = ""
                Msk.Item(2).Text = ""
                TxtTexto.Item(4).Text = ""
                TxtTexto.Item(2).Text = ""
End Sub


