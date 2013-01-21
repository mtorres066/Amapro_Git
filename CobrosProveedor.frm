VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CobrosProveedor 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Cobros A Proveedor"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8430
   Icon            =   "CobrosProveedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8430
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
      Height          =   6735
      Left            =   120
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DbGridBuscar 
         Height          =   5535
         Left            =   120
         TabIndex        =   40
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   9763
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
         Picture         =   "CobrosProveedor.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   41
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Txtbusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   39
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   5775
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   38
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   37
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
         TabIndex        =   42
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "CobrosProveedor.frx":3D6C
      Picture         =   "CobrosProveedor.frx":41AE
      Style           =   1  'Graphical
      TabIndex        =   66
      ToolTipText     =   "Ultimo Registro"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "CobrosProveedor.frx":46E0
      Picture         =   "CobrosProveedor.frx":4B22
      Style           =   1  'Graphical
      TabIndex        =   65
      ToolTipText     =   "Siguiente Registro"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "CobrosProveedor.frx":5054
      Picture         =   "CobrosProveedor.frx":5496
      Style           =   1  'Graphical
      TabIndex        =   64
      ToolTipText     =   "Registro Anterior"
      Top             =   6000
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "CobrosProveedor.frx":59C8
      Picture         =   "CobrosProveedor.frx":5E0A
      Style           =   1  'Graphical
      TabIndex        =   63
      ToolTipText     =   "Primer Registro"
      Top             =   6000
      Width           =   375
   End
   Begin TabDlg.SSTab TabPuestos 
      Height          =   5775
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   10186
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CobrosProveedor.frx":633C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePuestos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CobrosProveedor.frx":6656
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGrid"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CobrosProveedor.frx":6AA8
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "Label4"
      Tab(2).Control(2)=   "Label5"
      Tab(2).Control(3)=   "FrameOpciones"
      Tab(2).Control(4)=   "CmdBuscar(0)"
      Tab(2).Control(5)=   "CmdBuscar(1)"
      Tab(2).Control(6)=   "TxtBuscar"
      Tab(2).Control(7)=   "DtpFecIni"
      Tab(2).Control(8)=   "DtpFecFin"
      Tab(2).ControlCount=   9
      Begin MSDataGridLib.DataGrid DbGrid 
         Height          =   4935
         Left            =   -74880
         TabIndex        =   61
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   8705
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
         ColumnCount     =   14
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Proveedor"
            Caption         =   "Proveedor"
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
            DataField       =   "FichaTecnica"
            Caption         =   "FichaTecnica"
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
         BeginProperty Column04 
            DataField       =   "Boleta"
            Caption         =   "Boleta"
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
            DataField       =   "FechaRevision"
            Caption         =   "FechaRevision"
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
            DataField       =   "URevisadas"
            Caption         =   "URevisadas"
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
            DataField       =   "UNoConformes"
            Caption         =   "UNoConformes"
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
            DataField       =   "CostoxUnidad"
            Caption         =   "CostoxUnidad"
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
            DataField       =   "HorasHombre"
            Caption         =   "HorasHombre"
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
            DataField       =   "CostoxHora"
            Caption         =   "CostoxHora"
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
            DataField       =   "TazaCambio"
            Caption         =   "TazaCambio"
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
            DataField       =   "Serie"
            Caption         =   "Reclamo/Serie"
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
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column08 
            EndProperty
            BeginProperty Column09 
            EndProperty
            BeginProperty Column10 
            EndProperty
            BeginProperty Column11 
            EndProperty
            BeginProperty Column12 
            EndProperty
            BeginProperty Column13 
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -70080
         TabIndex        =   53
         Top             =   2640
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61669379
         CurrentDate     =   38146
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -70080
         TabIndex        =   52
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61669379
         CurrentDate     =   38146
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -70080
         TabIndex        =   23
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   3120
         Width           =   2085
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68880
         Picture         =   "CobrosProveedor.frx":6EFA
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68880
         Picture         =   "CobrosProveedor.frx":7204
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   3600
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
         Height          =   1695
         Left            =   -74880
         TabIndex        =   30
         Top             =   960
         Width           =   2565
         Begin VB.OptionButton OptBol 
            Caption         =   "No. Boleta"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            ToolTipText     =   " "
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton OptDef 
            Caption         =   "Fechas y Defecto"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   " "
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Fechas y Proveedor"
            Height          =   225
            Left            =   120
            TabIndex        =   21
            ToolTipText     =   " "
            Top             =   300
            Width           =   1815
         End
         Begin VB.OptionButton OptDescripcion 
            Caption         =   "Fechas y Ficha Tecnica"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            ToolTipText     =   " "
            Top             =   720
            Value           =   -1  'True
            Width           =   2175
         End
      End
      Begin VB.Frame FramePuestos 
         Caption         =   "Datos Del Cobro"
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
         Top             =   960
         Width           =   8175
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   5640
            MaxLength       =   20
            TabIndex        =   9
            Top             =   2520
            Width           =   2415
         End
         Begin VB.TextBox TxtTotal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
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
            Left            =   5640
            Locked          =   -1  'True
            TabIndex        =   60
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   3
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   1440
            Width           =   1455
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   0
            Left            =   1800
            TabIndex        =   5
            Top             =   2160
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Height          =   285
            Index           =   2
            Left            =   1800
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   4080
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   2
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   1
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   720
            Width           =   1455
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   6
            Top             =   2520
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
            Index           =   2
            Left            =   1800
            TabIndex        =   7
            Top             =   2880
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
            Left            =   1800
            TabIndex        =   8
            Top             =   3240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   4
            Left            =   5640
            TabIndex        =   10
            Top             =   2880
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   5
            Left            =   5640
            TabIndex        =   11
            Top             =   3240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   6
            Left            =   1800
            TabIndex        =   13
            Top             =   3720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   7
            Left            =   1800
            TabIndex        =   0
            Top             =   360
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
            Index           =   8
            Left            =   5640
            TabIndex        =   12
            Top             =   3600
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
            Caption         =   "Costo  Combustible"
            Height          =   195
            Index           =   12
            Left            =   4080
            TabIndex        =   67
            Top             =   3600
            Width           =   1350
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Reclamo/Serie"
            Height          =   195
            Index           =   11
            Left            =   4080
            TabIndex        =   62
            Top             =   2520
            Width           =   1065
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Costo Total"
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
            Left            =   4080
            TabIndex        =   59
            Top             =   3960
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   58
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Taza De Cambio"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   56
            Top             =   3720
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Mano De Obra"
            Height          =   195
            Index           =   9
            Left            =   4080
            TabIndex        =   50
            Top             =   2880
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Costo x Unidad"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   49
            Top             =   3240
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Unidades No Conformes"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   48
            Top             =   2880
            Width           =   1725
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Revision"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   47
            Top             =   2160
            Width           =   1110
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Boleta"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   46
            Top             =   1800
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Defecto"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   45
            Top             =   1440
            Width           =   570
         End
         Begin VB.Label LblDef 
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
            TabIndex        =   44
            Top             =   1440
            Width           =   4695
         End
         Begin VB.Label LblCur 
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
            TabIndex        =   43
            Top             =   1080
            Width           =   4695
         End
         Begin VB.Label LblEmp 
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
            TabIndex        =   35
            Top             =   720
            Width           =   4695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Costo x Hora"
            Height          =   195
            Index           =   3
            Left            =   4080
            TabIndex        =   34
            Top             =   3240
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Unidades Revisadas"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   2520
            Width           =   1470
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   3960
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Ficha Tecnica"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   1080
            Width           =   1020
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   -71040
         TabIndex        =   55
         Top             =   2640
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   -71040
         TabIndex        =   54
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   -72120
         TabIndex        =   31
         Top             =   3120
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6360
      MouseIcon       =   "CobrosProveedor.frx":7646
      Picture         =   "CobrosProveedor.frx":7A88
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5880
      Width           =   1125
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5280
      MouseIcon       =   "CobrosProveedor.frx":9AFA
      Picture         =   "CobrosProveedor.frx":9F3C
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "CobrosProveedor.frx":A46E
      Picture         =   "CobrosProveedor.frx":A8B0
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   3120
      MouseIcon       =   "CobrosProveedor.frx":ADE2
      Picture         =   "CobrosProveedor.frx":B224
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2040
      MouseIcon       =   "CobrosProveedor.frx":B756
      Picture         =   "CobrosProveedor.frx":BB98
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5880
      Width           =   1000
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   960
      MouseIcon       =   "CobrosProveedor.frx":C0CA
      Picture         =   "CobrosProveedor.frx":C50C
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5880
      Width           =   1000
   End
End
Attribute VB_Name = "CobrosProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim VMensaje As String
Dim buscar As String

Dim BProveedor As Boolean
Dim BFicha As Boolean
Dim BDefecto As Boolean
Dim BEditar As Boolean
Dim VLlave1 As String
Dim Vllave2 As String
Dim VTexto As String

Dim RBuscaProveedor As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaDefecto As New ADODB.Recordset
Dim RCobros As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VProveedor As String
Dim VCostoxHora As Currency
Dim VTaza As Currency


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
         FrameOpciones.Visible = False
         DbGrid.Visible = False
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
         FrameOpciones.Visible = True
         DbGrid.Visible = True
    End If
End Sub



Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        
            If Index = 0 Then
                    
                    Bandera = True
                    botones
                    Limpia_Campos
                    
                    Msk.Item(7).Text = Date
                    Msk.Item(7).SetFocus
                    TxtTexto.Item(0).Text = VProveedor
                    TxtTexto.Item(2).Text = GUsuario
                    Msk.Item(5).Text = VCostoxHora
                    Msk.Item(6).Text = VTaza
                    BEditar = False
            'EDITAR
            ElseIf Index = 1 Then
                    Bandera = True
                    botones
                    'GUARDA LA LLAVE
                    VLlave1 = TxtTexto.Item(5).Text
                    Vllave2 = Msk.Item(0).Text
                    Msk.Item(7).SetFocus
                    TxtTexto.Item(2).Text = GUsuario
                    BEditar = True
            'GRABAR
            ElseIf Index = 2 Then
                    Msk.Item(7).Text = Format(Msk.Item(7).Text, "dd/mm/yyyy")
                    Msk.Item(0).Text = Format(Msk.Item(0).Text, "dd/mm/yyyy")
                    'REVISA FECHA
                    If Not IsDate(Msk.Item(7).Text) Then
                        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(7).SetFocus
                        Exit Sub
                    End If
                    'REVISA FECHA DE REVISION
                    If Not IsDate(Msk.Item(0).Text) Then
                        MsgBox "Fecha De Revision Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(0).SetFocus
                        Exit Sub
                    End If
                    
                    'REVISA LAS UNIDADES REVISADAS
                    If Not IsNumeric(Msk.Item(1).Text) Then
                            MsgBox "Unidades Revisadas Incorrectas", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                    End If
                    
                    'REVISA LAS UNIDADES NO CONFORME
                    If Not IsNumeric(Msk.Item(2).Text) Then
                            MsgBox "Unidades No Conforme Incorrectas", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                    End If
                    'REVISA COSTO X UNIDAD
                    If Not IsNumeric(Msk.Item(3).Text) Then
                            MsgBox "Costo x Unidad", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                    End If
                    
                    'REVISA LAS HORAS HOMBRE
                    If Not IsNumeric(Msk.Item(4).Text) Then
                            MsgBox "Horas Mano De Obra Incorrectas", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                    End If
                    
                    'REVISA COSTO X HORA
                    If Not IsNumeric(Msk.Item(5).Text) Then
                            MsgBox "Costo x Hora Incorrectas", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                    End If
                    
                    VProveedor = TxtTexto.Item(0).Text
                    VCostoxHora = Msk.Item(5).Text
                    VTaza = Msk.Item(6).Text
                    
                     'AGREGAR
                    If BEditar = False Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "#" & Format(Msk.Item(7).Text, "mm/dd/yyyy") & "#, '" 'FECHA
                            Else 'ORACLE
                                 VTexto = "To_Date('" & Msk.Item(7).Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                            End If
                            VTexto = VTexto & TxtTexto.Item(0).Text & "', '" 'Descripcion
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', '" 'FICHA TECNICA
                            VTexto = VTexto & TxtTexto.Item(4).Text & "', '" 'DEFECTO
                            VTexto = VTexto & TxtTexto.Item(5).Text & "', " 'BOLETA
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "#" & Format(Msk.Item(0).Text, "mm/dd/yyyy") & "#, "  'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "To_Date('" & Msk.Item(0).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & Msk.Item(1).Text & ", " 'U REVISADAS
                            VTexto = VTexto & Msk.Item(2).Text & ", " 'U NO CONFORMES
                            VTexto = VTexto & Msk.Item(3).Text & ", " 'COSTO X UNIDAD
                            VTexto = VTexto & Msk.Item(4).Text & ", " 'HORAS HOMBRE
                            VTexto = VTexto & Msk.Item(5).Text & ", " 'COSTO X HORA
                            VTexto = VTexto & Msk.Item(6).Text & ", '" 'TAZA DE CAMBIO
                            VTexto = VTexto & TxtTexto(2).Text & "', '" 'USUARIO
                            VTexto = VTexto & TxtTexto(3).Text & "', " 'SERIE
                            VTexto = VTexto & Msk(8).Text 'COMBUSTIBLE
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into CobrosProveedor Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "Fecha = #" & Format(Msk.Item(7).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = "Fecha = To_Date('" & Msk.Item(7).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "Proveedor = '" & TxtTexto.Item(0).Text & "', " 'Descripcion
                            VTexto = VTexto & "FichaTecnica = '" & TxtTexto.Item(1).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & "Defecto = '" & TxtTexto.Item(4).Text & "', " 'DEFECTO
                            VTexto = VTexto & "Boleta = '" & TxtTexto.Item(5).Text & "', " 'BOLETA
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "FechaRevision = #" & Format(Msk.Item(0).Text, "mm/dd/yyyy") & "#, "  'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "FechaRevision = To_Date('" & Msk.Item(0).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "URevisadas = " & Msk.Item(1).Text & ", " 'U REVISADAS
                            VTexto = VTexto & "UNoConformes = " & Msk.Item(2).Text & ", " 'U NO CONFORMES
                            VTexto = VTexto & "CostoxUnidad = " & Msk.Item(3).Text & ", " 'COSTO X UNIDAD
                            VTexto = VTexto & "HorasHombre = " & Msk.Item(4).Text & ", " 'HORAS HOMBRE
                            VTexto = VTexto & "CostoxHora = " & Msk.Item(5).Text & ", " 'COSTO X HORA
                            VTexto = VTexto & "TazaCambio = " & Msk.Item(6).Text & ", " 'TAZA DE CAMBIO
                            VTexto = VTexto & "Usuario = '" & TxtTexto(2).Text & "', " 'USUARIO
                            VTexto = VTexto & "Serie = '" & TxtTexto(3).Text & "', " 'SERIE
                            VTexto = VTexto & "Combustible = " & Msk(8).Text 'COMBUSTIBLE
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & " Where Boleta = '" & VLlave1 & "' And FechaRevision = #" & Format(Vllave2, "mm/dd/yyyy") & "#"
                            Else 'OACLE
                                VTexto = VTexto & " Where UPPER(Boleta) = '" & UCase(VLlave1) & "' And FechaRevision = To_Date('" & Vllave2 & "', 'dd/mm/yyyy')"
                            End If
                            
                            Conexion.Execute "UPDATE CobrosProveedor SET " & VTexto
                    End If
                     
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Boleta y Fecha De Revision Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Boleta y Fecha De Revision Ya Existe", vbOKOnly + vbInformation, "Informacion"
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
                        RCobros.Requery
                        RCobros.MoveLast
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
                        RCobros.Delete
                        
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
                        RCobros.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RCobros.MoveLast
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
        RCobros.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RCobros.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RCobros.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RCobros.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RCobros.BOF Then
        RCobros.MoveFirst
    ElseIf RCobros.EOF Then
        RCobros.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
    If IsNumeric(Msk.Item(2).Text) And IsNumeric(Msk.Item(3).Text) And IsNumeric(Msk.Item(4).Text) And IsNumeric(Msk.Item(5).Text) And IsNumeric(Msk.Item(6).Text) And IsNumeric(Msk.Item(8).Text) Then
        If Msk.Item(6).Text > 0 Then
            TxtTotal.Text = (((Msk.Item(2) * Msk.Item(3)) + (Msk.Item(4) * Msk.Item(5)) + Msk.Item(8).Text) / Msk.Item(6).Text)
            TxtTotal.Text = Format(TxtTotal.Text, "#,###,##0.00")
        Else
            TxtTotal.Text = 0
        End If
    Else
        TxtTotal.Text = 0
    End If

    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
        
        Set RCobros = New ADODB.Recordset
        
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RCobros, "Select * from CobrosProveedor where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaRevision <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Proveedor Like '%" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RCobros, "Select * from CobrosProveedor where Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And FechaRevision <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Proveedor) Like '%" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RCobros, "Select * from CobrosProveedor where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaRevision <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And FichaTecnica Like '%" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RCobros, "Select * from CobrosProveedor where Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And FechaRevision <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(FichaTecnica) Like '%" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptDef.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RCobros, "Select * from CobrosProveedor where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaRevision <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Defecto Like '%" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RCobros, "Select * from CobrosProveedor where Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And FechaRevision <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Defecto) Like '%" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptBol.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RCobros, "Select * from CobrosProveedor where Boleta Like '%" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RCobros, "Select * from CobrosProveedor where UPPER(Boleta) Like '%" & UCase(TxtBuscar.Text) & "%'")
                End If
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RCobros, "Select * From CobrosProveedor")
        End If
                Set DbGrid.DataSource = RCobros
                TabPuestos.Tab = 1
End Sub


Private Sub CmdSale_Click()
    FrameBuscar.Visible = False
End Sub



Private Sub DbGrid_HeadClick(ByVal ColIndex As Integer)
        RCobros.Sort = RCobros.Fields(ColIndex).Name
End Sub


Private Sub DBGridBuscar_DblClick()
    If BProveedor = True Then
        TxtTexto.Item(0).Text = DbGridBuscar.Columns(0)
        TxtTexto.Item(0).SetFocus
    ElseIf BFicha = True Then
        TxtTexto.Item(1).Text = DbGridBuscar.Columns(0)
        TxtTexto.Item(1).SetFocus
    ElseIf BDefecto = True Then
        TxtTexto.Item(4).Text = DbGridBuscar.Columns(0)
        TxtTexto.Item(4).SetFocus
    End If
        FrameBuscar.Visible = False

End Sub

Private Sub DbGridBuscar_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
    If BProveedor = True Then
        TxtTexto.Item(0).Text = DbGridBuscar.Columns(0)
        TxtTexto.Item(0).SetFocus
    ElseIf BFicha = True Then
        TxtTexto.Item(1).Text = DbGridBuscar.Columns(0)
        TxtTexto.Item(1).SetFocus
    ElseIf BDefecto = True Then
        TxtTexto.Item(4).Text = DbGridBuscar.Columns(0)
        TxtTexto.Item(4).SetFocus
    End If
        FrameBuscar.Visible = False
End Sub

Private Sub Form_Load()
        Set RCobros = New ADODB.Recordset
        Call Abrir_Recordset(RCobros, "Select * From CobrosProveedor")
        Set DbGrid.DataSource = RCobros
        Llena_Campos
        
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date
End Sub

Private Sub Msk_Change(Index As Integer)
    If IsNumeric(Msk.Item(2).Text) And IsNumeric(Msk.Item(3).Text) And IsNumeric(Msk.Item(4).Text) And IsNumeric(Msk.Item(5).Text) And IsNumeric(Msk.Item(6).Text) And IsNumeric(Msk.Item(8).Text) Then
        If Msk.Item(6).Text > 0 Then
            TxtTotal.Text = (((Msk.Item(2) * Msk.Item(3)) + (Msk.Item(4) * Msk.Item(5)) + Msk.Item(8)) / Msk.Item(6).Text)
            TxtTotal.Text = Format(TxtTotal.Text, "#,###,##0.00")
        Else
            TxtTotal.Text = 0
        End If
    Else
        TxtTotal.Text = 0
    End If
End Sub

Private Sub Msk_GotFocus(Index As Integer)
        Msk.Item(Index).SelStart = 0
        Msk.Item(Index).SelLength = Len(Msk.Item(Index).Text)
End Sub

Private Sub Msk_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub Msk_LostFocus(Index As Integer)
    If IsNumeric(Msk.Item(2).Text) And IsNumeric(Msk.Item(3).Text) And IsNumeric(Msk.Item(4).Text) And IsNumeric(Msk.Item(5).Text) And IsNumeric(Msk.Item(6).Text) And IsNumeric(Msk.Item(8).Text) Then
        If Msk.Item(6).Text > 0 Then
            TxtTotal.Text = (((Msk.Item(2) * Msk.Item(3)) + (Msk.Item(4) * Msk.Item(5)) + Msk.Item(8).Text) / Msk.Item(6).Text)
            TxtTotal.Text = Format(TxtTotal.Text, "#,###,##0.00")
        Else
            TxtTotal.Text = 0
        End If
    Else
        TxtTotal.Text = 0
    End If
End Sub

Private Sub OptBol_Click()
        Label4.Visible = False
        Label5.Visible = False
        DtpFecIni.Visible = False
        DTPFecFin.Visible = False
        Lbletiqueta.Caption = "No. Boleta"
        TxtBuscar.SetFocus
End Sub

Private Sub OptCodigo_Click()
        Label4.Visible = True
        Label5.Visible = True
        DtpFecIni.Visible = True
        DTPFecFin.Visible = True
        Lbletiqueta.Caption = "Proveedor"
        TxtBuscar.SetFocus
End Sub

Private Sub OptDef_Click()
        Label4.Visible = True
        Label5.Visible = True
        DtpFecIni.Visible = True
        DTPFecFin.Visible = True
        Lbletiqueta.Caption = "Defecto"
        TxtBuscar.SetFocus
End Sub

Private Sub OptDescripcion_Click()
        Label4.Visible = True
        Label5.Visible = True
        DtpFecIni.Visible = True
        DTPFecFin.Visible = True
        Lbletiqueta.Caption = "Ficha Tecnica"
        TxtBuscar.SetFocus
End Sub

Private Sub TabPuestos_Click(PreviousTab As Integer)
        If TabPuestos.Tab = 0 Then
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

Private Sub Txtbusqueda_Change()
        Set RBusqueda = New ADODB.Recordset
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                        If BProveedor = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BFicha = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Descrip Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BDefecto = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip from Defectos Where Descrip Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip from Defectos Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        End If
                    'OPCION DE CODIGO
                    Else
                        If BProveedor = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where CodigoProveedor Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(CodigoProveedor) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BFicha = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Esp_Tec Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BDefecto = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip from Defectos Where Defecto Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip from Defectos Where UPPER(Defecto) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        End If
                    End If
                            
                            Set DbGridBuscar.DataSource = RBusqueda
                            DbGridBuscar.Columns(1).Width = "5000"
                            

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
        'PROVEEDORES
        If Index = 0 Then
            Set RBuscaProveedor = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtTexto.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                End If
                If RBuscaProveedor.RecordCount > 0 Then
                    LblEmp.Caption = RBuscaProveedor!Descripcion
                Else
                    LblEmp.Caption = ""
                End If
        'FICHA TECNICA
        ElseIf Index = 1 Then
            Set RBuscaFicha = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(1).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(1).Text) & "'")
                End If
                If RBuscaFicha.RecordCount > 0 Then
                    LblCur.Caption = RBuscaFicha!Descrip
                Else
                    LblCur.Caption = ""
                End If
        'DEFECTOS
        ElseIf Index = 4 Then
            Set RBuscaDefecto = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaDefecto, "Select Descrip From Defectos Where Defecto = '" & TxtTexto.Item(4).Text & "'")
                Else ' ORACLE
                    Call Abrir_Recordset(RBuscaDefecto, "Select Descrip From Defectos Where UPPER(Defecto) = '" & UCase(TxtTexto.Item(4).Text) & "'")
                End If
                If RBuscaDefecto.RecordCount > 0 Then
                    LblDef.Caption = RBuscaDefecto!Descrip
                Else
                    LblDef.Caption = ""
                End If
        
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
            
            If Index = 0 Or Index = 1 Or Index = 4 Then
                Set RBusqueda = New ADODB.Recordset
            End If
            
            If Index = 0 Then
                BProveedor = True
                BFicha = False
                BDefecto = False
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores")
            ElseIf Index = 1 Then
                BProveedor = False
                BFicha = True
                BDefecto = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            ElseIf Index = 4 Then
                BProveedor = False
                BFicha = False
                BDefecto = True
                Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip from Defectos")
            End If
            
            If Index = 0 Or Index = 1 Or Index = 4 Then
                Set DbGridBuscar.DataSource = RBusqueda
                FrameBuscar.Visible = True
                TxtBusqueda.SetFocus
                DbGridBuscar.Columns(1).Width = "4000"
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
            If Index = 0 Or Index = 1 Or Index = 4 Then
                Set RBusqueda = New ADODB.Recordset
            End If
            
            If Index = 0 Then
                BProveedor = True
                BFicha = False
                BDefecto = False
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores")
            ElseIf Index = 1 Then
                BProveedor = False
                BFicha = True
                BDefecto = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            ElseIf Index = 4 Then
                BProveedor = False
                BFicha = False
                BDefecto = True
                Call Abrir_Recordset(RBusqueda, "Select Defecto, Descrip from Defectos")
            End If
            
            If Index = 0 Or Index = 1 Or Index = 4 Then
                Set DbGridBuscar.DataSource = RBusqueda
                FrameBuscar.Visible = True
                TxtBusqueda.SetFocus
                DbGridBuscar.Columns(1).Width = "4000"
            End If
        End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
        'FECHA
            If IsNull(RCobros!fecha) Then
                Msk.Item(7).Text = ""
            Else
                Msk.Item(7).Text = RCobros!fecha
            End If
        'Descripcion
            If IsNull(RCobros!Proveedor) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = RCobros!Proveedor
            End If
        'FICHATECNICA
            If IsNull(RCobros!FichaTecnica) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = RCobros!FichaTecnica
            End If
        'DEFECTO
            If IsNull(RCobros!Defecto) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = RCobros!Defecto
            End If
        'BOLETA
            If IsNull(RCobros!Boleta) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = RCobros!Boleta
            End If
        'Fecha Revision
            If IsNull(RCobros!FechaRevision) Then
                Msk.Item(0).Text = ""
            Else
                Msk.Item(0).Text = RCobros!FechaRevision
            End If
        'URevisadas
            If IsNull(RCobros!URevisadas) Then
                Msk.Item(1).Text = ""
            Else
                Msk.Item(1).Text = RCobros!URevisadas
            End If
        'UNoConformes
            If IsNull(RCobros!UNoConformes) Then
                Msk.Item(2).Text = ""
            Else
                Msk.Item(2).Text = RCobros!UNoConformes
            End If
        'CostoxUnidad
            If IsNull(RCobros!CostoxUnidad) Then
                Msk.Item(3).Text = ""
            Else
                Msk.Item(3).Text = RCobros!CostoxUnidad
            End If
        'HorasHombre
            If IsNull(RCobros!HorasHombre) Then
                Msk.Item(4).Text = ""
            Else
                Msk.Item(4).Text = RCobros!HorasHombre
            End If
        'CostoxHora
            If IsNull(RCobros!CostoxHora) Then
                Msk.Item(5).Text = ""
            Else
                Msk.Item(5).Text = RCobros!CostoxHora
            End If
        'TazaCambio
            If IsNull(RCobros!TazaCambio) Then
                Msk.Item(6).Text = ""
            Else
                Msk.Item(6).Text = RCobros!TazaCambio
            End If
        'USUARIO
            If IsNull(RCobros!Usuario) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RCobros!Usuario
            End If
        'SERIE
            If IsNull(RCobros!Serie) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = RCobros!Serie
            End If
            'COMBUSTIBLE
            If IsNull(RCobros!Combustible) Then
                Msk.Item(8).Text = ""
            Else
                Msk.Item(8).Text = RCobros!Combustible
            End If
            
        If Err <> 0 Then
            'MsgBox Err.Description
        End If

End Sub

Public Sub Limpia_Campos()
        
        Msk.Item(7).Text = ""
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(4).Text = ""
        TxtTexto.Item(5).Text = ""
        Msk.Item(0).Text = ""
        Msk.Item(1).Text = 0
        Msk.Item(2).Text = 0
        Msk.Item(3).Text = 0
        Msk.Item(4).Text = 0
        Msk.Item(5).Text = 0
        Msk.Item(6).Text = 0
        Msk.Item(8).Text = 0
        TxtTexto.Item(2).Text = ""
        TxtTexto.Item(3).Text = ""
        
End Sub


