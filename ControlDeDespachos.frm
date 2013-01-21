VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ControlDeDespachos 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Productos En Transito"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "ControlDeDespachos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
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
      Height          =   6135
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4815
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   8493
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
         Picture         =   "ControlDeDespachos.frx":1CFA
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Txtbusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   34
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   5775
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   32
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
         TabIndex        =   37
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2040
      MouseIcon       =   "ControlDeDespachos.frx":3D6C
      Picture         =   "ControlDeDespachos.frx":41AE
      Style           =   1  'Graphical
      TabIndex        =   57
      Top             =   5280
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "ControlDeDespachos.frx":46E0
      Picture         =   "ControlDeDespachos.frx":4B22
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Primer Registro"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "ControlDeDespachos.frx":5054
      Picture         =   "ControlDeDespachos.frx":5496
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Registro Anterior"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "ControlDeDespachos.frx":59C8
      Picture         =   "ControlDeDespachos.frx":5E0A
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Siguiente Registro"
      Top             =   5400
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "ControlDeDespachos.frx":633C
      Picture         =   "ControlDeDespachos.frx":677E
      Style           =   1  'Graphical
      TabIndex        =   53
      ToolTipText     =   "Ultimo Registro"
      Top             =   5400
      Width           =   375
   End
   Begin TabDlg.SSTab TabPuestos 
      Height          =   5175
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   9128
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "ControlDeDespachos.frx":6CB0
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FramePuestos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "ControlDeDespachos.frx":6FCA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGrid"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "ControlDeDespachos.frx":741C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DtpFecFin"
      Tab(2).Control(1)=   "DtpFecIni"
      Tab(2).Control(2)=   "TxtBuscar"
      Tab(2).Control(3)=   "CmdBuscar(1)"
      Tab(2).Control(4)=   "CmdBuscar(0)"
      Tab(2).Control(5)=   "FrameOpciones"
      Tab(2).Control(6)=   "Label5"
      Tab(2).Control(7)=   "Label4"
      Tab(2).Control(8)=   "Lbletiqueta"
      Tab(2).ControlCount=   9
      Begin MSDataGridLib.DataGrid DbGrid 
         Height          =   4335
         Left            =   -74880
         TabIndex        =   52
         Top             =   720
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
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -72360
         TabIndex        =   45
         Top             =   3600
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51445763
         CurrentDate     =   38146
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -72360
         TabIndex        =   44
         Top             =   3240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51445763
         CurrentDate     =   38146
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -72360
         TabIndex        =   19
         ToolTipText     =   "Digite los datos para hacer la busqueda"
         Top             =   3960
         Width           =   2085
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -68880
         Picture         =   "ControlDeDespachos.frx":786E
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -68880
         Picture         =   "ControlDeDespachos.frx":7B78
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2760
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
         Height          =   2055
         Left            =   -74880
         TabIndex        =   26
         Top             =   960
         Width           =   3165
         Begin VB.OptionButton OptFac 
            Caption         =   "Factura"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   " "
            Top             =   1800
            Width           =   2895
         End
         Begin VB.OptionButton OptBol 
            Caption         =   "Fechas Arribo y Codigo"
            Height          =   195
            Left            =   120
            TabIndex        =   48
            ToolTipText     =   " "
            Top             =   1440
            Width           =   2895
         End
         Begin VB.OptionButton OptDef 
            Caption         =   "Fechas Arribo y Proveedor"
            Height          =   195
            Left            =   120
            TabIndex        =   43
            ToolTipText     =   " "
            Top             =   1080
            Width           =   2895
         End
         Begin VB.OptionButton OptCodigo 
            Caption         =   "Fechas Despacho y Proveedor"
            Height          =   225
            Left            =   120
            TabIndex        =   17
            ToolTipText     =   " "
            Top             =   300
            Width           =   2655
         End
         Begin VB.OptionButton OptDescripcion 
            Caption         =   "Fechas Despacho y Codigo"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            ToolTipText     =   " "
            Top             =   720
            Value           =   -1  'True
            Width           =   2895
         End
      End
      Begin VB.Frame FramePuestos 
         Caption         =   "Datos Del Despacho"
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
         Height          =   4335
         Left            =   120
         TabIndex        =   23
         Top             =   720
         Width           =   8175
         Begin VB.TextBox TxtDoc 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6600
            MaxLength       =   15
            TabIndex        =   58
            Top             =   240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.CheckBox ChkRec 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            Caption         =   "Factura Recibida ?"
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
            TabIndex        =   7
            Top             =   2520
            Width           =   2055
         End
         Begin VB.CheckBox ChkMultiplica 
            BackColor       =   &H80000004&
            Caption         =   "Laminas x Unidades"
            Height          =   255
            Left            =   1800
            TabIndex        =   4
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   10
            Top             =   3600
            Width           =   4575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1800
            MaxLength       =   30
            TabIndex        =   9
            Top             =   3240
            Width           =   4575
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   6
            Top             =   2520
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
            Format          =   "#,###,##0.00"
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
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   2
            ToolTipText     =   "signo '+' o doble click para ayuda"
            Top             =   1080
            Width           =   1455
         End
         Begin MSMask.MaskEdBox Msk 
            Height          =   285
            Index           =   1
            Left            =   1800
            TabIndex        =   8
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
            Left            =   1800
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
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Codigo"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   50
            Top             =   1440
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Despacho"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   49
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Transportista"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   42
            Top             =   3600
            Width           =   915
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Piloto"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   41
            Top             =   3240
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "No. Factura"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   40
            Top             =   2520
            Width           =   840
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   39
            Top             =   2160
            Width           =   630
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
            TabIndex        =   38
            Top             =   1440
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
            TabIndex        =   30
            Top             =   1080
            Width           =   4695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto En Dolares $"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   29
            Top             =   2880
            Width           =   1410
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   28
            Top             =   3960
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Arribo"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   720
            Width           =   900
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   -73320
         TabIndex        =   47
         Top             =   3600
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   -73320
         TabIndex        =   46
         Top             =   3240
         Width           =   900
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
         Left            =   -74400
         TabIndex        =   27
         Top             =   3960
         Width           =   1935
      End
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   6360
      MouseIcon       =   "ControlDeDespachos.frx":7FBA
      Picture         =   "ControlDeDespachos.frx":83FC
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   5280
      MouseIcon       =   "ControlDeDespachos.frx":A46E
      Picture         =   "ControlDeDespachos.frx":A8B0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5280
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   4200
      MouseIcon       =   "ControlDeDespachos.frx":ADE2
      Picture         =   "ControlDeDespachos.frx":B224
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5280
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   3120
      MouseIcon       =   "ControlDeDespachos.frx":B756
      Picture         =   "ControlDeDespachos.frx":BB98
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5280
      Width           =   1020
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   960
      MouseIcon       =   "ControlDeDespachos.frx":C0CA
      Picture         =   "ControlDeDespachos.frx":C50C
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5280
      Width           =   1020
   End
End
Attribute VB_Name = "ControlDeDespachos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BProveedor As Boolean
Dim BFicha As Boolean
Dim BEditar As Boolean
Dim Vllave As String
Dim VTexto As String

Dim RBuscaProveedor As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaCuerpos As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RTransito As New ADODB.Recordset
Dim RBuscaSigDoc As New ADODB.Recordset

Dim VProveedor As String
Dim VFechaArribo As Date
Dim VCantidad As Single
Dim VFactura As String
Dim VMonto As Currency
Dim VPiloto As String
Dim VTransportista As String


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
         
         FrameOpciones.Visible = True
         DbGrid.Visible = True
    End If
End Sub




Private Sub ChkMultiplica_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub ChkRec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
    
End Sub

Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        
            If Index = 0 Then
                    
                    'BUSCA EL DOCUMENTO MAXIMO Y LE ASIGNA 1
                    Set RBuscaSigDoc = New ADODB.Recordset
                        Call Abrir_Recordset(RBuscaSigDoc, "Select Max(Documento) from ProductoEnTransito")
                        If RBuscaSigDoc.RecordCount > 0 Then
                            If IsNull(RBuscaSigDoc(0)) Then
                                TxtDoc.Text = "1"
                            Else
                                TxtDoc.Text = RBuscaSigDoc(0) + 1
                            End If
                        End If
                        
                    Bandera = True
                    botones
                    Limpia_Campos
                    
                    Msk.Item(7).Text = Date
                    Msk.Item(7).SetFocus
                    
                    TxtTexto.Item(0).Text = VProveedor
                    Msk.Item(8).Text = Format(VFechaArribo, "dd/mm/yyyy")
                    Msk.Item(0).Text = VCantidad
                    TxtTexto.Item(5).Text = VFactura
                    Msk.Item(1).Text = VMonto
                    TxtTexto.Item(3).Text = VPiloto
                    TxtTexto.Item(4).Text = VTransportista
                    TxtTexto.Item(2).Text = GUsuario
                    BEditar = False
                    ChkRec.Value = 0
            'EDITAR
            ElseIf Index = 1 Then
                    Bandera = True
                    botones
                    'GUARDA LA LLAVE
                    Vllave = TxtDoc.Text
                    Msk.Item(7).SetFocus
                    TxtTexto.Item(2).Text = GUsuario
                    BEditar = True
            
            'GRABAR
            ElseIf Index = 2 Then
                    
                    Msk.Item(7).Text = Format(Msk.Item(7), "dd/mm/yyyy")
                    Msk.Item(8).Text = Format(Msk.Item(8), "dd/mm/yyyy")
                    
                    'REVISA FECHA DE DESPACHO
                    If Not IsDate(Msk.Item(7).Text) Then
                        MsgBox "Fecha De Despacho Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(7).SetFocus
                        Exit Sub
                    End If
                    'REVISA FECHA DE ARRIBO
                    If Not IsDate(Msk.Item(8).Text) Then
                        MsgBox "Fecha De Arribo Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Msk.Item(8).SetFocus
                        Exit Sub
                    End If
                    
                    'MONTO EN DOLARES
                    If Not IsNumeric(Msk.Item(1).Text) Then
                            MsgBox "Monto En Dolares Incorrecto", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                    End If
                    
                    'GRABA VARIABLES
                    VProveedor = TxtTexto.Item(0).Text
                    VFechaArribo = Msk.Item(8).Text
                    VCantidad = Msk.Item(0).Text
                    VFactura = TxtTexto.Item(5).Text
                    VMonto = Msk.Item(1).Text
                    VPiloto = TxtTexto.Item(3).Text
                    VTransportista = TxtTexto.Item(4).Text
                    
                        If BEditar = False Then
                    
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "#" & Format(Msk.Item(7).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = "To_Date('" & Msk.Item(7).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "#" & Format(Msk.Item(8).Text, "mm/dd/yyyy") & "#, '" 'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "To_Date('" & Msk.Item(8).Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                            End If
                            VTexto = VTexto & TxtTexto.Item(0).Text & "', '" 'PROVEEDOR
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & Msk.Item(0).Text & ", '" 'CANTIDAD
                            VTexto = VTexto & TxtTexto.Item(5).Text & "', " 'FACTURA
                            VTexto = VTexto & Msk.Item(1).Text & ", '" 'MONTO DOLARES
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', '" 'PILOTO
                            VTexto = VTexto & TxtTexto.Item(4).Text & "', '" 'TRANSPORTISTA
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', " 'USUARIO
                            If ChkRec.Value = "1" Then
                                VTexto = VTexto & "-1, " 'RECIBIDA
                            Else
                                VTexto = VTexto & "0, "  'NO RECIBIDA
                            End If
                            VTexto = VTexto & TxtDoc.Text 'DOCUMENTO
                            
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into ProductoEnTransito Values(" & VTexto & ")"
                            
                    'EDITAR
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = "FechaDespacho = #" & Format(Msk.Item(7).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = "FechaDespacho = To_Date('" & Msk.Item(7).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            If GOrigenDeDatos = "AmaproAccess" Then
                                 VTexto = VTexto & "FechaArribo = #" & Format(Msk.Item(8).Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 VTexto = VTexto & "FechaArribo = To_Date('" & Msk.Item(8).Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "Proveedor = '" & TxtTexto.Item(0).Text & "', " 'PROVEEDOR
                            VTexto = VTexto & "Codigo = '" & TxtTexto.Item(1).Text & "', " 'FICHA TECNICA
                            VTexto = VTexto & "Cantidad = " & Msk.Item(0).Text & ", " 'CANTIDAD
                            VTexto = VTexto & "Factura = '" & TxtTexto.Item(5).Text & "', " 'FACTURA
                            VTexto = VTexto & "MontoDolares = " & Msk.Item(1).Text & ", " 'MONTO DOLARES
                            VTexto = VTexto & "Piloto = '" & TxtTexto.Item(3).Text & "', " 'PILOTO
                            VTexto = VTexto & "Transportista = '" & TxtTexto.Item(4).Text & "', " 'TRANSPORTISTA
                            VTexto = VTexto & "Usuario = '" & TxtTexto.Item(2).Text & "', " 'USUARIO
                            If ChkRec.Value = "1" Then
                                VTexto = VTexto & "Recibida = -1" 'RECIBIDA
                            Else
                                VTexto = VTexto & "Recibida = 0"  'NO RECIBIDA
                            End If
                            VTexto = VTexto & " Where Documento = " & Vllave
                            
                            Conexion.Execute "UPDATE ProductoEnTransito SET " & VTexto
                    End If
                    
                    
                     
                            'SI SE DUPLICA LA LLAVE
                             If GOrigenDeDatos = "AmaproAccess" Then
                              'SI ES CUALQUIER OTRO ERROR
                                If Err <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                    Err.Clear
                                    Exit Sub
                                End If
                            Else 'ORACLE
                                'I ES CUALQUIER OTRO ERROR
                                If Err <> 0 Then
                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                    Err.Clear
                                    Exit Sub
                                End If
                            End If
                                         
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
                        
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RTransito.Requery
                        RTransito.MoveLast
                        Llena_Campos
            'CANCELAR
            ElseIf Index = 3 Then
                    'CANCELA LOS CAMBIOS Y DEJA LOS DATOS COMO ESTABAN
                    Bandera = False
                    botones
                    Llena_Campos
            'BORRAR
            ElseIf Index = 4 Then
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RTransito.Delete
                        
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
                        RTransito.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RTransito.MoveLast
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
        RTransito.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RTransito.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RTransito.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RTransito.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RTransito.BOF Then
        RTransito.MoveFirst
    ElseIf RTransito.EOF Then
        RTransito.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0


End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    Set RTransito = New ADODB.Recordset
    
        'SELECCIONAR DATOS
        If Index = 0 Then
            If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where FechaDespacho >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaDespacho <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Proveedor Like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where FechaDespacho >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And FechaDespacho <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Proveedor) Like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where FechaDespacho >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaDespacho <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Codigo Like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where FechaDespacho >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And FechaDespacho <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Codigo) Like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptDef.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where FechaArribo >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaArribo <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Proveedor Like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where FechaArribo >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And FechaArribo <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Proveedor) Like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptBol.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where FechaArribo >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And FechaArribo <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Codigo Like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where FechaArribo >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And FechaArribo <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Codigo) Like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            ElseIf OptFac.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where Factura Like '" & TxtBuscar.Text & "%'")
                Else 'ORACLE
                    Call Abrir_Recordset(RTransito, "Select * from ProductoEnTransito where UPPER(Factura) Like '" & UCase(TxtBuscar.Text) & "%'")
                End If
            End If
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                Call Abrir_Recordset(RTransito, "Select * From ProductoEnTransito")
        End If
    
        Set DbGrid.DataSource = RTransito
        TabPuestos.Tab = 1
End Sub


Private Sub CmdSale_Click()
    FrameBuscar.Visible = False
End Sub



Private Sub DbGrid_HeadClick(ByVal ColIndex As Integer)
            RTransito.Sort = RTransito.Fields(ColIndex).Name
End Sub


Private Sub DBGridBusqueda_DblClick()
    If BProveedor = True Then
        TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(0).SetFocus
    ElseIf BFicha = True Then
        TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(1).SetFocus
    End If
        FrameBuscar.Visible = False

End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If BProveedor = True Then
        TxtTexto.Item(0).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(0).SetFocus
    ElseIf BFicha = True Then
        TxtTexto.Item(1).Text = DBGridBusqueda.Columns(0)
        TxtTexto.Item(1).SetFocus
    End If
        FrameBuscar.Visible = False
End Sub

Private Sub Form_Load()
        Set RTransito = New ADODB.Recordset
        Call Abrir_Recordset(RTransito, "Select * From ProductoEnTransito")
            Set DbGrid.DataSource = RTransito
            Llena_Campos
        
        
        DtpFecIni.Value = Date
        DTPFecFin.Value = Date

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
    
    If Index = 0 Then
            'SI DESEA MULTIPLICAR LAS LAMINAS QUE VIENEN, BUSCA POR EL CODIGO CUANTAS LAMINA TIENE CADA CODIGO Y LAS MULTIPLICA
            If ChkMultiplica.Value = 1 Then
                If IsNumeric(Msk.Item(0).Text) Then
                    Set RBuscaCuerpos = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaCuerpos, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(1).Text & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaCuerpos, "Select UnidadesxLamina From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(1).Text) & "'")
                        End If
                        If RBuscaCuerpos.RecordCount > 0 Then
                            Msk.Item(0).Text = Msk.Item(0).Text * RBuscaCuerpos!UnidadesxLamina
                        End If
                End If
            End If
    End If

End Sub

Private Sub OptBol_Click()
        Label4.Visible = False
        Label5.Visible = False
        DtpFecIni.Visible = False
        DTPFecFin.Visible = False
        Lbletiqueta.Caption = "Codigo"
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
        Lbletiqueta.Caption = "Proveedor"
        TxtBuscar.SetFocus
End Sub

Private Sub OptDescripcion_Click()
        Label4.Visible = True
        Label5.Visible = True
        DtpFecIni.Visible = True
        DTPFecFin.Visible = True
        Lbletiqueta.Caption = "Codigo"
        TxtBuscar.SetFocus
End Sub

Private Sub OptFac_Click()
        Label4.Visible = False
        Label5.Visible = False
        DtpFecIni.Visible = False
        DTPFecFin.Visible = False
        Lbletiqueta.Caption = "Factura"
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

Private Sub TxtBusqueda_Change()
            
                Set RBusqueda = New ADODB.Recordset
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                        If BProveedor = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where Proveedor Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(Proveedor) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        ElseIf BFicha = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where Descrip Like '%" & TxtBusqueda.Text & "%'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
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
                        End If
                    End If
                            
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "5000"
                            

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
        End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
            Set RBusqueda = New ADODB.Recordset
            If Index = 0 Then
                BProveedor = True
                BFicha = False
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores")
            ElseIf Index = 1 Then
                BProveedor = False
                BFicha = True
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            End If
            
            If Index = 0 Or Index = 1 Then
                
                Set DBGridBusqueda.DataSource = RBusqueda
                FrameBuscar.Visible = True
                TxtBusqueda.SetFocus
                DBGridBusqueda.Columns(1).Width = "4000"
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
            If Index = 0 Then
                BProveedor = True
                BFicha = False
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores")
            ElseIf Index = 1 Then
                BProveedor = False
                BFicha = True
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica")
            End If
            
            If Index = 0 Or Index = 1 Then
                
                Set DBGridBusqueda.DataSource = RBusqueda
                FrameBuscar.Visible = True
                TxtBusqueda.SetFocus
                DBGridBusqueda.Columns(1).Width = "4000"
            End If
        End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
        'FECHA
            If IsNull(RTransito!FechaDespacho) Then
                Msk.Item(7).Text = ""
            Else
                Msk.Item(7).Text = RTransito!FechaDespacho
            End If
        'FECHA ARRIBO
            If IsNull(RTransito!FechaArribo) Then
                Msk.Item(8).Text = ""
            Else
                Msk.Item(8).Text = RTransito!FechaArribo
            End If
        'PROVEEDOR
            If IsNull(RTransito!Proveedor) Then
                TxtTexto.Item(0).Text = ""
            Else
                TxtTexto.Item(0).Text = RTransito!Proveedor
            End If
        'CODIGO
            If IsNull(RTransito!Codigo) Then
                TxtTexto.Item(1).Text = ""
            Else
                TxtTexto.Item(1).Text = RTransito!Codigo
            End If
        'CANTIDAD
            If IsNull(RTransito!Cantidad) Then
                Msk.Item(0).Text = ""
            Else
                Msk.Item(0).Text = RTransito!Cantidad
            End If
        'FACTURA
            If IsNull(RTransito!Factura) Then
                TxtTexto.Item(5).Text = ""
            Else
                TxtTexto.Item(5).Text = RTransito!Factura
            End If
        'MONTO DOLARES
            If IsNull(RTransito!MontoDolares) Then
                Msk.Item(1).Text = ""
            Else
                Msk.Item(1).Text = RTransito!MontoDolares
            End If
        'PILOTO
            If IsNull(RTransito!Piloto) Then
                TxtTexto.Item(3).Text = ""
            Else
                TxtTexto.Item(3).Text = RTransito!Piloto
            End If
        'TRANSPORTISTA
            If IsNull(RTransito!Transportista) Then
                TxtTexto.Item(4).Text = ""
            Else
                TxtTexto.Item(4).Text = RTransito!Transportista
            End If
        'USUARIO
            If IsNull(RTransito!Usuario) Then
                TxtTexto.Item(2).Text = ""
            Else
                TxtTexto.Item(2).Text = RTransito!Usuario
            End If
        'RECIBIDA
            If GOrigenDeDatos = "AmaproAccess" Then
                    If RTransito!Recibida = True Then
                        ChkRec.Value = "1"
                    Else
                        ChkRec.Value = "0"
                    End If
            Else
                If RTransito!Recibida = "-1" Then
                    ChkRec.Value = "1"
                Else
                    ChkRec.Value = "0"
                End If
            End If
        'DOCUMENTO
            If IsNull(RTransito!Documento) Then
                TxtDoc.Text = ""
            Else
                TxtDoc.Text = RTransito!Documento
            End If
        
            
        If Err <> 0 Then
            'MsgBox Err.Description
        End If

End Sub

Public Sub Limpia_Campos()
        
        Msk.Item(7).Text = ""
        Msk.Item(8).Text = ""
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        Msk.Item(0).Text = ""
        TxtTexto.Item(5).Text = ""
        Msk.Item(1).Text = 0
        TxtTexto.Item(3).Text = ""
        TxtTexto.Item(4).Text = ""
        TxtTexto.Item(2).Text = ""
        
End Sub



