VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CierrePedidosProveedores 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre De Pedidos De Proveedores"
   ClientHeight    =   8550
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   Icon            =   "CierrePedidosProveedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
      Caption         =   "Buscar Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8535
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbGridBuscar 
         Height          =   7455
         Left            =   120
         TabIndex        =   30
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   13150
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
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   10920
         Picture         =   "CierrePedidosProveedores.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1680
         TabIndex        =   28
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   29
         Top             =   600
         Width           =   6255
      End
   End
   Begin TabDlg.SSTab TabInformacion 
      Height          =   5775
      Left            =   0
      TabIndex        =   14
      Top             =   2280
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Detalle Pedido"
      TabPicture(0)   =   "CierrePedidosProveedores.frx":293C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DbGridDetalle"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Informacion De Codigo"
      TabPicture(1)   =   "CierrePedidosProveedores.frx":2D96
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtDatos"
      Tab(1).Control(1)=   "TxtDatos2"
      Tab(1).Control(2)=   "lblFieldLabel(5)"
      Tab(1).Control(3)=   "lblFieldLabel(6)"
      Tab(1).ControlCount=   4
      Begin MSDataGridLib.DataGrid DbGridDetalle 
         Height          =   3255
         Left            =   240
         TabIndex        =   52
         ToolTipText     =   "doble click o flechas abajo o arriba para seleccionar datos"
         Top             =   1800
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5741
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
         ColumnCount     =   5
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Codigo"
            Caption         =   "Codigo"
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
            DataField       =   "Descrip"
            Caption         =   "Descripcion"
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
            DataField       =   "Pedido"
            Caption         =   "Pedido"
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
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   4245.166
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameDetalle 
         BorderStyle     =   0  'None
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
         ForeColor       =   &H00FF0000&
         Height          =   4935
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   11565
         Begin VB.Frame FrameDetalle2 
            Enabled         =   0   'False
            Height          =   975
            Left            =   120
            TabIndex        =   16
            Top             =   0
            Width           =   11295
            Begin VB.TextBox TxtPed 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   8160
               MaxLength       =   12
               TabIndex        =   18
               Top             =   480
               Width           =   1455
            End
            Begin VB.TextBox TxtDesPro 
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
               Left            =   1920
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   480
               Width           =   6135
            End
            Begin VB.TextBox TxtCod 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   17
               ToolTipText     =   "signo + o doble click para ayuda"
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox TxtDocDet 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1920
               TabIndex        =   36
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
            End
            Begin MSMask.MaskEdBox MskCan 
               Height          =   285
               Left            =   9720
               TabIndex        =   19
               Top             =   480
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "#,###,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
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
               Index           =   0
               Left            =   120
               TabIndex        =   39
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "No. Pedido"
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
               Left            =   8160
               TabIndex        =   38
               Top             =   240
               Width           =   960
            End
            Begin VB.Label Label1 
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
               Height          =   255
               Index           =   5
               Left            =   9720
               TabIndex        =   37
               Top             =   240
               Width           =   1095
            End
         End
         Begin VB.CommandButton CmdAgregar2 
            Caption         =   "A&gregar"
            Height          =   495
            Left            =   120
            Picture         =   "CierrePedidosProveedores.frx":9030
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   4440
            Visible         =   0   'False
            Width           =   2200
         End
         Begin VB.CommandButton CmdGrabar2 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2400
            Picture         =   "CierrePedidosProveedores.frx":9562
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   4440
            Visible         =   0   'False
            Width           =   2200
         End
         Begin VB.CommandButton CmdTerminar 
            Caption         =   "&Terminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9240
            Picture         =   "CierrePedidosProveedores.frx":9A94
            TabIndex        =   25
            Top             =   4440
            Visible         =   0   'False
            Width           =   2200
         End
         Begin VB.CommandButton CmdCancelar2 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4680
            Picture         =   "CierrePedidosProveedores.frx":9FC6
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4440
            Visible         =   0   'False
            Width           =   2200
         End
         Begin VB.CommandButton CmdBorrar2 
            Caption         =   "B&orrar"
            Height          =   495
            Left            =   6960
            Picture         =   "CierrePedidosProveedores.frx":A4F8
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   4440
            Visible         =   0   'False
            Width           =   2200
         End
      End
      Begin VB.TextBox TxtDatos 
         BackColor       =   &H0000C000&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2295
         Left            =   -74760
         MultiLine       =   -1  'True
         TabIndex        =   41
         Top             =   840
         Width           =   11295
      End
      Begin VB.TextBox TxtDatos2 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   3480
         Width           =   11295
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Cierres De Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   5
         Left            =   -66000
         TabIndex        =   43
         Top             =   3120
         Width           =   2580
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         Caption         =   "Datos Del Pedido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Index           =   6
         Left            =   -65880
         TabIndex        =   42
         Top             =   480
         Width           =   2430
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   11400
      MouseIcon       =   "CierrePedidosProveedores.frx":AA2A
      Picture         =   "CierrePedidosProveedores.frx":AE6C
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Ultimo Registro"
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   11040
      MouseIcon       =   "CierrePedidosProveedores.frx":B39E
      Picture         =   "CierrePedidosProveedores.frx":B7E0
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Siguiente Registro"
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   360
      MouseIcon       =   "CierrePedidosProveedores.frx":BD12
      Picture         =   "CierrePedidosProveedores.frx":C154
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Registro Anterior"
      Top             =   8040
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   0
      MouseIcon       =   "CierrePedidosProveedores.frx":C686
      Picture         =   "CierrePedidosProveedores.frx":CAC8
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Primer Registro"
      Top             =   8040
      Width           =   375
   End
   Begin VB.Frame FrameEncabezado 
      Caption         =   "Encabezado de Cierre Pedido"
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
      Height          =   2295
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   480
         Left            =   8760
         Picture         =   "CierrePedidosProveedores.frx":CFFA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1680
         Width           =   1400
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&EDITAR"
         Height          =   480
         Left            =   1560
         Picture         =   "CierrePedidosProveedores.frx":D52C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   1400
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&USCAR"
         Height          =   480
         Left            =   7320
         Picture         =   "CierrePedidosProveedores.frx":DA5E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1680
         Width           =   1400
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   10200
         Picture         =   "CierrePedidosProveedores.frx":DF90
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Salida"
         Top             =   1680
         Width           =   1400
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&BORRAR"
         Height          =   480
         Left            =   5880
         Picture         =   "CierrePedidosProveedores.frx":10002
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1680
         Width           =   1400
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&CANCELAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   4440
         Picture         =   "CierrePedidosProveedores.frx":10534
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   1400
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&GRABAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   3000
         Picture         =   "CierrePedidosProveedores.frx":10A66
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1680
         Width           =   1400
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&AGREGAR"
         Height          =   480
         Left            =   120
         Picture         =   "CierrePedidosProveedores.frx":10F98
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1680
         Width           =   1400
      End
      Begin VB.Frame FrameEncabezado2 
         Enabled         =   0   'False
         Height          =   1335
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   11535
         Begin VB.TextBox TxtNumDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7920
            MaxLength       =   12
            TabIndex        =   2
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox TxtTipDoc 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   3
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   4
            Top             =   960
            Width           =   8055
         End
         Begin VB.TextBox TxtUsu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   9720
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   960
            Width           =   1695
         End
         Begin MSMask.MaskEdBox MskFec 
            Height          =   285
            Left            =   1560
            TabIndex        =   0
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtDoc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   3960
            Locked          =   -1  'True
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
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
            TabIndex        =   47
            Top             =   600
            Width           =   1410
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Numero De Documento"
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
            Left            =   5880
            TabIndex        =   46
            Top             =   240
            Width           =   1980
         End
         Begin VB.Label LblDoc 
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
            TabIndex        =   45
            Top             =   600
            Width           =   8415
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
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
            Left            =   120
            TabIndex        =   44
            Top             =   960
            Width           =   1275
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   10920
            Picture         =   "CierrePedidosProveedores.frx":114CA
            Top             =   120
            Width           =   480
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Transaccion"
            ForeColor       =   &H8000000D&
            Height          =   195
            Left            =   3000
            TabIndex        =   35
            Top             =   240
            Width           =   885
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "CierrePedidosProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VMensaje As String
Dim VTransaccion As Long
Dim VCantidadMateriaPrima As Double
Dim VCodigoMateriaPrima As String
Dim VBodega As String
Dim VNumeroPedido As String
Dim VFechaPedido As Date
Dim VTexto As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim BEstaEditanto As Boolean
Dim BMateriaPrima As Boolean
Dim BNumeroIngreso As Boolean
Dim BDocumento As Boolean
Dim BPedido As Boolean
Dim BEditarEncabezado As Boolean
Dim BEditarDetalle As Boolean

Dim VCantidadEntrada As Single
Dim VFechaEntrada As Date
Dim VCodigo As String
Dim VDiasDeAtraso As Integer
Dim VNumeroDocumento As String
Dim VTipoDocumento As String

Dim RBuscaMateriaPrima As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset
Dim RBuscaSigDoc As New ADODB.Recordset
Dim RBuscaTipoDocumento As New ADODB.Recordset
Dim RBuscaDetalle As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim RBuscaPedido As New ADODB.Recordset
Dim RBuscaInventario As New ADODB.Recordset
Dim RBuscaPedidosPendientes As New ADODB.Recordset
Dim RSumaSaldoPedidos As New ADODB.Recordset
Dim RBuscaPedido2 As New ADODB.Recordset
Dim RBuscaPedido3 As New ADODB.Recordset
Dim RBuscaSaldoEncabezado As New ADODB.Recordset
Dim RBuscaSaldoDetalle As New ADODB.Recordset
Dim RBuscaCierrePedidos As New ADODB.Recordset
Dim REncabezado As New ADODB.Recordset
Dim RDetalle As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Sub Botones1()
    If Bandera = True Then
         FrameEncabezado2.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
                
    Else
         FrameEncabezado2.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
                
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalle2.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalle2.Enabled = False
         CmdAgregar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
    End If
End Sub

Sub BotonesVisiblesDetalle()
    If Bandera3 = True Then
         CmdAgregar2.Visible = True
         CmdGrabar2.Visible = True
         CmdTerminar.Visible = True
         CmdBorrar2.Visible = True
         CmdCancelar2.Visible = True
    Else
         CmdAgregar2.Visible = False
         CmdGrabar2.Visible = False
         CmdTerminar.Visible = False
         CmdBorrar2.Visible = False
         CmdCancelar2.Visible = False
    End If
End Sub
Sub BotonesVisiblesEncabezado()
    If Bandera4 = True Then
         CmdAgregar.Visible = True
         CmdEditar.Visible = True
         CmdGrabar.Visible = True
         CmdCancelar.Visible = True
         CmdBorrar.Visible = True
         CmdBuscar.Visible = True
         CmdImprimir.Visible = True
         CmdSalida.Visible = True
    Else
         CmdAgregar.Visible = False
         CmdEditar.Visible = False
         CmdGrabar.Visible = False
         CmdCancelar.Visible = False
         CmdBorrar.Visible = False
         CmdBuscar.Visible = False
         CmdImprimir.Visible = False
         CmdSalida.Visible = False
    End If

End Sub





Private Sub CmdAgregar2_Click()
    
    Bandera2 = True
    Botones2
    Limpia_CamposDetalle
    
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DbGridDetalle.Enabled = False
        
    BEditarDetalle = False
    TxtDocDet.Text = VTransaccion
    TxtCod.SetFocus
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next

            If GBorrarPedidos = True Then
                  'NO HACE NADA PORQUE SI TIENE ACCESO A ESTA FUNCION
            Else
                  MsgBox "Usted No Tiene Acceso a Esta Funcion, Consulte al Encargado", vbOKOnly + vbInformation, "Informacion"
                  Exit Sub
            End If

            VTransaccion = TxtDoc.Text

            VMensaje = MsgBox("¿Esta Seguro De Borrar El Registro?", vbOKCancel + vbExclamation + vbDefaultButton2, "Esta Seguro")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If VMensaje = vbOK Then
                MousePointer = 11
                
                'BUSCA EL DETALLE DEL DOCUMENTO
                Set RBuscaDetalle = New ADODB.Recordset
                    Call Abrir_Recordset(RBuscaDetalle, "Select * from DetalleCierrePedidosProve where Documento = " & VTransaccion & " Order by Codigo")
                    If RBuscaDetalle.RecordCount > 0 Then
                        
                        Conexion.BeginTrans
                        
                        Do Until RBuscaDetalle.EOF
                        
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Conexion.Execute "Update DetallePedidosProveedores Set CantidadEntregada = CantidadEntregada - " & RBuscaDetalle!Cantidad & ", SaldoPorEntregar = SaldoPorEntregar + " & RBuscaDetalle!Cantidad & " Where Documento = '" & RBuscaDetalle!Pedido & "' And Codigo = '" & RBuscaDetalle!Codigo & "'"
                                    Else 'ORACLE
                                        Conexion.Execute "Update DetallePedidosProveedores Set CantidadEntregada = CantidadEntregada - " & RBuscaDetalle!Cantidad & ", SaldoPorEntregar = SaldoPorEntregar + " & RBuscaDetalle!Cantidad & " Where UPPER(Documento) = '" & UCase(RBuscaDetalle!Pedido) & "' And UPPER(Codigo) = '" & UCase(RBuscaDetalle!Codigo) & "'"
                                    End If
                                
                                    If Err <> 0 Then
                                        Conexion.RollbackTrans
                                        MsgBox "Error, No se actualizarion la cantidad entregada y el saldo del pedido " & Err.Number & " " & Err.Description, vbOKCancel + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                                        
                                '---------- PEDIDO ------------------------------------------------------
                                'BUSCA EL DOCUMENTO DE PEDIDO PARA SUMAR LA CANTIDAD ENTREGADA Y RESTAR EL SALDO POR ENTREGAR DE LA MATERIA PRIMA
                                Set RBuscaPedido = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaPedido, "Select SaldoPorEntregar From DetallePedidosProveedores Where Documento = '" & RBuscaDetalle!Pedido & "' And Codigo = '" & RBuscaDetalle!Codigo & "'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBuscaPedido, "Select SaldoPorEntregar From DetallePedidosProveedores Where UPPER(Documento) = '" & UCase(RBuscaDetalle!Pedido) & "' And UPPER(Codigo) = '" & UCase(RBuscaDetalle!Codigo) & "'")
                                    End If
                                
                                    If RBuscaPedido.RecordCount > 0 Then
                                            'SI EL SALDO POR ENTREGAR ES MAYOR QUE CERO CAMBIA LA FECHA
                                            If RBuscaPedido!SaldoPorEntregar > 0 Then
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Conexion.Execute "Update DetallePedidosProveedores Set FechaEntregatotal = '', DiasDeAtraso = 0 Where Documento = '" & RBuscaDetalle!Pedido & "' And Codigo = '" & RBuscaDetalle!Codigo & "'"
                                                    Else 'ORACLE
                                                        Conexion.Execute "Update DetallePedidosProveedores Set FechaEntregatotal = '', DiasDeAtraso = 0 Where UPPER(Documento) = '" & UCase(RBuscaDetalle!Pedido) & "' And UPPER(Codigo) = '" & UCase(RBuscaDetalle!Codigo) & "'"
                                                    End If
                                                    
                                                    If Err <> 0 Then
                                                        Conexion.RollbackTrans
                                                        MsgBox "Error, No se actualizarion la fecha entrega total y los dias de atraso " & Err.Number & " " & Err.Description, vbOKCancel + vbCritical, "Error"
                                                        Err.Clear
                                                        Exit Sub
                                                    End If
                                                
                                            End If
                                        
                                    End If
                        
                            RBuscaDetalle.MoveNext
                        Loop
                    End If
                        
                    REncabezado.Delete
                    
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147217873 Then
                                Conexion.RollbackTrans
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                        
                    'TERMINA LA TRANSACCION
                    Conexion.CommitTrans
                    
                    'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        REncabezado.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        REncabezado.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_CamposEncabezado
                        
                            Set RDetalle = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
                                    End If
                                        Llena_CamposDetalle
                                        Set DbGridDetalle.DataSource = RDetalle

                    
                MousePointer = 0
            End If
  
            
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
    
                'If GBorrar = False Then
                '      MsgBox "Usted No Tiene Acceso a Esta Funcion, Consulte al Encargado", vbOKOnly + vbInformation, "Informacion"
                '      Exit Sub
                'End If
        
                VTransaccion = TxtDocDet.Text
                VCantidadEntrada = MskCan.Text
                VFechaEntrada = MskFec.Text
                VNumeroPedido = TxtPed.Text
                VCodigo = TxtCod.Text
        
                VMensaje = MsgBox("Esta seguro de borrar el registro", vbYesNo + vbDefaultButton2 + vbExclamation, "Verificar")
                
                If VMensaje = vbYes Then
                        
                        Conexion.BeginTrans
                        
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Conexion.Execute "Update DetallePedidosProveedores Set CantidadEntregada = CantidadEntregada - " & VCantidadEntrada & ", SaldoPorEntregar = SaldoPorEntregar + " & VCantidadEntrada & " Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'"
                                    Else 'ORACLE
                                        Conexion.Execute "Update DetallePedidosProveedores Set CantidadEntregada = CantidadEntregada - " & VCantidadEntrada & ", SaldoPorEntregar = SaldoPorEntregar + " & VCantidadEntrada & " Where UPPER(Documento) = '" & UCase(VNumeroPedido) & "' And UPPER(Codigo) = '" & UCase(VCodigo) & "'"
                                    End If
                                
                                    If Err <> 0 Then
                                        Conexion.RollbackTrans
                                        MsgBox "Error, No se actualizarion la cantidad entregada y el saldo del pedido " & Err.Number & " " & Err.Description, vbOKCancel + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                                        
                                '---------- PEDIDO ------------------------------------------------------
                                'BUSCA EL DOCUMENTO DE PEDIDO PARA SUMAR LA CANTIDAD ENTREGADA Y RESTAR EL SALDO POR ENTREGAR DE LA MATERIA PRIMA
                                Set RBuscaPedido = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaPedido, "Select SaldoPorEntregar From DetallePedidosProveedores Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBuscaPedido, "Select SaldoPorEntregar From DetallePedidosProveedores Where UPPER(Documento) = '" & UCase(VNumeroPedido) & "' And UPPER(Codigo) = '" & UCase(VCodigo) & "'")
                                    End If
                                
                                    If RBuscaPedido.RecordCount > 0 Then
                                            'SI EL SALDO POR ENTREGAR ES MAYOR QUE CERO CAMBIA LA FECHA
                                            If RBuscaPedido!SaldoPorEntregar > 0 Then
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Conexion.Execute "Update DetallePedidosProveedores Set FechaEntregatotal = '', DiasDeAtraso = 0 Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'"
                                                    Else 'ORACLE
                                                        Conexion.Execute "Update DetallePedidosProveedores Set FechaEntregatotal = '', DiasDeAtraso = 0 Where UPPER(Documento) = '" & UCase(VNumeroPedido) & "' And UPPER(Codigo) = '" & UCase(VCodigo) & "'"
                                                    End If
                                                    
                                                    If Err <> 0 Then
                                                        Conexion.RollbackTrans
                                                        MsgBox "Error, No se actualizarion la fecha entrega total y los dias de atraso " & Err.Number & " " & Err.Description, vbOKCancel + vbCritical, "Error"
                                                        Err.Clear
                                                        Exit Sub
                                                    End If
                                                
                                            End If
                                        
                                    End If
                        
                    
                        'BORRA EL REGISTRO
                        Conexion.Execute "Delete From DetalleCierrePedidosProve Where Documento = " & VTransaccion & " And Codigo = '" & VCodigo & "' And Pedido = '" & VNumeroPedido & "'"
                    
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147217873 Then
                                Conexion.RollbackTrans
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            End If
                        End If
                        
                    'TERMINA LA TRANSACCION
                    Conexion.CommitTrans
                    
                    'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RDetalle.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RDetalle.MoveNext
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                                                
                        Llena_CamposDetalle
                                                                    
                MousePointer = 0
            End If
                                
    
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
On Error Resume Next
MousePointer = 11
    If Index = 1 Then
        REncabezado.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        REncabezado.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        REncabezado.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        REncabezado.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If REncabezado.BOF Then
        REncabezado.MoveFirst
    ElseIf REncabezado.EOF Then
        REncabezado.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_CamposEncabezado
    
            Set RDetalle = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
            Else 'ORACLE
                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
            End If
                Llena_CamposDetalle
                Set DbGridDetalle.DataSource = RDetalle
    
MousePointer = 0
End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
    VMensaje = InputBox("Numero De Documento a Buscar")
    If VMensaje <> "" Then
                'REncabezado.MoveFirst
                'REncabezado.Find " Documento = " & VMensaje
                
                Set REncabezado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(REncabezado, "Select * From EncabezadoCierrePedidosProve Where NumeroDocumento Like '" & VMensaje & "%' Order By Documento")
                    Else
                        Call Abrir_Recordset(REncabezado, "Select * From EncabezadoCierrePedidosProve Where UPPER(NumeroDocumento) Like '" & UCase(VMensaje) & "%' Order By Documento")
                    End If
                        Llena_CamposEncabezado
        
                                                
                Llena_CamposEncabezado
                
                        Set RDetalle = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
                            End If
                                Llena_CamposDetalle
                                Set DbGridDetalle.DataSource = RDetalle
    End If
End Sub

Private Sub CmdCancelar_Click()
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    Llena_CamposEncabezado
    FrameDetalle.Visible = True
    DbGridDetalle.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
    
    DbGridDetalle.Enabled = True
    Bandera2 = False
    Botones2
    Llena_CamposDetalle

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
        
    BEditarEncabezado = True
    
    'VALIDA SI TIENE ACCESO
    If GEditarPedidos = True Then
    Else
        MsgBox "Usted No Esta Autorizado Para Editar Un Cierre De Pedido Llame Al Encargado", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If

    
    Bandera = True
    Botones1
    
    MskFec.SetFocus
    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False
    TxtUsu.Text = GUsuario
    
End Sub



Private Sub CmdGrabar2_Click()
On Error Resume Next
    
                'REVISA SI ES NUMERICO LA CANTIDAD DE ENTRADA
                If Not IsNumeric(MskCan.Text) Then
                        MsgBox "Cantidad  De Entrada Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
                                
                'REVISA SI EXISTE EL PEDIDO
                Set RBuscaPedido2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaPedido2, "Select * From DetallePedidosProveedores Where Documento = '" & TxtPed.Text & "' And Codigo = '" & TxtCod.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaPedido2, "Select * From DetallePedidosProveedores Where UPPER(Documento) = '" & UCase(TxtPed.Text) & "' And UPPER(Codigo) = '" & UCase(TxtCod.Text) & "'")
                    End If
                        
                        If RBuscaPedido2.RecordCount > 0 Then
                        Else
                                 MsgBox "No. Pedido y Codigo De Producto, No Existe En Pedidos De Proveedores", vbOKOnly + vbInformation, "Informacion"
                                 Exit Sub
                        End If
                
                
                VCantidadEntrada = MskCan.Text
                VFechaEntrada = MskFec.Text
                VNumeroPedido = TxtPed.Text
                VCodigo = TxtCod.Text
                
                    'INICIA LA TRANSACCION
                    Conexion.BeginTrans
                        
                                    VTexto = "Values(" & TxtDocDet.Text & ", '" ' DOCUMENTO
                                    VTexto = VTexto & TxtCod.Text & "', '" 'CODIGO
                                    VTexto = VTexto & TxtPed.Text & "', " 'PEDIDO
                                    VTexto = VTexto & MskCan.Text & ")" 'CANTIDAD
                                    
                                    Conexion.Execute "Insert Into DetalleCierrePedidosProve " & VTexto
                                    
                                    If Err <> 0 Then
                                        Conexion.RollbackTrans
                                        MsgBox "Error, No se actualizarion la cantidad entregada y el saldo del pedido " & Err.Number & " " & Err.Description, vbOKCancel + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                        
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Conexion.Execute "Update DetallePedidosProveedores Set CantidadEntregada = CantidadEntregada + " & VCantidadEntrada & ", SaldoPorEntregar = SaldoPorEntregar - " & VCantidadEntrada & " Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'"
                                    Else 'ORACLE
                                        Conexion.Execute "Update DetallePedidosProveedores Set CantidadEntregada = CantidadEntregada + " & VCantidadEntrada & ", SaldoPorEntregar = SaldoPorEntregar - " & VCantidadEntrada & " Where UPPER(Documento) = '" & UCase(VNumeroPedido) & "' And UPPER(Codigo) = '" & UCase(VCodigo) & "'"
                                    End If
                                
                                    If Err <> 0 Then
                                        Conexion.RollbackTrans
                                        MsgBox "Error, No se actualizarion la cantidad entregada y el saldo del pedido " & Err.Number & " " & Err.Description, vbOKCancel + vbCritical, "Error"
                                        Err.Clear
                                        Exit Sub
                                    End If
                                        
                                '---------- PEDIDO ------------------------------------------------------
                                'BUSCA EL DOCUMENTO DE PEDIDO PARA SUMAR LA CANTIDAD ENTREGADA Y RESTAR EL SALDO POR ENTREGAR DE LA MATERIA PRIMA
                                Set RBuscaPedido = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaPedido, "Select SaldoPorEntregar, FechaParaEntregar From DetallePedidosProveedores Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBuscaPedido, "Select SaldoPorEntregar, FechaParaEntregar From DetallePedidosProveedores Where UPPER(Documento) = '" & UCase(VNumeroPedido) & "' And UPPER(Codigo) = '" & UCase(VCodigo) & "'")
                                    End If
                                
                                    If RBuscaPedido.RecordCount > 0 Then
                                            'SI EL SALDO POR ENTREGAR ES MAYOR QUE CERO CAMBIA LA FECHA
                                            If RBuscaPedido!SaldoPorEntregar <= 0 Then
                                                        'CALCULA LOS DIAS DE ATRASO
                                                        VDiasDeAtraso = (DateValue(RBuscaPedido!FechaParaEntregar) - DateValue(VFechaEntrada))
                                                                        
                                                        'SI LA VARIABLE VDIASDEATRASO ES MENOR QUE CERO ES PORQUE ENTREGO EL PEDIDO ANTES DE LA FECHA
                                                        If VDiasDeAtraso < 0 Then
                                                            VDiasDeAtraso = VDiasDeAtraso * -1
                                                        Else
                                                            VDiasDeAtraso = 0
                                                        End If
                                            
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Conexion.Execute "Update DetallePedidosProveedores Set FechaEntregatotal = '" & VFechaEntrada & "', DiasDeAtraso = " & VDiasDeAtraso & " Where Documento = '" & VNumeroPedido & "' And Codigo = '" & VCodigo & "'"
                                                    Else 'ORACLE
                                                        Conexion.Execute "Update DetallePedidosProveedores Set FechaEntregatotal = '" & VFechaEntrada & "', DiasDeAtraso = " & VDiasDeAtraso & " Where UPPER(Documento) = '" & UCase(VNumeroPedido) & "' And UPPER(Codigo) = '" & UCase(VCodigo) & "'"
                                                    End If
                                                    
                                                    If Err <> 0 Then
                                                        Conexion.RollbackTrans
                                                        MsgBox "Error, No se actualizarion la fecha entrega total y los dias de atraso " & Err.Number & " " & Err.Description, vbOKCancel + vbCritical, "Error"
                                                        Err.Clear
                                                        Exit Sub
                                                    End If
                                            End If
                                    End If
                                    
                                'AGREGA UN REGISTRO A LOS % DE CONFORME POR PEDIDO Y CODIGO
                                            
                                            VTexto = "'" & VNumeroDocumento & "', '" ' NUMERO DOCUMENTO
                                            VTexto = VTexto & VTipoDocumento & "', '" 'DOCUMENTO
                                            VTexto = VTexto & VNumeroPedido & "', '" 'PEDIDO
                                            VTexto = VTexto & VCodigo & "', " 'CODIGO
                                            VTexto = VTexto & "100" & ", '" '%
                                            VTexto = VTexto & GUsuario & "'" 'USUARIO
                                            
                                            Conexion.Execute "Insert Into PedidosProveedoresPorcentajeNo Values(" & VTexto & ")"
                                    
                                
                                            If Err <> 0 Then
                                                        Conexion.RollbackTrans
                                                        MsgBox "Error, No se Grabo El % De Producto Conforme " & Err.Number & " " & Err.Description, vbOKCancel + vbCritical, "Error"
                                                        Err.Clear
                                                        Exit Sub
                                            End If
                                                    
                                    
                                    
                    'FINALIZA LA TRANSACCION
                    Conexion.CommitTrans
                                
        
                        Bandera2 = False
                        Botones2
                        RDetalle.Requery
                        RDetalle.MoveLast
                        Llena_CamposDetalle
                        DbGridDetalle.Enabled = True
                        CmdAgregar2.SetFocus
End Sub


Private Sub CmdAgregar_Click()
    On Error Resume Next
    
    TxtDoc.Enabled = True
    Bandera = True
    Botones1
    BEditarEncabezado = False
    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False
    Limpia_CamposEncabezado
    'ASIGNA EL USUARIO
    TxtUsu.Text = GUsuario
    'ASIGNA LA FECHA DEL DIA
    MskFec.Text = Date
    MskFec.SetFocus
    
    'BUSCA EL DOCUMENTO MAXIMO Y LE ASIGNA 1
    Set RBuscaSigDoc = New ADODB.Recordset
        Call Abrir_Recordset(RBuscaSigDoc, "Select Max(Documento) from EncabezadoCierrePedidosProve")
        If RBuscaSigDoc.RecordCount > 0 Then
            If IsNull(RBuscaSigDoc(0)) Then
                TxtDoc.Text = "1"
            Else
                TxtDoc.Text = RBuscaSigDoc(0) + 1
            End If
        End If
    
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    
    
    MskFec.Text = Format(MskFec.Text, "dd/mm/yyyy")
    
    
    'VERIFICA LA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Error"
        MskFec.SetFocus
        Exit Sub
    End If
    
    'VERIFICA EL NUMERO DE DOCUMENTO
    If TxtNumDoc.Text = "" Then
        MsgBox "Documento Incorrecto", vbOKOnly + vbInformation, "Error"
        TxtNumDoc.SetFocus
        Exit Sub
    End If
    
    'VERIFICA EL NUMERO DE DOCUMENTO
    If TxtTipDoc.Text = "" Then
        MsgBox "Tipo Documento Incorrecto", vbOKOnly + vbInformation, "Error"
        TxtTipDoc.SetFocus
        Exit Sub
    End If
    
    
    'SI ESTA AGREGANDO UN REGISTRO LO REVISA
    If BEditarEncabezado = False Then
            'VERIFICA QUE EL DOCUMENTO YA EXISTE
            Set RBuscaDocumento = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaDocumento, "Select * From EncabezadoCierrePedidosProve Where TipoDeDocumento = '" & TxtTipDoc.Text & "' And NumeroDocumento = '" & TxtNumDoc.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaDocumento, "Select * From EncabezadoCierrePedidosProve Where Upper(TipoDeDocumento) = '" & UCase(TxtTipDoc.Text) & "' And Upper(NumeroDocumento) = '" & UCase(TxtNumDoc.Text) & "'")
                End If
                If RBuscaDocumento.RecordCount > 0 Then
                    MsgBox "Numero De Documento y Tipo De Documento Ya Existe", vbOKOnly + vbInformation, "Error"
                    TxtNumDoc.SetFocus
                    Exit Sub
                End If
    End If
    
    VTransaccion = TxtDoc.Text
    VFechaPedido = MskFec.Text
    VNumeroDocumento = TxtNumDoc.Text
    VTipoDocumento = TxtTipDoc.Text
    
                    'AGREGAR
                    If BEditarEncabezado = False Then
                            VTexto = "Values(" & TxtDoc.Text & ", " 'DOCUMENTO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                            Else 'ORACLE
                                VTexto = VTexto & "To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                            End If
                            VTexto = VTexto & TxtTipDoc.Text & "', '" 'TIPO DE DOCUMENTO
                            VTexto = VTexto & TxtNumDoc & "', '" 'NUMERO DE DOCUMENTO
                            VTexto = VTexto & TxtObs & "', '" 'NUMERO DE DOCUMENTO
                            VTexto = VTexto & GUsuario & "')" 'NUMERO DE DOCUMENTO
                            
                            Conexion.Execute "Insert Into EncabezadoCierrePedidosProve " & VTexto
                    'EDITAR
                    Else
                            
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = "Fecha = #" & Format(MskFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = "Fecha = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "TipoDeDocumento = '" & UCase(TxtTipDoc.Text) & "', " 'TIPO DE DOCUMENTO
                            VTexto = VTexto & "NumeroDocumento = '" & UCase(TxtNumDoc) & "', " 'NUMERO DE DOCUMENTO
                            VTexto = VTexto & "Observaciones = '" & TxtObs & "', " 'OBSERVACIONES
                            VTexto = VTexto & "Usuario = '" & GUsuario & "' " 'USUARIO
                            VTexto = VTexto & "Where Documento = " & VTransaccion & " " 'DOCUMENTO
                            
                            Conexion.Execute "UPDATE EncabezadoCierrePedidosProve SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Transaccion Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtNumDoc.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Transaccion Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtNumDoc.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        Botones1
                        TxtNumDoc.Enabled = True
                        
                        Set REncabezado = New ADODB.Recordset
                        Call Abrir_Recordset(REncabezado, "Select * From EncabezadoCierrePedidosProve Where Documento = " & VTransaccion & " Order By Documento")
                        
                        Llena_CamposEncabezado
   
                        Set RDetalle = New ADODB.Recordset
                                 If GOrigenDeDatos = "AmaproAccess" Then
                                     Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & VTransaccion & " And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                                 Else 'ORACLE
                                     Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & VTransaccion & " And E.Documento = D.Documento And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
                                 End If
                                     Llena_CamposDetalle
                                     Set DbGridDetalle.DataSource = RDetalle

                
                'ESCONDE LOS BOTONES DEL ENCABEZADO
                Bandera4 = False
                BotonesVisiblesEncabezado
                
                'VISUALIZA LOS BOTONES DEL DETALLE
                Bandera3 = True
                BotonesVisiblesDetalle
                
    
                'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
                'BOTONES DE DATA
                CmdBotones2.Item(1).Visible = False
                CmdBotones2.Item(2).Visible = False
                CmdBotones2.Item(3).Visible = False
                CmdBotones2.Item(4).Visible = False
                
    
                FrameDetalle.Visible = True
                DbGridDetalle.Visible = True
                FrameDetalle.Enabled = True
                FrameEncabezado.Enabled = False
                CmdAgregar2.SetFocus
End Sub

Private Sub CmdImprimir_Click()
MousePointer = 11
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "CierrePedidosProveedoresResumen.rpt"
                Else
                    GNombreReporte = "CierrePedidosProveedoresResumenO.rpt"
                End If
                GCriteriaReporte = "{EncabezadoCierrePedidosProve.Documento} = " & TxtDoc.Text
                FrmReporte.Show
        
MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdTerminar_Click()
If CmdCancelar2.Enabled = True Then
     CmdCancelar2_Click
End If
    
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    CmdBotones2.Item(1).Visible = True
    CmdBotones2.Item(2).Visible = True
    CmdBotones2.Item(3).Visible = True
    CmdBotones2.Item(4).Visible = True
    FrameDetalle.Visible = True
    FrameEncabezado.Enabled = True
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesVisiblesEncabezado
    
    'ESCONDE LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle
    
    Set REncabezado = New ADODB.Recordset
            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoCierrePedidosProve Order By Documento")
            REncabezado.MoveLast
                Llena_CamposEncabezado
                
        Set RDetalle = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
            Else 'ORACLE
                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
            End If
                Llena_CamposDetalle
                Set DbGridDetalle.DataSource = RDetalle


End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub



Private Sub DBGridBuscar_DblClick()
        If BMateriaPrima = True Then
            TxtCod.Text = DbGridBuscar.Columns(0)
            TxtCod.SetFocus
        ElseIf BPedido = True Then
            TxtPed.Text = DbGridBuscar.Columns(1)
            TxtPed.SetFocus
        ElseIf BDocumento = True Then
            TxtTipDoc.Text = DbGridBuscar.Columns(0)
            TxtTipDoc.SetFocus
        End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
End Sub

Private Sub DbGridBuscar_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BMateriaPrima = True Then
                TxtCod.Text = DbGridBuscar.Columns(0)
                TxtCod.SetFocus
            ElseIf BPedido = True Then
                TxtPed.Text = DbGridBuscar.Columns(1)
                TxtPed.SetFocus
            ElseIf BDocumento = True Then
                TxtTipDoc.Text = DbGridBuscar.Columns(0)
                TxtTipDoc.SetFocus
            End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
        End If
End Sub





Private Sub DbGridDetalle_HeadClick(ByVal ColIndex As Integer)
        RDetalle.Sort = RDetalle.Fields(ColIndex).Name
End Sub


Private Sub DbGridDetalle_SelChange(Cancel As Integer)
Llena_CamposDetalle
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
    

End Sub

Private Sub Form_Load()
        Set REncabezado = New ADODB.Recordset
            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoCierrePedidosProve Order By Documento")
                Llena_CamposEncabezado
                
        Set RDetalle = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
            Else 'ORACLE
                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.Pedido, D.Cantidad From EncabezadoCierrePedidosProve E, DetalleCierrePedidosProve D, FichaTecnica F Where E.Documento = " & TxtDoc.Text & " And E.Documento = D.Documento And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
            End If
                Llena_CamposDetalle
                Set DbGridDetalle.DataSource = RDetalle

    
End Sub

Private Sub MskCan_GotFocus()
    MskCan.SelStart = 0
    MskCan.SelLength = Len(MskCan.Text)
End Sub

Private Sub MskCan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub
Private Sub MskFec_GotFocus()
    MskFec.SelStart = 0
    MskFec.SelLength = Len(MskFec.Text)
End Sub

Private Sub MskFec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
End Sub

Private Sub TabInformacion_Click(PreviousTab As Integer)
    'TAB DE INFORMACION DEL CODIGO
    If TabInformacion.Tab = 1 Then
            
        'CAMBIA EL CAPTION DEL TAB CON LA DESCRIPCION DEL CODIGO
        TabInformacion.Caption = TxtDesPro.Text
        
       'BUSCA EL ENCABEZADO DE PEDIDO
        Set RBuscaSaldoEncabezado = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaSaldoEncabezado, "Select EP.Fecha, P.Descripcion, EP.Observaciones, EP.Documento From EncabezadoPedidosProveedores EP, Proveedores P Where EP.Documento = '" & TxtPed.Text & "' And EP.Proveedor = P.CodigoProveedor")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaSaldoEncabezado, "Select EP.Fecha, P.Descripcion, EP.Observaciones, EP.Documento From EncabezadoPedidosProveedores EP, Proveedores P Where UPPER(EP.Documento) = '" & UCase(TxtPed.Text) & "' And UPPER(EP.Proveedor) = UPPER(P.CodigoProveedor)")
            End If
                   If RBuscaSaldoEncabezado.RecordCount > 0 Then
                       TxtDatos.Text = ""
                       TxtDatos.Text = TxtDatos.Text & "No. Pedido       " & RBuscaSaldoEncabezado(3) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Fecha Pedido     " & RBuscaSaldoEncabezado(0) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Proveedor          " & RBuscaSaldoEncabezado(1) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "Observaciones    " & RBuscaSaldoEncabezado(2) & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "         Pedido           Entregado              Saldo     Dias      Entrega    Entregado   Atraso" & vbCrLf
                       TxtDatos.Text = TxtDatos.Text & "__________________________________________________________________________________________________" & vbCrLf
                       'BUSCA EL DETALLE DEL PEDIDO
                       Set RBuscaSaldoDetalle = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaSaldoDetalle, "Select * From DetallePedidosProveedores Where Documento = '" & TxtPed.Text & "' And Codigo = '" & TxtCod.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaSaldoDetalle, "Select * From DetallePedidosProveedores Where UPPER(Documento) = '" & UCase(TxtPed.Text) & "' And UPPER(Codigo) = '" & UCase(TxtCod.Text) & "'")
                            End If
                           Do Until RBuscaSaldoDetalle.EOF
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!CantidadPedido) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!CantidadEntregada) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatSingle(RBuscaSaldoDetalle!SaldoPorEntregar) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatInteger5(RBuscaSaldoDetalle!DiasPedido) & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & RBuscaSaldoDetalle!FechaParaEntregar & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & RBuscaSaldoDetalle!FechaEntregaTotal & Space(3)
                                   TxtDatos.Text = TxtDatos.Text & FormatInteger5(RBuscaSaldoDetalle!DiasDeAtraso) & Space(3) & vbCrLf
                               RBuscaSaldoDetalle.MoveNext
                           Loop
                   Else
                       TxtDatos.Text = ""
                   End If
                   
                    'BUSCA TODOS LOS CIERRES QUE TIENE EL PEDIDO
                    Set RBuscaCierrePedidos = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaCierrePedidos, "Select EC.Fecha, D.Descripcion, EC.NumeroDocumento, DC.Cantidad From DetalleCierrePedidosProve DC, EncabezadoCierrePedidosProve EC, Documentos as D Where DC.Pedido = '" & TxtPed.Text & "' And DC.Codigo = '" & TxtCod.Text & "' And DC.Documento = EC.Documento And EC.TipoDeDocumento = D.CodigoDocumento")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaCierrePedidos, "Select EC.Fecha, D.Descripcion, EC.NumeroDocumento, DC.Cantidad From DetalleCierrePedidosProve DC, EncabezadoCierrePedidosProve EC, Documentos D Where UPPER(DC.Pedido) = '" & UCase(TxtPed.Text) & "' And UPPER(DC.Codigo) = '" & UCase(TxtCod.Text) & "' And DC.Documento = EC.Documento And UPPER(EC.TipoDeDocumento) = UPPER(D.CodigoDocumento)")
                        End If
                        If RBuscaCierrePedidos.RecordCount > 0 Then
                                TxtDatos2.Text = ""
                                TxtDatos2.Text = TxtDatos2.Text & "Fecha      Documento                     No. Documento               Cantidad" & vbCrLf
                                TxtDatos2.Text = TxtDatos2.Text & "___________________________________________________________________________________________________" & vbCrLf
                                    Do Until RBuscaCierrePedidos.EOF
                                            TxtDatos2.Text = TxtDatos2.Text & RBuscaCierrePedidos(0) & " " & Left(RBuscaCierrePedidos(1) & Space(30), 30) & FormatString15(RBuscaCierrePedidos(2)) & Space(5) & FormatSingle(RBuscaCierrePedidos(3)) & vbCrLf
                                        RBuscaCierrePedidos.MoveNext
                                    Loop
                        Else
                                TxtDatos2.Text = ""
                        End If
        
    End If
End Sub

Private Sub Txtbuscar_Change()
        Set RBusqueda = New ADODB.Recordset
        'MATERIA PRIMA
        If BMateriaPrima = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '%" & TxtBuscar.Text & "%' Order by Esp_Tec")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Esp_Tec")
                    End If
            ElseIf OptDescripcion.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Esp_Tec")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Esp_Tec")
                    End If
            End If
        'DOCUMENTO
        ElseIf BDocumento = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion From Documentos Where CodigoDocumento Like '%" & TxtBuscar.Text & "%' Order by Descripcion")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion From Documentos Where UPPER(CodigoDocumento) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Descripcion")
                    End If
            ElseIf OptDescripcion.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion From Documentos Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by Descripcion")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion From Documentos Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Descripcion")
                    End If
            End If
        End If
            
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"

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


Private Sub TxtCod_Change()
                'BUSCA LA DESCRIPCION DEL CODIGO
                Set RBuscaMateriaPrima = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCod.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCod.Text) & "'")
                    End If
                
                If RBuscaMateriaPrima.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaMateriaPrima!Descrip
                Else
                        TxtDesPro.Text = ""
                End If
End Sub
Private Sub TxtCod_DblClick()
            BMateriaPrima = True
            BDocumento = False
            BPedido = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Order by Esp_Tec")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"
End Sub
Private Sub TxtCod_GotFocus()
        TxtCod.SelStart = 0
        TxtCod.SelLength = Len(TxtCod.Text)
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
        'SI PRECIONA ENTER
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
        'SI PRECIONA LA TECLA DE SIGNO +
        If KeyAscii = 43 Then
            BMateriaPrima = True
            BDocumento = False
            BPedido = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Order by Esp_Tec")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"
        
        End If
End Sub


Private Sub TxtDesPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtDoc_GotFocus()
    TxtDoc.SelStart = 0
    TxtDoc.SelLength = Len(TxtDoc.Text)
End Sub

Private Sub TxtDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub


Private Sub TxtNumDoc_GotFocus()
        TxtNumDoc.SelStart = 0
        TxtNumDoc.SelLength = Len(TxtNumDoc.Text)
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtObs_GotFocus()
        TxtObs.SelStart = 0
        TxtObs.SelLength = Len(TxtObs.Text)
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtPed_DblClick()
        Set RBusqueda = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBusqueda, "Select P.Fecha, P.Documento, DP.CantidadPedido, DP.CantidadEntregada, DP.SaldoPorEntregar, DP.FechaParaEntregar, Pr.Descripcion From EncabezadoPedidosProveedores P, Proveedores Pr, DetallePedidosProveedores DP Where DP.Codigo = '" & TxtCod.Text & "' And DP.SaldoPorEntregar > 0 And P.Documento = DP.Documento And P.Proveedor = Pr.CodigoProveedor Order By Fecha")
            Else 'ORACLE
                Call Abrir_Recordset(RBusqueda, "Select P.Fecha, P.Documento, DP.CantidadPedido, DP.CantidadEntregada, DP.SaldoPorEntregar, DP.FechaParaEntregar, Pr.Descripcion From EncabezadoPedidosProveedores P, Proveedores Pr, DetallePedidosProveedores DP Where UPPER(DP.Codigo) = '" & UCase(TxtCod.Text) & "' And DP.SaldoPorEntregar > 0 And P.Documento = DP.Documento And UPPER(P.Proveedor) = UPPER(Pr.CodigoProveedor) Order By Fecha")
            End If
        
        BDocumento = False
        BMateriaPrima = False
        BPedido = True
        Set DbGridBuscar.DataSource = RBusqueda
                        DbGridBuscar.Columns(0).Width = 1000
                        DbGridBuscar.Columns(1).Width = 1200
                        DbGridBuscar.Columns(2).Width = 1200
                        DbGridBuscar.Columns(3).Width = 1200
                        DbGridBuscar.Columns(4).Width = 1200
                        DbGridBuscar.Columns(5).Width = 1200
                        DbGridBuscar.Columns(6).Width = 2500
                        DbGridBuscar.Columns(0).Caption = "Fecha"
                        DbGridBuscar.Columns(1).Caption = "Pedido"
                        DbGridBuscar.Columns(2).Caption = "Inicio"
                        DbGridBuscar.Columns(3).Caption = "Entregado"
                        DbGridBuscar.Columns(4).Caption = "Saldo"
                        DbGridBuscar.Columns(5).Caption = "Entregar"
                        DbGridBuscar.Columns(6).Caption = "Proveedor"
                        DbGridBuscar.Columns(2).NumberFormat = "#,###,##0"
                        DbGridBuscar.Columns(3).NumberFormat = "#,###,##0"
                        DbGridBuscar.Columns(4).NumberFormat = "#,###,##0"
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus

End Sub

Private Sub TxtPed_GotFocus()
    TxtPed.SelStart = 0
    TxtPed.SelLength = Len(TxtPed.Text)
End Sub

Private Sub TxtPed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    ElseIf KeyAscii = 43 Then
        Set RBusqueda = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBusqueda, "Select P.Fecha, P.Documento, DP.CantidadPedido, DP.CantidadEntregada, DP.SaldoPorEntregar, DP.FechaParaEntregar, Pr.Descripcion From EncabezadoPedidosProveedores P, Proveedores Pr, DetallePedidosProveedores DP Where DP.Codigo = '" & TxtCod.Text & "' And DP.SaldoPorEntregar > 0 And P.Documento = DP.Documento And P.Proveedor = Pr.CodigoProveedor Order By Fecha")
            Else 'ORACLE
                Call Abrir_Recordset(RBusqueda, "Select P.Fecha, P.Documento, DP.CantidadPedido, DP.CantidadEntregada, DP.SaldoPorEntregar, DP.FechaParaEntregar, Pr.Descripcion From EncabezadoPedidosProveedores P, Proveedores Pr, DetallePedidosProveedores DP Where UPPER(DP.Codigo) = '" & UCase(TxtCod.Text) & "' And DP.SaldoPorEntregar > 0 And P.Documento = DP.Documento And UPPER(P.Proveedor) = UPPER(Pr.CodigoProveedor) Order By Fecha")
            End If
        BDocumento = False
        BMateriaPrima = False
        BPedido = True
        
        Set DbGridBuscar.DataSource = RBusqueda
                        DbGridBuscar.Columns(0).Width = 1000
                        DbGridBuscar.Columns(1).Width = 1200
                        DbGridBuscar.Columns(2).Width = 1200
                        DbGridBuscar.Columns(3).Width = 1200
                        DbGridBuscar.Columns(4).Width = 1200
                        DbGridBuscar.Columns(5).Width = 1200
                        DbGridBuscar.Columns(6).Width = 2500
                        DbGridBuscar.Columns(0).Caption = "Fecha"
                        DbGridBuscar.Columns(1).Caption = "Pedido"
                        DbGridBuscar.Columns(2).Caption = "Inicio"
                        DbGridBuscar.Columns(3).Caption = "Entregado"
                        DbGridBuscar.Columns(4).Caption = "Saldo"
                        DbGridBuscar.Columns(5).Caption = "Entregar"
                        DbGridBuscar.Columns(6).Caption = "Proveedor"
                        DbGridBuscar.Columns(2).NumberFormat = "#,###,##0"
                        DbGridBuscar.Columns(3).NumberFormat = "#,###,##0"
                        DbGridBuscar.Columns(4).NumberFormat = "#,###,##0"
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
    End If

End Sub

Private Sub TxtTipDoc_Change()
        Set RBuscaTipoDocumento = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaTipoDocumento, "Select Descripcion From Documentos Where CodigoDocumento = '" & TxtTipDoc.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaTipoDocumento, "Select Descripcion From Documentos Where UPPER(CodigoDocumento) = '" & UCase(TxtTipDoc.Text) & "'")
            End If
            If RBuscaTipoDocumento.RecordCount > 0 Then
                LblDoc.Caption = RBuscaTipoDocumento!Descripcion
            Else
                LblDoc.Caption = ""
            End If
        
End Sub

Private Sub TxtTipDoc_DblClick()
            BMateriaPrima = False
            BDocumento = True
            BPedido = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Order by Descripcion")
            Set DbGridBuscar.DataSource = RBusqueda
            
            DbGridBuscar.Columns(1).Width = "4000"

End Sub

Private Sub TxtTipDoc_GotFocus()
        TxtTipDoc.SelStart = 0
        TxtTipDoc.SelLength = Len(TxtTipDoc.Text)
End Sub

Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BMateriaPrima = False
            BDocumento = True
            BPedido = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Order by Descripcion")
            Set DbGridBuscar.DataSource = RBusqueda
            DbGridBuscar.Columns(1).Width = "4000"
        End If
End Sub


Public Sub Llena_CamposEncabezado()
On Error Resume Next
            If REncabezado.RecordCount > 0 Then
                TxtDoc.Text = REncabezado!Documento
                MskFec.Text = REncabezado!fecha
                TxtTipDoc.Text = REncabezado!TipoDeDocumento
                TxtNumDoc.Text = REncabezado!NumeroDocumento
                TxtObs.Text = REncabezado!Observaciones
                TxtUsu.Text = REncabezado!Usuario
            Else
                TxtDoc.Text = ""
                MskFec.Text = ""
                TxtTipDoc.Text = ""
                TxtNumDoc.Text = ""
                TxtObs.Text = ""
                TxtUsu.Text = ""
            End If
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Llena_CamposDetalle()
On Error Resume Next
            If RDetalle.RecordCount > 0 Then
                TxtDocDet.Text = RDetalle!Documento
                TxtCod.Text = RDetalle!Codigo
                TxtPed.Text = RDetalle!Pedido
                MskCan.Text = RDetalle!Cantidad
            Else
                TxtDocDet.Text = ""
                TxtCod.Text = ""
                TxtPed.Text = ""
                MskCan.Text = ""
            End If
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Limpia_CamposEncabezado()
                TxtDoc.Text = ""
                MskFec.Text = ""
                TxtTipDoc.Text = ""
                TxtNumDoc.Text = ""
                TxtObs.Text = ""
                TxtUsu.Text = ""
End Sub

Public Sub Limpia_CamposDetalle()
                TxtDocDet.Text = ""
                TxtCod.Text = ""
                TxtPed.Text = ""
                MskCan.Text = ""
End Sub


