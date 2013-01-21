VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form PedidosProveedores 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pedidos De Proveedores"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "PedidosProveedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
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
      Height          =   8055
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6855
         Left            =   120
         TabIndex        =   38
         Top             =   960
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   12091
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
         Picture         =   "PedidosProveedores.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1680
         TabIndex        =   36
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Width           =   6255
      End
   End
   Begin TabDlg.SSTab TabInformacion 
      Height          =   5652
      Left            =   0
      TabIndex        =   15
      Top             =   2400
      Width           =   11772
      _ExtentX        =   20770
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Detalle Pedido"
      TabPicture(0)   =   "PedidosProveedores.frx":293C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameDetalle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "CmdBotones2(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdBotones2(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "CmdBotones2(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "CmdBotones2(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DbGridDetalle"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Informacion De Codigo"
      TabPicture(1)   =   "PedidosProveedores.frx":4646
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtDatosInventario"
      Tab(1).Control(1)=   "TxtDatosPedido"
      Tab(1).Control(2)=   "Label6(5)"
      Tab(1).Control(3)=   "Label6(4)"
      Tab(1).Control(4)=   "Label3"
      Tab(1).Control(5)=   "LblSalPed"
      Tab(1).Control(6)=   "Label6(3)"
      Tab(1).ControlCount=   7
      Begin MSDataGridLib.DataGrid DbGridDetalle 
         Height          =   2895
         Left            =   240
         TabIndex        =   66
         Top             =   1920
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5106
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
            DataField       =   "Precio"
            Caption         =   "Precio"
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
            DataField       =   "CantidadPedido"
            Caption         =   "Requerido"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "CantidadEntregada"
            Caption         =   "Entregado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "SaldoPorEntregar"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "DiasPedido"
            Caption         =   "Dias Plazo"
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
            DataField       =   "FechaParaEntregar"
            Caption         =   "Entrega"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column09 
            DataField       =   "FechaEntregadoTotal"
            Caption         =   "Entregado"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "DiasAtraso"
            Caption         =   "Dias Atraso"
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
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1170.142
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   3750.236
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   374.74
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   959.811
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               ColumnWidth     =   345.26
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   4
         Left            =   11160
         MouseIcon       =   "PedidosProveedores.frx":9778
         Picture         =   "PedidosProveedores.frx":9BBA
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Ultimo Registro"
         Top             =   5040
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   3
         Left            =   10800
         MouseIcon       =   "PedidosProveedores.frx":A0EC
         Picture         =   "PedidosProveedores.frx":A52E
         Style           =   1  'Graphical
         TabIndex        =   64
         ToolTipText     =   "Siguiente Registro"
         Top             =   5040
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   2
         Left            =   600
         MouseIcon       =   "PedidosProveedores.frx":AA60
         Picture         =   "PedidosProveedores.frx":AEA2
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Registro Anterior"
         Top             =   5040
         Width           =   375
      End
      Begin VB.CommandButton CmdBotones2 
         BackColor       =   &H00C0C0C0&
         Height          =   465
         Index           =   1
         Left            =   240
         MouseIcon       =   "PedidosProveedores.frx":B3D4
         Picture         =   "PedidosProveedores.frx":B816
         Style           =   1  'Graphical
         TabIndex        =   62
         ToolTipText     =   "Primer Registro"
         Top             =   5040
         Width           =   375
      End
      Begin VB.TextBox TxtDatosInventario 
         Appearance      =   0  'Flat
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
         Height          =   1125
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   56
         Top             =   1080
         Width           =   11535
      End
      Begin VB.TextBox TxtDatosPedido 
         Appearance      =   0  'Flat
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
         Height          =   2565
         Left            =   -74880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   2640
         Width           =   11535
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
         Height          =   4815
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   11565
         Begin VB.Frame FrameDetalle2 
            Enabled         =   0   'False
            Height          =   1215
            Left            =   120
            TabIndex        =   17
            Top             =   0
            Width           =   11295
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
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   480
               Width           =   6135
            End
            Begin VB.TextBox TxtCod 
               Appearance      =   0  'Flat
               DataField       =   "Codigo"
               DataSource      =   "DataDetalle"
               Height          =   285
               Left            =   120
               MaxLength       =   15
               TabIndex        =   18
               ToolTipText     =   "signo + o doble click para ayuda"
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox TxtDocDet 
               Appearance      =   0  'Flat
               DataField       =   "Documento"
               DataSource      =   "DataDetalle"
               Height          =   285
               Left            =   6120
               MaxLength       =   15
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
            End
            Begin VB.TextBox TxtDiaAtr 
               Appearance      =   0  'Flat
               DataField       =   "DiasDeAtraso"
               DataSource      =   "DataDetalle"
               Enabled         =   0   'False
               Height          =   285
               Left            =   8160
               MaxLength       =   3
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   840
               Width           =   700
            End
            Begin VB.TextBox TxtDiaPla 
               Appearance      =   0  'Flat
               DataField       =   "DiasPedido"
               DataSource      =   "DataDetalle"
               Enabled         =   0   'False
               Height          =   285
               Left            =   3720
               TabIndex        =   25
               ToolTipText     =   "signo + o doble click para ayuda"
               Top             =   840
               Width           =   700
            End
            Begin MSMask.MaskEdBox MskFecEntTot 
               DataField       =   "FechaEntregaTotal"
               DataSource      =   "DataDetalle"
               Height          =   285
               Left            =   6000
               TabIndex        =   46
               Top             =   840
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskFecEnt 
               DataField       =   "FechaParaEntregar"
               DataSource      =   "DataDetalle"
               Height          =   285
               Left            =   1680
               TabIndex        =   24
               Top             =   840
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskSal 
               DataField       =   "SaldoPorEntregar"
               DataSource      =   "DataDetalle"
               Height          =   285
               Left            =   9720
               TabIndex        =   27
               Top             =   840
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Enabled         =   0   'False
               Format          =   "#,###,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskCanEnt 
               DataField       =   "CantidadEntregada"
               DataSource      =   "DataDetalle"
               Height          =   285
               Left            =   4080
               TabIndex        =   23
               Top             =   120
               Visible         =   0   'False
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Enabled         =   0   'False
               Format          =   "#,###,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskCanPed 
               DataField       =   "CantidadPedido"
               DataSource      =   "DataDetalle"
               Height          =   285
               Left            =   8160
               TabIndex        =   21
               Top             =   480
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,###,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskPre 
               DataField       =   "CantidadPedido"
               DataSource      =   "DataDetalle"
               Height          =   285
               Left            =   9720
               TabIndex        =   22
               Top             =   480
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
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
               TabIndex        =   54
               Top             =   240
               Width           =   735
            End
            Begin VB.Label Label1 
               Caption         =   "Fecha a Entregar"
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
               Index           =   3
               Left            =   120
               TabIndex        =   53
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad Pedido"
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
               TabIndex        =   52
               Top             =   240
               Width           =   1410
            End
            Begin VB.Label Label1 
               Caption         =   "Dias Plazo"
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
               Index           =   1
               Left            =   2760
               TabIndex        =   51
               Top             =   840
               Width           =   975
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Dias Atraso"
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
               Left            =   7080
               TabIndex        =   50
               Top             =   840
               Width           =   990
            End
            Begin VB.Label Label1 
               Caption         =   "Precio"
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
               TabIndex        =   49
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Saldo"
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
               Left            =   9120
               TabIndex        =   48
               Top             =   840
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Entregado"
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
               Left            =   4440
               TabIndex        =   47
               Top             =   840
               Width           =   1470
            End
         End
         Begin VB.CommandButton CmdAgregar2 
            Caption         =   "A&gregar"
            Height          =   495
            Left            =   960
            Picture         =   "PedidosProveedores.frx":BD48
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   4320
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CommandButton CmdGrabar2 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4080
            Picture         =   "PedidosProveedores.frx":C27A
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4320
            Visible         =   0   'False
            Width           =   1500
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
            Left            =   9000
            Picture         =   "PedidosProveedores.frx":C7AC
            TabIndex        =   33
            Top             =   4320
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.CommandButton CmdCancelar2 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   5640
            Picture         =   "PedidosProveedores.frx":CCDE
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   4320
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.CommandButton CmdBorrar2 
            Caption         =   "B&orrar"
            Height          =   495
            Left            =   7320
            Picture         =   "PedidosProveedores.frx":D210
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   4320
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.CommandButton CmdEditar2 
            Caption         =   "Editar"
            Height          =   495
            Left            =   2520
            Picture         =   "PedidosProveedores.frx":D742
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   4320
            Visible         =   0   'False
            Width           =   1500
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Inventario"
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
         Left            =   -74880
         TabIndex        =   61
         Top             =   720
         Width           =   1380
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Pedidos Pendientes"
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
         Index           =   4
         Left            =   -74880
         TabIndex        =   60
         Top             =   2280
         Width           =   2820
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Total"
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
         Left            =   -66960
         TabIndex        =   59
         Top             =   5280
         Width           =   990
      End
      Begin VB.Label LblSalPed 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   -65760
         TabIndex        =   58
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "                  Bultos                  Unidades                                             Peso"
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
         Index           =   3
         Left            =   -71880
         TabIndex        =   57
         Top             =   840
         Width           =   8535
      End
   End
   Begin VB.Frame FrameEncabezado 
      Caption         =   "Encabezado de Pedido"
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
      Height          =   2415
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   11775
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   480
         Left            =   8760
         Picture         =   "PedidosProveedores.frx":DC74
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "&EDITAR"
         Height          =   480
         Left            =   1560
         Picture         =   "PedidosProveedores.frx":E1A6
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "B&USCAR"
         Height          =   480
         Left            =   7320
         Picture         =   "PedidosProveedores.frx":E6D8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdSalida 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   10200
         Picture         =   "PedidosProveedores.frx":EC0A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Salida"
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdBorrar 
         Caption         =   "&BORRAR"
         Height          =   480
         Left            =   5880
         Picture         =   "PedidosProveedores.frx":10C7C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&CANCELAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   4440
         Picture         =   "PedidosProveedores.frx":111AE
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&GRABAR"
         Enabled         =   0   'False
         Height          =   480
         Left            =   3000
         Picture         =   "PedidosProveedores.frx":116E0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1800
         Width           =   1400
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&AGREGAR"
         Height          =   480
         Left            =   120
         Picture         =   "PedidosProveedores.frx":11C12
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1800
         Width           =   1400
      End
      Begin VB.Frame FrameEncabezado2 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   120
         TabIndex        =   41
         Top             =   240
         Width           =   11535
         Begin VB.TextBox TxtIva 
            Appearance      =   0  'Flat
            DataField       =   "Cliente"
            DataSource      =   "DataEncabezado"
            Height          =   285
            Left            =   3600
            MaxLength       =   10
            TabIndex        =   1
            TabStop         =   0   'False
            Text            =   "12"
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   360
            Width           =   375
         End
         Begin VB.TextBox TxtUsu 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "Usuario"
            DataSource      =   "DataEncabezado"
            Height          =   285
            Left            =   7440
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox TxtCli 
            Appearance      =   0  'Flat
            DataField       =   "Cliente"
            DataSource      =   "DataEncabezado"
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   3
            ToolTipText     =   "signo + o doble click para ayuda"
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox TxtObs 
            Appearance      =   0  'Flat
            DataField       =   "Observaciones"
            DataSource      =   "DataEncabezado"
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   5
            Top             =   1080
            Width           =   7335
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataEncabezado"
            Height          =   285
            Left            =   1560
            TabIndex        =   0
            Top             =   360
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
            DataField       =   "Documento"
            DataSource      =   "DataEncabezado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   5160
            MaxLength       =   12
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
         Begin MSMask.MaskEdBox MskSubTot 
            DataField       =   "CantidadPedido"
            DataSource      =   "DataDetalle"
            Height          =   285
            Left            =   9960
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskIva 
            DataField       =   "CantidadPedido"
            DataSource      =   "DataDetalle"
            Height          =   285
            Left            =   9960
            TabIndex        =   72
            TabStop         =   0   'False
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskTot 
            DataField       =   "CantidadPedido"
            DataSource      =   "DataDetalle"
            Height          =   285
            Left            =   9960
            TabIndex        =   73
            TabStop         =   0   'False
            Top             =   1080
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "#,###,##0.00"
            PromptChar      =   "_"
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Total"
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
            Left            =   9000
            TabIndex        =   70
            Top             =   1080
            Width           =   450
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Iva"
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
            Left            =   9000
            TabIndex        =   69
            Top             =   720
            Width           =   285
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total"
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
            Left            =   9000
            TabIndex        =   68
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label6 
            Caption         =   "% Iva"
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
            Index           =   6
            Left            =   3000
            TabIndex        =   67
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Proveedor"
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
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label LblCli 
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
            TabIndex        =   4
            Top             =   720
            Width           =   4335
         End
         Begin VB.Label Label6 
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
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   1080
            Width           =   1335
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
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
            Left            =   4080
            TabIndex        =   43
            Top             =   360
            Width           =   975
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
            TabIndex        =   42
            Top             =   360
            Width           =   855
         End
      End
   End
End
Attribute VB_Name = "PedidosProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As String
Dim VDocumentoDetalle As String
Dim VCantidadMateriaPrima As Double
Dim VCodigoMateriaPrima As String
Dim VBodega As String
Dim VNumeroPedido As String
Dim VProveedor As String
Dim VFechaPedido As Date

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim BMateriaPrima As Boolean
Dim BNumeroIngreso As Boolean
Dim BProveedor As Boolean
Dim BTransportista As Boolean
Dim BPedido As Boolean
Dim BEsFichaTecnica As Boolean
Dim BEditarEncabezado As Boolean
Dim BEditarDetalle As Boolean


Dim RBuscaMateriaPrima As New ADODB.Recordset
Dim RBuscaNumeroIngreso As New ADODB.Recordset
Dim RBuscaCliente As New ADODB.Recordset
Dim RBuscaSigDoc As New ADODB.Recordset
Dim RBuscaDetalle As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim RBuscaPedido As New ADODB.Recordset
Dim RBuscaTransportista As New ADODB.Recordset
Dim RBuscaInventario As New ADODB.Recordset
Dim RBuscaPedidosPendientes As New ADODB.Recordset
Dim RSumaSaldoPedidos As New ADODB.Recordset
Dim RBuscaSiEsFichaTecnica As New ADODB.Recordset
Dim RBuscaInventarioPT As New ADODB.Recordset
Dim RBuscaEntregas As New ADODB.Recordset

Dim RSumaPrecios As New ADODB.Recordset
Dim REncabezado As New ADODB.Recordset
Dim RDetalle As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim VTexto As String
Dim VCodigo As String


Dim VIva As Currency
Dim VSubTotal As Currency
Dim VPorIva As Currency


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
         CmdEditar2.Enabled = False
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
    Else
         FrameDetalle2.Enabled = False
         CmdAgregar2.Enabled = True
         CmdEditar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
    End If
End Sub

Sub BotonesVisiblesDetalle()
    If Bandera3 = True Then
         CmdAgregar2.Visible = True
         CmdEditar2.Visible = True
         CmdGrabar2.Visible = True
         CmdTerminar.Visible = True
         CmdBorrar2.Visible = True
         CmdCancelar2.Visible = True
    Else
         CmdAgregar2.Visible = False
         CmdEditar2.Visible = False
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
On Error Resume Next
        
    Bandera2 = True
    Botones2
    Limpia_CamposDetalle
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DbGridDetalle.Enabled = False
    'SE ASIGNA AL DOCUMENTO DE DETALLE EL DOCUMENTO DEL ENCABEZADO
    TxtDocDet.Text = VDocumento
    TxtCod.SetFocus
    
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next

            If GBorrarPedidos = False Then
                   MsgBox "Usted No Tiene Acceso a Esta Funcion, Consulte al Encargado", vbOKOnly + vbInformation, "Informacion"
                   Exit Sub
            End If
            
            Set RBuscaEntregas = New ADODB.Recordset
                Call Abrir_Recordset(RBuscaEntregas, "Select Sum(CantidadEntregada) From DetallepedidosProveedores Where Documento = '" & TxtDoc.Text & "'")
                    If RBuscaEntregas.RecordCount > 0 Then
                        If IsNull(RBuscaEntregas(0)) Then
                        
                        ElseIf RBuscaEntregas(0) > 0 Then
                            MsgBox "Este Pedido No Se Puede Borrar Porque Ya Hay Cierres De Pedido", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                    Else
                    
                    End If
                
            

            VDocumento = TxtDoc.Text

            mensaje = MsgBox("¿Esta Seguro De Borrar El Pedido?", vbOKCancel + vbExclamation + vbDefaultButton2, "Esta Seguro")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
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
                                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
                                                End If
                                                
                                                Llena_CamposDetalle
                                                Set DbGridDetalle.DataSource = RDetalle
                   
            End If
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
            
            Set RBuscaEntregas = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaEntregas, "Select Sum(CantidadEntregada) From DetallepedidosProveedores Where Documento = '" & TxtDocDet.Text & "' And Codigo = '" & TxtCod.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaEntregas, "Select Sum(CantidadEntregada) From DetallepedidosProveedores Where UPPER(Documento) = '" & UCase(TxtDocDet.Text) & "' And UPPER(Codigo) = '" & UCase(TxtCod.Text) & "'")
                End If
                    
                    If RBuscaEntregas.RecordCount > 0 Then
                        If IsNull(RBuscaEntregas(0)) Then
                        
                        ElseIf RBuscaEntregas(0) > 0 Then
                            MsgBox "Este Codigo No Se Puede Borrar Porque Ya Tiene Cantidad Entregada", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                    Else
                    
                    End If
            
            
            VDocumento = TxtDocDet.Text
            VCodigo = TxtCod.Text
                                       
    
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
            
                Conexion.BeginTrans
                
                        'BORRA EL REGISTRO
                        Conexion.Execute "Delete From DetallePedidosProveedores Where Documento = '" & VDocumento & "' And Codigo = '" & VCodigo & "'"
                    
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
                        
                          'SUMA LOS PRECIOS
                                Set RSumaPrecios = New ADODB.Recordset
                                    Call Abrir_Recordset(RSumaPrecios, "Select Sum(Precio) From DetallePedidosProveedores Where Documento = '" & VDocumento & "'")
                                    If RSumaPrecios.RecordCount > 0 Then
                                        If IsNull(RSumaPrecios(0)) Then
                                            VSubTotal = 0
                                        Else
                                            VSubTotal = RSumaPrecios(0)
                                        End If
                                    Else
                                            VSubTotal = 0
                                    End If
                                   
                                    'SACA EL MONTO DEL IVA
                                    VIva = Val(VSubTotal) * Val(VPorIva)
                                    
                                    Conexion.Execute "Update EncabezadoPedidosProveedores Set Iva = " & VIva & ", Total = " & VSubTotal & " Where Documento = '" & VDocumento & "'"
                                    
                                    'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                        Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    Else 'ORACLE
                                    
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    MskSubTot.Text = VSubTotal
                                    MskIva.Text = VIva
                                    MskTot.Text = Val(VSubTotal) + Val(VIva)
                                    
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
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
                                End If
                                Llena_CamposDetalle
                                Set DbGridDetalle.DataSource = RDetalle
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click()
    
    mensaje = InputBox("Pedido a Buscar")
    If mensaje <> "" Then
                REncabezado.MoveFirst
                If GOrigenDeDatos = "AmaproAccess" Then
                    REncabezado.Find "Documento = '" & mensaje & "'"
                Else
                    REncabezado.Find "Documento = '" & UCase(mensaje) & "'"
                End If
    
                                                
                Llena_CamposEncabezado
                
                        Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
                                End If
                                Llena_CamposDetalle
                                Set DbGridDetalle.DataSource = RDetalle
    End If
    
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    Llena_CamposEncabezado
    FrameDetalle.Visible = True
    DbGridDetalle.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    
    DbGridDetalle.Enabled = True
    Bandera2 = False
    Botones2
    Llena_CamposDetalle
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next

    'VALIDA SI TIENE ACCESO
    If GEditarPedidos = True Then
    Else
        MsgBox "Usted No Esta Autorizado Para Modificar Un Pedido Llame Al Encargado", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    BEditarEncabezado = True
    Bandera = True
    Botones1
    TxtDoc.Enabled = False
    MskFec.SetFocus
    TxtUsu.Text = GUsuario
    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False
End Sub


Private Sub CmdEditar2_Click()
On Error Resume Next

    'VALIDA SI TIENE ACCESO
    If GEditar = True Then
        MskCanPed.Enabled = True
    Else
        MskCanPed.Enabled = False
    End If
    
    Set RBuscaEntregas = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaEntregas, "Select Sum(CantidadEntregada) From DetallepedidosProveedores Where Documento = '" & TxtDocDet.Text & "' And Codigo = '" & TxtCod.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaEntregas, "Select Sum(CantidadEntregada) From DetallepedidosProveedores Where UPPER(Documento) = '" & UCase(TxtDocDet.Text) & "' And UPPER(Codigo) = '" & UCase(TxtCod.Text) & "'")
                End If
                    
                    If RBuscaEntregas.RecordCount > 0 Then
                        If IsNull(RBuscaEntregas(0)) Then
                        
                        ElseIf RBuscaEntregas(0) > 0 Then
                            MsgBox "Este Codigo No Se Puede Editar Porque Ya Tiene Cantidad Entregada", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                    Else
                    
                    End If
    
    
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    VCodigo = TxtCod.Text
    DbGridDetalle.Enabled = False
    BEditarDetalle = True
    Bandera2 = True
    Botones2
    TxtCod.SetFocus
End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
    
                        VDocumento = TxtDocDet.Text
                        
                        If IsNumeric(MskPre.Text) Then
                        Else
                            MsgBox "Precio Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            MskPre.SetFocus
                            Exit Sub
                        End If
                        
                        If IsNumeric(MskCanPed.Text) Then
                        Else
                            MsgBox "Cantidad De Pedido Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            MskCanPed.SetFocus
                            Exit Sub
                        End If
                        
                        If IsDate(MskFecEnt.Text) Then
                        Else
                            MsgBox "Fecha De Entrega Incorrecta", vbOKOnly + vbInformation, "Informacion"
                            MskFecEnt.SetFocus
                            Exit Sub
                        End If
                        
                        'REVISA SI EXISTE LA MATERIA PRIMA
                        Set RBuscaMateriaPrima = New ADODB.Recordset
                            Call Abrir_Recordset(RBuscaMateriaPrima, "Select * From FichaTecnica Where Esp_Tec = '" & TxtCod.Text & "'")
                                If RBuscaMateriaPrima.RecordCount > 0 Then
                                Else
                                    MsgBox "El Codigo No Existe ", vbOKOnly + vbInformation, "Informacion"
                                    Exit Sub
                                End If
                                  
                        'SALDO DE PEDIDO ES IGUAL A CANTIDAD DE PEDIDO MENOS CANTIDAD ENTREGADA
                        MskSal.Text = Val(MskCanPed.Text) - Val(MskCanEnt.Text)
                        
                        MskFecEnt.Text = Format(MskFecEnt.Text, "dd/mm/yyyy")
                        
                        'EMPIEZA LA CONEXION
                        Conexion.BeginTrans
                        
                            If BEditarDetalle = False Then
                                    VTexto = "'" & TxtDocDet.Text & "', '" ' DOCUMENTO
                                    VTexto = VTexto & TxtCod.Text & "', " 'CODIGO
                                    VTexto = VTexto & MskCanPed.Text & ", " 'PEDIDO
                                    VTexto = VTexto & MskCanEnt.Text & ", " 'ENTREGADO
                                    VTexto = VTexto & MskSal.Text & ", " 'SALDO
                                    VTexto = VTexto & TxtDiaPla.Text & ", " 'DIAS PLAZO
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        VTexto = VTexto & "#" & Format(MskFecEnt.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                                    Else 'ORACLE
                                        VTexto = VTexto & "To_Date('" & MskFecEnt.Text & "', 'dd/mm/yyyy')" & ", '"  'FECHA
                                    End If
                                    VTexto = VTexto & MskFecEntTot.Text & "', " 'FECHA ENTREGA TOTAL
                                    VTexto = VTexto & TxtDiaAtr.Text & ", " 'DIAS DE ATRASO
                                    VTexto = VTexto & MskPre.Text  'PRECIO
                                    
                                    Conexion.Execute "Insert Into DetallePedidosProveedores Values(" & VTexto & ")"
                            Else 'SI ESTA EDITANDO
                                    'VTexto = "'" & TxtDocDet.Text & "', '" ' DOCUMENTO
                                    'VTexto = VTexto & TxtLin.Text & "', '" 'LINEA
                                    
                                    VTexto = "CantidadPedido = " & MskCanPed.Text & ", "  'PEDIDO
                                    VTexto = VTexto & "Codigo = '" & TxtCod.Text & "', " 'ENTREGADO
                                    VTexto = VTexto & "CantidadEntregada = " & MskCanEnt.Text & ", " 'ENTREGADO
                                    VTexto = VTexto & "SaldoPorEntregar = " & MskSal.Text & ", " 'SALDO
                                    VTexto = VTexto & "DiasPedido = " & TxtDiaPla.Text & ", " 'DIAS PLAZO
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        VTexto = VTexto & "FechaParaEntregar = #" & Format(MskFecEnt.Text, "mm/dd/yyyy") & "#, " 'FECHA
                                    Else 'ORACLE
                                        VTexto = VTexto & "FechaParaEntregar = To_Date('" & MskFecEnt.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                                    End If
                                    VTexto = VTexto & "FechaEntregaTotal = '" & MskFecEntTot.Text & "', " 'FECHA ENTREGA TOTAL
                                    VTexto = VTexto & "DiasDeAtraso = " & TxtDiaAtr.Text & ", " 'DIAS DE ATRASO
                                    VTexto = VTexto & "Precio = " & MskPre.Text 'PRECIO
                                    VTexto = VTexto & " Where Documento = '" & TxtDocDet.Text & "' And Codigo = '" & VCodigo & "'"
                                    
                                    Conexion.Execute "Update DetallePedidosProveedores Set " & VTexto
                            End If
                                        
                                    'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    Else 'ORACLE
                                        If Err = -2147217873 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Orden y Codigo Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                            TxtCod.SetFocus
                                            Exit Sub
                                      'SI ES CUALQUIER OTRO ERROR
                                        ElseIf Err <> -2147217873 And Err <> 0 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    End If
                                    
                                'SUMA LOS PRECIOS
                                Set RSumaPrecios = New ADODB.Recordset
                                    Call Abrir_Recordset(RSumaPrecios, "Select Sum(Precio) From DetallePedidosProveedores Where Documento = '" & VDocumento & "'")
                                    If RSumaPrecios.RecordCount > 0 Then
                                        If IsNull(RSumaPrecios(0)) Then
                                            VSubTotal = 0
                                        Else
                                            VSubTotal = RSumaPrecios(0)
                                        End If
                                    Else
                                            VSubTotal = 0
                                    End If
                                   
                                    'SACA EL MONTO DEL IVA
                                    VIva = Val(VSubTotal) * Val(VPorIva)
                                    
                                    Conexion.Execute "Update EncabezadoPedidosProveedores Set Iva = " & VIva & ", Total = " & VSubTotal & " Where Documento = '" & VDocumento & "'"
                                    
                                    'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                        Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    Else 'ORACLE
                                    
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                            Conexion.RollbackTrans
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    End If
                                    
                                    MskSubTot.Text = VSubTotal
                                    MskIva.Text = VIva
                                    MskTot.Text = Val(VSubTotal) + Val(VIva)
                                    
                        
                        'TERMINA LA TRANSACCION
                        Conexion.CommitTrans
                        
    
    Bandera2 = False
    Botones2
    TxtCod.Enabled = True
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
    TxtUsu.Text = GUsuario
    MskFec.Text = Date
    MskFec.SetFocus
    TxtDoc.SetFocus
    
   
    
    'BUSCA EL DOCUMENTO MAXIMO Y LE ASIGNA 1
    'Set RBuscaSigDoc = Db.OpenRecordset("Select Max(Documento) from EncabezadoPedidosProveedores")
    '    If RBuscaSigDoc.RecordCount > 0 Then
    '        If IsNull(RBuscaSigDoc(0)) Then
    '            TxtDoc.Text = "1"
    '        Else
    '            TxtDoc.Text = RBuscaSigDoc(0) + 1
    '        End If
    '    End If
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next
    
    VDocumento = TxtDoc.Text
    VProveedor = TxtCli.Text
    MskFec.Text = Format(MskFec.Text, "dd/mm/yyyy")
    VFechaPedido = Format(MskFec.Text, "dd/mm/yyyy")
    If IsNumeric(TxtIva.Text) Then
        'CALCULA EL % DE IVA
        VPorIva = "0." & TxtIva.Text
    Else
        MsgBox "El % De Iva Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
        
                    If BEditarEncabezado = False Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                            Else 'ORACLE
                                VTexto = "To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", '"  'FECHA
                            End If
                            VTexto = VTexto & TxtDoc.Text & "', '" 'DOCUMENTO
                            VTexto = VTexto & TxtCli.Text & "', '" 'PROVEEDOR
                            VTexto = VTexto & TxtObs.Text & "', '" 'OBSERCACIONES
                            VTexto = VTexto & GUsuario & "', 0, 0" 'USUARIO, IVA , SUBTOTAL
                            
                            Conexion.Execute "Insert Into EncabezadoPedidosProveedores Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            VTexto = "Proveedor = '" & UCase(TxtCli.Text) & "', " 'Proveedor
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "Fecha = #" & Format(MskFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = VTexto & "Fecha = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "Observaciones = '" & UCase(TxtObs.Text) & "', " 'OBSERVACIONES
                            VTexto = VTexto & "Usuario = '" & GUsuario & "'" 'USUARIO
                            VTexto = VTexto & " Where Documento = '" & VDocumento & "'" 'DOCUMENTO
                            
                            Conexion.Execute "UPDATE EncabezadoPedidosProveedores SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                      'SI ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Pedido Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtDoc.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    TxtDoc.Enabled = True
    
                        REncabezado.Requery
                        'Set REncabezado = New ADODB.Recordset
                        'Call Abrir_Recordset(REncabezado, "Select * From EncabezadoPedidosProveedores Where Documento = '" & VDocumento & "' Order By Fecha")
                        
                        'Llena_CamposEncabezado
                        
                        Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
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
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDoc.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'gtituloreporte = "letras = '" & VLetras & "'"
        'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "PedidosProveedoresDetalle.rpt"
                Else
                    GNombreReporte = "PedidosProveedoresDetalleO.rpt"
                End If
                GCriteriaReporte = "{EncabezadoPedidosProveedores.Documento} = '" & TxtDoc.Text & "'"
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
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
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
            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoPedidosProveedores Order By Fecha")
            REncabezado.MoveLast
                Llena_CamposEncabezado
                
                        Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
                                End If
                                Llena_CamposDetalle
                                Set DbGridDetalle.DataSource = RDetalle
                
    
End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub



Private Sub DBGridBusqueda_DblClick()
        If BMateriaPrima = True Then
            TxtCod.Text = DBGridBusqueda.Columns(0)
            TxtCod.SetFocus
        ElseIf BProveedor = True Then
            TxtCli.Text = DBGridBusqueda.Columns(0)
            TxtCli.SetFocus
        End If
            TxtBuscar.Text = ""
            FrameBuscar.Visible = False
End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            If BMateriaPrima = True Then
                TxtCod.Text = DBGridBusqueda.Columns(0)
                TxtCod.SetFocus
            ElseIf BProveedor = True Then
                TxtCli.Text = DBGridBusqueda.Columns(0)
                TxtCli.SetFocus
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
On Error Resume Next
    Set REncabezado = New ADODB.Recordset
            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoPedidosProveedores Order By Fecha")
            REncabezado.MoveLast
            If Err <> 0 Then
            End If
                Llena_CamposEncabezado
                
                        Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where E.Documento = '" & TxtDoc.Text & "' And E.Documento = D.Documento And D.Codigo = F.Esp_Tec")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.Codigo, F.Descrip, D.CantidadPedido, D.CantidadEntregada, D.SaldoPorEntregar, D.Diaspedido, D.FechaParaEntregar, D.FechaEntregaTotal, D.DiasDeAtraso, D.Precio From EncabezadoPedidosProveedores E, DetallePedidosProveedores D, FichaTecnica F Where UPPER(E.Documento) = '" & UCase(TxtDoc.Text) & "' And UPPER(E.Documento) = UPPER(D.Documento) And UPPER(D.Codigo) = UPPER(F.Esp_Tec)")
                                End If
                                Llena_CamposDetalle
                                Set DbGridDetalle.DataSource = RDetalle
                
        
End Sub

Private Sub MskCanEnt_GotFocus()
    MskCanEnt.SelStart = 0
    MskCanEnt.SelLength = Len(MskCanEnt.Text)
End Sub

Private Sub MskCanEnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub


Private Sub MskCanPed_GotFocus()
        MskCanPed.SelStart = 0
        MskCanPed.SelLength = Len(MskCanPed.Text)
End Sub

Private Sub MskCanPed_KeyPress(KeyAscii As Integer)
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

Private Sub MskFecEnt_GotFocus()
        MskFecEnt.SelStart = 0
        MskFecEnt.SelLength = Len(MskFecEnt.Text)
End Sub

Private Sub MskFecEnt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub MskFecEnt_LostFocus()
        If IsDate(MskFecEnt.Text) Then
            TxtDiaPla.Text = DateValue(MskFecEnt.Text) - DateValue(VFechaPedido)
        End If

End Sub

Private Sub MskFecEntTot_GotFocus()
        MskFecEntTot.SelStart = 0
        MskFecEntTot.SelLength = Len(MskFecEntTot.Text)
End Sub

Private Sub MskFecEntTot_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub MskPre_GotFocus()
        MskPre.SelStart = 0
        MskPre.SelLength = Len(MskPre.Text)
End Sub

Private Sub MskPre_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskSal_GotFocus()
        MskSal.SelStart = 0
        MskSal.SelLength = Len(MskSal.Text)
End Sub

Private Sub MskSal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub


Private Sub TabInformacion_Click(PreviousTab As Integer)
    If TabInformacion.Tab = 1 Then
        
        'CAMBIA EL CAPTION DEL TAB CON LA DESCRIPCION DEL CODIGO
        TabInformacion.Caption = TxtDesPro.Text
        
        'BUSCA EL INVENTARIO ACTUAL DE FICHA TECNICA
                Set RBuscaInventarioPT = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaInventarioPT, "SELECT De.Bodega, B.Descripcion, Count(DE.Saldo), Sum(DE.Saldo) From DetalleEntradasInventario DE, BodegasInventario B Where DE.Fichatecnica = '" & TxtCod.Text & "' And DE.Saldo > 0 AND DE.Bodega = B.CodigoBodega Group By DE.Bodega, B.Descripcion")
                    Else
                        Call Abrir_Recordset(RBuscaInventarioPT, "SELECT De.Bodega, B.Descripcion, Count(DE.Saldo), Sum(DE.Saldo) From DetalleEntradasInventario DE, BodegasInventario B Where UPPER(DE.Fichatecnica) = '" & UCase(TxtCod.Text) & "' And DE.Saldo > 0 AND UPPER(DE.Bodega) = UPPER(B.CodigoBodega) Group By DE.Bodega, B.Descripcion")
                    End If
                    If RBuscaInventarioPT.RecordCount > 0 Then
                            TxtDatosInventario.Text = ""
                            Do Until RBuscaInventarioPT.EOF
                                    TxtDatosInventario.Text = TxtDatosInventario.Text & Left(RBuscaInventarioPT(0) & Space(3), 3) & Space(2)
                                    TxtDatosInventario.Text = TxtDatosInventario.Text & Left(RBuscaInventarioPT(1) & Space(30), 30) & Space(2)
                                    TxtDatosInventario.Text = TxtDatosInventario.Text & FormatInteger5(RBuscaInventarioPT(2)) & Space(2)
                                    TxtDatosInventario.Text = TxtDatosInventario.Text & FormatSingle(RBuscaInventarioPT(3)) & Space(2) & vbCrLf
                                RBuscaInventarioPT.MoveNext
                            Loop
                    Else
                            TxtDatosInventario.Text = "No Hay Inventario"
                    End If
                
                'BUSCA PEDIDOS PENDIENTES
                Set RBuscaPedidosPendientes = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaPedidosPendientes, "Select EP.Fecha, P.Descripcion, EP.Observaciones, EP.Documento, DP.* From EncabezadoPedidosProveedores EP, Proveedores P, DetallePedidosProveedores DP Where EP.Documento = DP.Documento And EP.Proveedor = P.CodigoProveedor And DP.Codigo = '" & TxtCod.Text & "' And DP.SaldoPorEntregar > 0 ")
                    Else
                        Call Abrir_Recordset(RBuscaPedidosPendientes, "Select EP.Fecha, P.Descripcion, EP.Observaciones, EP.Documento, DP.* From EncabezadoPedidosProveedores EP, Proveedores P, DetallePedidosProveedores DP Where UPPER(EP.Documento) = UPPER(DP.Documento) And UPPER(EP.Proveedor) = UPPER(P.CodigoProveedor) And UPPER(DP.Codigo) = '" & UCase(TxtCod.Text) & "' And DP.SaldoPorEntregar > 0 ")
                    End If
                   If RBuscaPedidosPendientes.RecordCount > 0 Then
                            TxtDatosPedido.Text = ""
                      Do Until RBuscaPedidosPendientes.EOF
                                TxtDatosPedido.Text = TxtDatosPedido.Text & "No. Pedido: " & " " & Left(RBuscaPedidosPendientes(3) & Space(30), 30) & Space(3) & "Fecha Pedido:  " & " " & RBuscaPedidosPendientes(0) & vbCrLf
                                TxtDatosPedido.Text = TxtDatosPedido.Text & "Proveedor:    " & " " & Left(RBuscaPedidosPendientes(1) & Space(30), 30) & Space(3) & "Observaciones: " & " " & Left(RBuscaPedidosPendientes(2) & Space(30), 30) & vbCrLf
                                'TxtDatosPedido.Text = TxtDatosPedido.Text & vbCrLf
                                'TxtDatosPedido.Text = TxtDatosPedido.Text & "         Pedido           Entregado              Saldo     Dias      Entrega" & vbCrLf
                                'TxtDatosPedido.Text = TxtDatosPedido.Text & FormatInteger5(RBuscaPedidosPendientes!DiasPedido) & Space(3)
                                TxtDatosPedido.Text = TxtDatosPedido.Text & "Entregar: " & RBuscaPedidosPendientes!FechaParaEntregar & Space(3)
                                TxtDatosPedido.Text = TxtDatosPedido.Text & "Pedido: " & FormatSingle(RBuscaPedidosPendientes!CantidadPedido) & Space(3)
                                TxtDatosPedido.Text = TxtDatosPedido.Text & "Entregado: " & FormatSingle(RBuscaPedidosPendientes!CantidadEntregada) & Space(3)
                                TxtDatosPedido.Text = TxtDatosPedido.Text & "Saldo: " & FormatSingle(RBuscaPedidosPendientes!SaldoPorEntregar) & Space(3) & vbCrLf
                                TxtDatosPedido.Text = TxtDatosPedido.Text & "___________________________________________________________________________________________________________" & vbCrLf
                            RBuscaPedidosPendientes.MoveNext
                       Loop
                   Else
                       TxtDatosPedido.Text = "No Hay Pedidos Pendientes"
                   End If
                   
                   'SUMA TODOS LOS SALDOS EN PEDIDOS DE EL CODIGO
                   Set RSumaSaldoPedidos = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RSumaSaldoPedidos, "Select Sum(SaldoPorEntregar) From DetallePedidosProveedores Where Codigo = '" & TxtCod.Text & "'")
                        Else
                            Call Abrir_Recordset(RSumaSaldoPedidos, "Select Sum(SaldoPorEntregar) From DetallePedidosProveedores Where UPPER(Codigo) = '" & UCase(TxtCod.Text) & "'")
                        End If
                            If RSumaSaldoPedidos.RecordCount > 0 Then
                                If Not IsNull(RSumaSaldoPedidos(0)) Then
                                    LblSalPed.Caption = Format(RSumaSaldoPedidos(0), "#,###,##0.00")
                                Else
                                    LblSalPed.Caption = ""
                                End If
                            Else
                                LblSalPed.Caption = ""
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
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Esp_Tec")
                    End If
            ElseIf OptDescripcion.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '%" & TxtBuscar.Text & "%' Order by Esp_Tec")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Esp_Tec")
                    End If
            End If
        'Descripcion
        ElseIf BProveedor = True Then
            'SI VA A BUSCAR POR CODIGO O POR DESCRIPCION
            If OptCodigo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores Where CodigoProveedor Like '%" & TxtBuscar.Text & "%' Order by CodigoProveedor")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores Where UPPER(CodigoProveedor) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoProveedor")
                    End If
            ElseIf OptDescripcion.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by CodigoProveedor")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoProveedor")
                    End If
            End If
        End If
            
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"

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

Private Sub TxtDiaAtr_GotFocus()
        TxtDiaAtr.SelStart = 0
        TxtDiaAtr.SelLength = Len(TxtDiaAtr.Text)
End Sub

Private Sub TxtDiaAtr_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub TxtDiaPla_GotFocus()
        TxtDiaPla.SelStart = 0
        TxtDiaPla.SelLength = Len(TxtDiaPla.Text)
End Sub

Private Sub TxtDiaPla_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub

Private Sub TxtIva_GotFocus()
        TxtIva.SelStart = 0
        TxtIva.SelLength = Len(TxtIva.Text)
End Sub

Private Sub TxtIva_KeyPress(KeyAscii As Integer)
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

Private Sub TxtCli_Change()
            Set RBuscaCliente = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtCli.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtCli.Text) & "'")
                End If
                If RBuscaCliente.RecordCount > 0 Then
                    LblCli.Caption = RBuscaCliente!Descripcion
                Else
                    LblCli.Caption = ""
                End If
End Sub

Private Sub TxtCli_DblClick()
            BMateriaPrima = False
            BProveedor = True
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Order by CodigoProveedor")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtCli_GotFocus()
            TxtCli.SelStart = 0
            TxtCli.SelLength = Len(TxtCli.Text)
End Sub

Private Sub TxtCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
            BMateriaPrima = False
            BProveedor = True
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
           Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Order by CodigoProveedor")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
    End If
End Sub

Private Sub TxtCod_Change()
                Set RBuscaMateriaPrima = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaMateriaPrima, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCod.Text & "'")
                    Else
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
            BProveedor = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Order by Esp_Tec")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
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
            BProveedor = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
           'SELECCIONA TODO EL CATALOGO
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip from FichaTecnica Order by Esp_Tec")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
        
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


Public Sub Llena_CamposEncabezado()
On Error Resume Next
            If REncabezado.RecordCount > 0 Then
                If IsNull(REncabezado!Documento) Then
                    TxtDoc.Text = ""
                Else
                    TxtDoc.Text = REncabezado!Documento
                End If
                If IsNull(REncabezado!Proveedor) Then
                    TxtCli.Text = ""
                Else
                    TxtCli.Text = REncabezado!Proveedor
                End If
                If IsNull(REncabezado!fecha) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = REncabezado!fecha
                End If
                If IsNull(REncabezado!Observaciones) Then
                    TxtObs.Text = ""
                Else
                    TxtObs.Text = REncabezado!Observaciones
                End If
                If IsNull(REncabezado!Usuario) Then
                    TxtUsu.Text = ""
                Else
                    TxtUsu.Text = REncabezado!Usuario
                End If
                If IsNull(REncabezado!total) Then
                    MskSubTot.Text = "0"
                Else
                    MskSubTot.Text = REncabezado!total
                End If
                If IsNull(REncabezado!Iva) Then
                    MskIva.Text = "0"
                Else
                    MskIva.Text = REncabezado!Iva
                End If
                MskTot.Text = Val(MskSubTot.Text) + Val(MskIva)
            Else
                TxtDoc.Text = ""
                TxtObs.Text = ""
                MskFec.Text = ""
                TxtCli.Text = ""
                TxtUsu.Text = ""
                MskIva.Text = 0
                MskSubTot.Text = 0
                MskTot.Text = 0
            End If
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Llena_CamposDetalle()
On Error Resume Next
            If RDetalle.RecordCount > 0 Then
                If IsNull(RDetalle!Documento) Then
                    TxtDocDet.Text = ""
                Else
                    TxtDocDet.Text = RDetalle!Documento
                End If
                If IsNull(RDetalle!Codigo) Then
                    TxtCod.Text = ""
                Else
                    TxtCod.Text = RDetalle!Codigo
                End If
                If IsNull(RDetalle!CantidadPedido) Then
                    MskCanPed.Text = ""
                Else
                    MskCanPed.Text = RDetalle!CantidadPedido
                End If
                If IsNull(RDetalle!CantidadEntregada) Then
                    MskCanEnt.Text = 0
                Else
                    MskCanEnt.Text = RDetalle!CantidadEntregada
                End If
                If IsNull(RDetalle!SaldoPorEntregar) Then
                    MskSal.Text = 0
                Else
                    MskSal.Text = RDetalle!SaldoPorEntregar
                End If
                If IsNull(RDetalle!DiasPedido) Then
                    TxtDiaPla.Text = 0
                Else
                    TxtDiaPla.Text = RDetalle!DiasPedido
                End If
                If IsNull(RDetalle!FechaParaEntregar) Then
                    MskFecEnt.Text = 0
                Else
                    MskFecEnt.Text = RDetalle!FechaParaEntregar
                End If
                If IsNull(RDetalle!FechaEntregaTotal) Then
                    MskFecEntTot.Text = ""
                Else
                    MskFecEntTot.Text = RDetalle!FechaEntregaTotal
                End If
                If IsNull(RDetalle!DiasDeAtraso) Then
                    TxtDiaAtr.Text = ""
                Else
                    TxtDiaAtr.Text = RDetalle!DiasDeAtraso
                End If
                If IsNull(RDetalle!Precio) Then
                    MskPre.Text = "0"
                Else
                    MskPre.Text = RDetalle!Precio
                End If
            Else
                Limpia_CamposDetalle
            End If
            
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Limpia_CamposEncabezado()
                TxtDoc.Text = ""
                TxtCli.Text = ""
                MskFec.Text = ""
                TxtObs.Text = ""
                TxtUsu.Text = ""
                MskIva.Text = 0
                MskSubTot.Text = 0
                MskTot.Text = 0
End Sub

Public Sub Limpia_CamposDetalle()
                TxtDocDet.Text = ""
                TxtCod.Text = ""
                MskCanPed.Text = 0
                MskCanEnt.Text = 0
                MskSal.Text = 0
                TxtDiaPla.Text = 0
                MskFecEnt.Text = ""
                MskFecEntTot.Text = ""
                TxtDiaAtr.Text = 0
                MskPre.Text = "0"
End Sub




