VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReportesPedidos 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes De Pedidos"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11865
   Icon            =   "ReportesPedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   11865
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
      Height          =   6015
      Left            =   120
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   11655
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4815
         Left            =   120
         TabIndex        =   34
         Top             =   1080
         Width           =   11415
         _ExtentX        =   20135
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
         Left            =   10680
         Picture         =   "ReportesPedidos.frx":5C12
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   33
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   3975
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
         TabIndex        =   31
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
         TabIndex        =   32
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10080
      Picture         =   "ReportesPedidos.frx":7C84
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   10080
      Picture         =   "ReportesPedidos.frx":85B6
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   120
      Width           =   1695
   End
   Begin TabDlg.SSTab TabReportes 
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   10610
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   1058
      BackColor       =   49152
      TabCaption(0)   =   "Pedidos Proveedores"
      TabPicture(0)   =   "ReportesPedidos.frx":8D00
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblpedprofecfin"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblpedprofecini"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblPedProDes"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblPedProEti"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "DTPPedProFecIni"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "DTPPedProFecFin"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "OptPedPro(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "OptPedPro(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "OptPedPro(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FramePedPro"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtPedPro"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FramePedProOpc"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "OptPedPro(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "FramePedPro3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "OptPedPro(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Pedidos Clientes"
      TabPicture(1)   =   "ReportesPedidos.frx":95DA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblPedCliDes"
      Tab(1).Control(1)=   "LblPedCliEti"
      Tab(1).Control(2)=   "lblpedclifecini"
      Tab(1).Control(3)=   "lblpedclifecFin"
      Tab(1).Control(4)=   "DTPPedCliFecIni"
      Tab(1).Control(5)=   "DTPPedCliFecFin"
      Tab(1).Control(6)=   "OptPedCli(2)"
      Tab(1).Control(7)=   "OptPedCli(1)"
      Tab(1).Control(8)=   "OptPedCli(0)"
      Tab(1).Control(9)=   "TxtPedCli"
      Tab(1).Control(10)=   "FramePedCli"
      Tab(1).Control(11)=   "FramePedCliOpc"
      Tab(1).Control(12)=   "OptPedCli(3)"
      Tab(1).Control(13)=   "FramePedCli3"
      Tab(1).Control(14)=   "OptPedCli(4)"
      Tab(1).ControlCount=   15
      TabCaption(2)   =   "Cierre Pedidos Proveedores"
      TabPicture(2)   =   "ReportesPedidos.frx":9EB4
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "OptCiePedPro(3)"
      Tab(2).Control(1)=   "TxtCiePedPro"
      Tab(2).Control(2)=   "FrameCiePedPro"
      Tab(2).Control(3)=   "OptCiePedPro(2)"
      Tab(2).Control(4)=   "OptCiePedPro(1)"
      Tab(2).Control(5)=   "OptCiePedPro(0)"
      Tab(2).Control(6)=   "DTPCiePedProFecFin"
      Tab(2).Control(7)=   "DTPCiePedProFecIni"
      Tab(2).Control(8)=   "LblCiePedProFecFin"
      Tab(2).Control(9)=   "LblCiePedProFecIni"
      Tab(2).Control(10)=   "LblCiePedProDes"
      Tab(2).Control(11)=   "LblCiePedProEti"
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "Cierre Pedidos Clientes"
      TabPicture(3)   =   "ReportesPedidos.frx":B646
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LblCiePedCliEti"
      Tab(3).Control(1)=   "LblCiePedCliDes"
      Tab(3).Control(2)=   "LblCiePedCliFecFin"
      Tab(3).Control(3)=   "LblCiePedCliFecIni"
      Tab(3).Control(4)=   "DTPCiePedCliFecIni"
      Tab(3).Control(5)=   "DTPCiePedCliFecFin"
      Tab(3).Control(6)=   "OptCiePedCli(0)"
      Tab(3).Control(7)=   "OptCiePedCli(1)"
      Tab(3).Control(8)=   "OptCiePedCli(2)"
      Tab(3).Control(9)=   "FrameCiePedCli"
      Tab(3).Control(10)=   "TxtCiePedCli"
      Tab(3).Control(11)=   "OptCiePedCli(3)"
      Tab(3).ControlCount=   12
      Begin VB.OptionButton OptPedCli 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   4
         Left            =   -74520
         TabIndex        =   87
         Top             =   2160
         Width           =   1455
      End
      Begin VB.Frame FramePedCli3 
         Caption         =   "Opciones De Seleccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   -67800
         TabIndex        =   83
         Top             =   960
         Width           =   2415
         Begin VB.OptionButton OptPedCli3 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   86
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptPedCli3 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   85
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton OptPedCli3 
            Caption         =   "Igual"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   84
            Top             =   1080
            Width           =   735
         End
      End
      Begin VB.OptionButton OptPedPro 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   4
         Left            =   480
         TabIndex        =   82
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Frame FramePedPro3 
         Caption         =   "Opciones De Seleccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   7200
         TabIndex        =   78
         Top             =   960
         Width           =   2415
         Begin VB.OptionButton OptPedPro3 
            Caption         =   "Igual"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   81
            Top             =   1080
            Width           =   735
         End
         Begin VB.OptionButton OptPedPro3 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   80
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton OptPedPro3 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   79
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.OptionButton OptCiePedCli 
         Caption         =   "No. Documento"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   49
         Top             =   2160
         Width           =   2055
      End
      Begin VB.OptionButton OptCiePedPro 
         Caption         =   "No. Documento"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   39
         Top             =   2160
         Width           =   1575
      End
      Begin VB.OptionButton OptPedCli 
         Caption         =   "No. Pedido"
         Height          =   195
         Index           =   3
         Left            =   -74520
         TabIndex        =   75
         Top             =   2520
         Width           =   1455
      End
      Begin VB.OptionButton OptPedPro 
         Caption         =   "No. Pedido"
         Height          =   195
         Index           =   3
         Left            =   480
         TabIndex        =   74
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Frame FramePedCliOpc 
         Caption         =   "Opciones De Reporte"
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
         Left            =   -72720
         TabIndex        =   69
         Top             =   960
         Width           =   2415
         Begin VB.OptionButton OptPedCliOpcRep 
            Caption         =   "Fechas Entrega"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   73
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton OptPedCliOpcRep 
            Caption         =   "Pedidos Entregados"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   72
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton OptPedCliOpcRep 
            Caption         =   "Pedidos Pendientes"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   71
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptPedCliOpcRep 
            Caption         =   "Fechas Pedido"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   70
            Top             =   1080
            Width           =   1575
         End
      End
      Begin VB.Frame FramePedProOpc 
         Caption         =   "Opciones De Reporte"
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
         Left            =   2280
         TabIndex        =   64
         Top             =   960
         Width           =   2415
         Begin VB.OptionButton OptPedProOpcRep 
            Caption         =   "Fechas Entrega"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   68
            Top             =   1440
            Width           =   1815
         End
         Begin VB.OptionButton OptPedProOpcRep 
            Caption         =   "Pedidos Entregados"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton OptPedProOpcRep 
            Caption         =   "Pedidos Pendientes"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   66
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptPedProOpcRep 
            Caption         =   "Fechas Pedido"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   65
            Top             =   1080
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtCiePedCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73080
         TabIndex        =   53
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Frame FrameCiePedCli 
         Caption         =   "Tipo De Reporte"
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
         Height          =   855
         Left            =   -68160
         TabIndex        =   50
         Top             =   1080
         Width           =   2535
         Begin VB.OptionButton OptCiePedCliRes 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   120
            TabIndex        =   52
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptCiePedCliDet 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   1440
            TabIndex        =   51
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.OptionButton OptCiePedCli 
         Caption         =   "Fechas Y Codigo"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   48
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton OptCiePedCli 
         Caption         =   "Pedido"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   47
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton OptCiePedCli 
         Caption         =   "Fechas "
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   46
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtCiePedPro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73080
         TabIndex        =   41
         Top             =   4560
         Width           =   1815
      End
      Begin VB.Frame FrameCiePedPro 
         Caption         =   "Tipo De Reporte"
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
         Height          =   855
         Left            =   -68160
         TabIndex        =   40
         Top             =   1080
         Width           =   2535
         Begin VB.OptionButton OptCiePedProDet 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   1440
            TabIndex        =   45
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton OptCiePedProRes 
            Caption         =   "Resumen"
            Height          =   195
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.OptionButton OptCiePedPro 
         Caption         =   "Fechas Y Codigo"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   38
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton OptCiePedPro 
         Caption         =   "Pedido"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   37
         Top             =   1440
         Width           =   1935
      End
      Begin VB.OptionButton OptCiePedPro 
         Caption         =   "Fechas "
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   36
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Frame FramePedCli 
         Caption         =   "Tipo De Reporte"
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
         Height          =   1815
         Left            =   -70200
         TabIndex        =   20
         Top             =   960
         Width           =   2295
         Begin VB.OptionButton OptPedCliTipRep 
            Caption         =   "Resumen x Producto"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   89
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptPedCliTipRep 
            Caption         =   "Resumen Cuadricula"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   77
            Top             =   1440
            Width           =   1935
         End
         Begin VB.OptionButton OptPedCliTipRep 
            Caption         =   "Resumen x No. Pedido"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   720
            Width           =   1935
         End
         Begin VB.OptionButton OptPedCliTipRep 
            Caption         =   "Detalle x No. Pedido"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   1935
         End
      End
      Begin VB.TextBox TxtPedCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         MaxLength       =   15
         TabIndex        =   18
         ToolTipText     =   "sigon '+' o doble click para ayuda"
         Top             =   5160
         Width           =   1695
      End
      Begin VB.OptionButton OptPedCli 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74520
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton OptPedCli 
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   15
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton OptPedCli 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   2
         Left            =   -74520
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
      End
      Begin VB.TextBox TxtPedPro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3240
         MaxLength       =   15
         TabIndex        =   5
         ToolTipText     =   "sigon '+' o doble click para ayuda"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame FramePedPro 
         Caption         =   "Tipo De Reporte"
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
         Height          =   1815
         Left            =   4800
         TabIndex        =   7
         Top             =   960
         Width           =   2295
         Begin VB.OptionButton OptPedProTipRep 
            Caption         =   "Resumen x Producto"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   88
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OptPedProTipRep 
            Caption         =   "Resumen Cuadricula"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   76
            Top             =   1440
            Width           =   1815
         End
         Begin VB.OptionButton OptPedProTipRep 
            Caption         =   "Resumen x No. Pedido"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton OptPedProTipRep 
            Caption         =   "Detalle x No. Pedido"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   1935
         End
      End
      Begin VB.OptionButton OptPedPro 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   1080
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptPedPro 
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.OptionButton OptPedPro 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   2
         Left            =   480
         TabIndex        =   3
         Top             =   1800
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPPedProFecFin 
         Height          =   255
         Left            =   8040
         TabIndex        =   13
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37603
      End
      Begin MSComCtl2.DTPicker DTPPedProFecIni 
         Height          =   255
         Left            =   5520
         TabIndex        =   11
         Top             =   4200
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37603
      End
      Begin MSComCtl2.DTPicker DTPPedCliFecFin 
         Height          =   255
         Left            =   -66960
         TabIndex        =   26
         Top             =   4440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37603
      End
      Begin MSComCtl2.DTPicker DTPPedCliFecIni 
         Height          =   255
         Left            =   -69480
         TabIndex        =   24
         Top             =   4440
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37603
      End
      Begin MSComCtl2.DTPicker DTPCiePedProFecFin 
         Height          =   255
         Left            =   -66960
         TabIndex        =   56
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37603
      End
      Begin MSComCtl2.DTPicker DTPCiePedProFecIni 
         Height          =   255
         Left            =   -69480
         TabIndex        =   57
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37603
      End
      Begin MSComCtl2.DTPicker DTPCiePedCliFecFin 
         Height          =   255
         Left            =   -66960
         TabIndex        =   60
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37603
      End
      Begin MSComCtl2.DTPicker DTPCiePedCliFecIni 
         Height          =   255
         Left            =   -69480
         TabIndex        =   61
         Top             =   3120
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37603
      End
      Begin VB.Label LblCiePedCliFecIni 
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
         Left            =   -70200
         TabIndex        =   63
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label LblCiePedCliFecFin 
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
         Left            =   -67560
         TabIndex        =   62
         Top             =   3120
         Width           =   510
      End
      Begin VB.Label LblCiePedProFecFin 
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
         Left            =   -67560
         TabIndex        =   59
         Top             =   3120
         Width           =   510
      End
      Begin VB.Label LblCiePedProFecIni 
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
         Left            =   -70200
         TabIndex        =   58
         Top             =   3120
         Width           =   555
      End
      Begin VB.Label LblCiePedCliDes 
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
         Left            =   -71160
         TabIndex        =   55
         Top             =   4560
         Width           =   5895
      End
      Begin VB.Label LblCiePedCliEti 
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
         Left            =   -74880
         TabIndex        =   54
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label LblCiePedProDes 
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
         Left            =   -71160
         TabIndex        =   43
         Top             =   4560
         Width           =   5895
      End
      Begin VB.Label LblCiePedProEti 
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
         Left            =   -74880
         TabIndex        =   42
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Label lblpedclifecFin 
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
         Left            =   -67560
         TabIndex        =   25
         Top             =   4440
         Width           =   510
      End
      Begin VB.Label lblpedclifecini 
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
         Left            =   -70080
         TabIndex        =   23
         Top             =   4440
         Width           =   555
      End
      Begin VB.Label LblPedCliEti 
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
         Left            =   -74520
         TabIndex        =   17
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Label LblPedCliDes 
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
         Left            =   -69960
         TabIndex        =   19
         Top             =   5160
         Width           =   4575
      End
      Begin VB.Label LblPedProEti 
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
         Left            =   480
         TabIndex        =   4
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Label LblPedProDes 
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
         Left            =   5040
         TabIndex        =   6
         Top             =   5160
         Width           =   4335
      End
      Begin VB.Label lblpedprofecini 
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
         Left            =   4920
         TabIndex        =   10
         Top             =   4200
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label lblpedprofecfin 
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
         Left            =   7440
         TabIndex        =   12
         Top             =   4200
         Visible         =   0   'False
         Width           =   510
      End
   End
End
Attribute VB_Name = "ReportesPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaProveedor As New ADODB.Recordset
Dim RBuscaCliente As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VDia, VMes, VAño, VDia2, VMes2, VAño2 As String
Dim VTexto As String

'VARIABLES PARA CARPETA DE PEDIDOS DE Descripcion
Dim BPedProCodPro As Boolean
Dim BPedProCodArt As Boolean

'VARIABLES PARA CARPETA DE PEDIDOS DE CLIENTES
Dim BPedCliCodCli As Boolean
Dim BPedCliCodArt As Boolean

'VARIABLES PARA CARPETA DE CIERRE PEDIDOS DE Descripcion
Dim BCiePedProCodPro As Boolean
Dim BCiePedProCodArt As Boolean

'VARIABLES PARA CARPETA DE CIERRE PEDIDOS DE CLIENTES
Dim BCiePedCliCodCli As Boolean
Dim BCiePedCliCodArt As Boolean




Private Sub CmdImprimir_Click()
On Error Resume Next
    MousePointer = 11
        
    If TabReportes.Tab = 0 Then
            'VA AL PROCEDIMIENTO DE PROVEEDORES
            PedidosProveedor
    ElseIf TabReportes.Tab = 1 Then
            'VA AL PROCEDIMIENTO DE CLIENTES
            PedidosClientes
    ElseIf TabReportes.Tab = 2 Then
            'VA AL PROCEDIMIENTO DE PROVEEDORES
            CierrePedidosProveedor
    ElseIf TabReportes.Tab = 3 Then
            'VA AL PROCEDIMIENTO DE CLIENTES
            CierrePedidosCliente
    End If
    
             'DESPLIEGA EL REPORTE
             FrmReporte.Show
             
             MousePointer = 0
             If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Exit Sub
             End If
                
    

End Sub
Public Sub PedidosProveedor()
            VDia = Day(DTPPedProFecIni.Value)
            VMes = Month(DTPPedProFecIni.Value)
            VAño = Year(DTPPedProFecIni.Value)
            VDia2 = Day(DTPPedProFecFin.Value)
            VMes2 = Month(DTPPedProFecFin.Value)
            VAño2 = Year(DTPPedProFecFin.Value)
            
            If OptPedPro3.Item(0).Value = True Then
                    VTexto = "Like '" & TxtPedPro.Text & "*'"
            ElseIf OptPedPro3.Item(1).Value = True Then
                    VTexto = "Like '*" & TxtPedPro.Text & "*'"
            ElseIf OptPedPro3.Item(2).Value = True Then
                    VTexto = "= '*" & TxtPedPro.Text & "'"
            End If
            
                    'FECHAS _____________________________________________________________________________
                    If OptPedPro.Item(0).Value = True Then
                                'FECHAS DE PEDIDOS
                                If OptPedProOpcRep.Item(0).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{EncabezadoPedidosProveedores.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                'PEDIDOS PENDIENTES
                                ElseIf OptPedProOpcRep.Item(1).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy") & " Que Estan Pendientes "
                                    'GCriteriaReporte = "{EncabezadoPedidosProveedores.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {DetallePedidosProveedores.SaldoPorEntregar} > 0"
                                    GCriteriaReporte = "{DetallePedidosProveedores.SaldoPorEntregar} > 0"
                                'PEDIDOS ENTREGADOS
                                ElseIf OptPedProOpcRep.Item(2).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy") & " Que Estan Entregados "
                                    GCriteriaReporte = "{DetallePedidosProveedores.SaldoPorEntregar} <= 0"
                                'FECHAS DE ENTREGA
                                ElseIf OptPedProOpcRep.Item(3).Value = True Then
                                    GTituloReporte = "Fechas De Entrega De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{DetallePedidosProveedores.FechaParaEntregar} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                End If
                    'Descripcion___________________________________________________________________________
                    ElseIf OptPedPro.Item(1).Value = True Then
                                'FECHAS DE PEDIDOS
                                If OptPedProOpcRep.Item(0).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{EncabezadoPedidosProveedores.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {EncabezadoPedidosProveedores.Proveedor} " & VTexto
                                'PEDIDOS PENDIENTES
                                ElseIf OptPedProOpcRep.Item(1).Value = True Then
                                    GTituloReporte = "Pedidos Pendientes De Descripcion " & LblPedProDes.Caption
                                    GCriteriaReporte = "{DetallePedidosProveedores.SaldoPorEntregar} > 0" & " And {EncabezadoPedidosProveedores.Proveedor} " & VTexto
                                'PEDIDOS ENTREGADOS
                                ElseIf OptPedProOpcRep.Item(2).Value = True Then
                                    GTituloReporte = "Pedidos Entregados De Descripcion " & LblPedProDes.Caption
                                    GCriteriaReporte = "{DetallePedidosProveedores.SaldoPorEntregar} <= 0" & " And {EncabezadoPedidosProveedores.Proveedor} " & VTexto
                                'FECHAS DE ENTREGA
                                ElseIf OptPedProOpcRep.Item(3).Value = True Then
                                    GTituloReporte = "Fechas De Entrega De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{DetallePedidosProveedores.FechaParaEntregar} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {EncabezadoPedidosProveedores.Proveedor} " & VTexto
                                End If
                    'CODIGO________________________________________________________________________
                    ElseIf OptPedPro.Item(2).Value = True Then
                                'FECHAS DE PEDIDOS
                                If OptPedProOpcRep.Item(0).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{EncabezadoPedidosProveedores.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {DetallePedidosProveedores.Codigo} " & VTexto
                                'PEDIDOS PENDIENTES
                                ElseIf OptPedProOpcRep.Item(1).Value = True Then
                                    GTituloReporte = "Pedidos Pendientes De Ficha Tecnica " & LblPedProDes.Caption
                                    GCriteriaReporte = "{DetallePedidosProveedores.SaldoPorEntregar} > 0" & " And {DetallePedidosProveedores.Codigo} " & VTexto
                                'PEDIDOS ENTREGADOS
                                ElseIf OptPedProOpcRep.Item(2).Value = True Then
                                    GTituloReporte = "Pedidos Entregados De Ficha Tecnica " & LblPedProDes.Caption
                                    GCriteriaReporte = "{DetallePedidosProveedores.SaldoPorEntregar} <= 0" & " And {DetallePedidosProveedores.Codigo} " & VTexto
                                'FECHAS DE ENTREGA
                                ElseIf OptPedProOpcRep.Item(3).Value = True Then
                                    GTituloReporte = "Fechas De Entrega De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{DetallePedidosProveedores.FechaParaEntregar} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {DetallePedidosProveedores.Codigo} " & VTexto
                                End If
                    'DESCRIPCION __________________________________________________________________
                    ElseIf OptPedPro.Item(4).Value = True Then
                                'FECHAS DE PEDIDOS
                                If OptPedProOpcRep.Item(0).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{EncabezadoPedidosProveedores.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {CorrelativosMateriaPrima.Descripcion} " & VTexto
                                'PEDIDOS PENDIENTES
                                ElseIf OptPedProOpcRep.Item(1).Value = True Then
                                    GTituloReporte = "Pedidos Pendientes De Ficha Tecnica " & LblPedProDes.Caption
                                    GCriteriaReporte = "{DetallePedidosProveedores.SaldoPorEntregar} > 0" & " And {CorrelativosMateriaPrima.Descripcion} " & VTexto
                                'PEDIDOS ENTREGADOS
                                ElseIf OptPedProOpcRep.Item(2).Value = True Then
                                    GTituloReporte = "Pedidos Entregados De Ficha Tecnica " & LblPedProDes.Caption
                                    GCriteriaReporte = "{DetallePedidosProveedores.SaldoPorEntregar} <= 0" & " And {CorrelativosMateriaPrima.Descripcion} " & VTexto
                                'FECHAS DE ENTREGA
                                ElseIf OptPedProOpcRep.Item(3).Value = True Then
                                    GTituloReporte = "Fechas De Entrega De " & Format(DTPPedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedProFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{DetallePedidosProveedores.FechaParaEntregar} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {CorrelativosMateriaPrima.Descripcion} " & VTexto
                                End If
                    
                    'No. PEDIDO __________________________________________________________________
                    ElseIf OptPedPro.Item(3).Value = True Then
                                    GTituloReporte = "Por Numero De Pedido"
                                    GCriteriaReporte = "{EncabezadoPedidosProveedores.Documento} " & VTexto
                    
                    End If
                                                          
                
                'ELIGE REPORTE DE ACUERDO A LA OPCION
                If OptPedProTipRep.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "PedidosProveedoresResumen.rpt"
                    Else
                        GNombreReporte = "PedidosProveedoresResumenO.rpt"
                    End If
                ElseIf OptPedProTipRep.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "PedidosProveedoresDetalle.rpt"
                    Else
                        GNombreReporte = "PedidosProveedoresDetalleO.rpt"
                    End If
                ElseIf OptPedProTipRep.Item(2).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "PedidosProveedoresResumenCuadricula.rpt"
                    Else
                        GNombreReporte = "PedidosProveedoresResumenCuadriculaO.rpt"
                    End If
                ElseIf OptPedProTipRep.Item(3).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "PedidosProveedoresResumenxProducto.rpt"
                    Else
                        GNombreReporte = "PedidosProveedoresResumenxProductoO.rpt"
                    End If
                End If
    
End Sub


Public Sub PedidosClientes()
            VDia = Day(DTPPedCliFecIni.Value)
            VMes = Month(DTPPedCliFecIni.Value)
            VAño = Year(DTPPedCliFecIni.Value)
            VDia2 = Day(DTPPedCliFecFin.Value)
            VMes2 = Month(DTPPedCliFecFin.Value)
            VAño2 = Year(DTPPedCliFecFin.Value)
            
            If OptPedCli3.Item(0).Value = True Then
                    VTexto = "Like '" & TxtPedCli.Text & "*'"
            ElseIf OptPedCli3.Item(1).Value = True Then
                    VTexto = "Like '*" & TxtPedCli.Text & "*'"
            ElseIf OptPedCli3.Item(2).Value = True Then
                    VTexto = "= '*" & TxtPedCli.Text & "'"
            End If

                    
                    'FECHAS _____________________________________________________________________________
                    If OptPedCli.Item(0).Value = True Then
                                'FECHAS DE PEDIDOS
                                If OptPedCliOpcRep.Item(0).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{EncabezadoPedidosClientes.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                'PEDIDOS PENDIENTES
                                ElseIf OptPedCliOpcRep.Item(1).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy") & " Que Estan Pendientes "
                                    GCriteriaReporte = "{DetallepedidosClientes.SaldoPorEntregar} > 0"
                                'PEDIDOS ENTREGADOS
                                ElseIf OptPedCliOpcRep.Item(2).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy") & " Que Estan Entregados"
                                    GCriteriaReporte = "{DetallepedidosClientes.SaldoPorEntregar} <= 0"
                                'FECHAS DE ENTREGA
                                ElseIf OptPedCliOpcRep.Item(3).Value = True Then
                                    GTituloReporte = "Fechas De Entrega De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{DetallepedidosClientes.FechaParaEntregar} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                End If
                    'Descripcion___________________________________________________________________________
                    ElseIf OptPedCli.Item(1).Value = True Then
                                'FECHAS DE PEDIDOS
                                If OptPedCliOpcRep.Item(0).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{EncabezadoPedidosClientes.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {EncabezadoPedidosClientes.Cliente} " & VTexto
                                'PEDIDOS PENDIENTES
                                ElseIf OptPedCliOpcRep.Item(1).Value = True Then
                                    GTituloReporte = "Pedidos Pendientes De Cliente " & LblPedCliDes.Caption
                                    GCriteriaReporte = "{DetallepedidosClientes.SaldoPorEntregar} > 0" & " And {EncabezadoPedidosClientes.Cliente} " & VTexto
                                'PEDIDOS ENTREGADOS
                                ElseIf OptPedCliOpcRep.Item(2).Value = True Then
                                    GTituloReporte = "Pedidos Entregados De Cliente " & LblPedCliDes.Caption
                                    GCriteriaReporte = "{DetallepedidosClientes.SaldoPorEntregar} <= 0" & " And {EncabezadoPedidosClientes.Cliente} " & VTexto
                                'FECHAS DE ENTREGA
                                ElseIf OptPedCliOpcRep.Item(3).Value = True Then
                                    GTituloReporte = "Fechas De Entrega De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{DetallepedidosClientes.FechaParaEntregar} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {EncabezadoPedidosClientes.Cliente} " & VTexto
                                End If
                    'CODIGO ________________________________________________________________________
                    ElseIf OptPedCli.Item(2).Value = True Then
                                'FECHAS DE PEDIDOS
                                If OptPedCliOpcRep.Item(0).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{EncabezadoPedidosClientes.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {DetallepedidosClientes.Codigo} " & VTexto
                                'PEDIDOS PENDIENTES
                                ElseIf OptPedCliOpcRep.Item(1).Value = True Then
                                    GTituloReporte = "Pedidos Pendientes De Ficha Tecnica " & LblPedCliDes.Caption
                                    GCriteriaReporte = "{DetallepedidosClientes.SaldoPorEntregar} > 0" & " And {DetallepedidosClientes.Codigo} " & VTexto
                                'PEDIDOS ENTREGADOS
                                ElseIf OptPedCliOpcRep.Item(2).Value = True Then
                                    GTituloReporte = "Pedidos Entregados De Ficha Tecnica " & LblPedCliDes.Caption
                                    GCriteriaReporte = "{DetallepedidosClientes.SaldoPorEntregar} <= 0" & " And {DetallepedidosClientes.Codigo} " & VTexto
                                'FECHAS DE ENTREGA
                                ElseIf OptPedCliOpcRep.Item(3).Value = True Then
                                    GTituloReporte = "Fechas De Entrega De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{DetallepedidosClientes.FechaParaEntregar} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {DetallepedidosClientes.Codigo} " & VTexto
                                End If
                    'DESCRIPCION____________________________________________________________________
                    ElseIf OptPedCli.Item(4).Value = True Then
                                'FECHAS DE PEDIDOS
                                If OptPedCliOpcRep.Item(0).Value = True Then
                                    GTituloReporte = "Fechas De Pedido De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{EncabezadoPedidosClientes.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {CorrelativosMateriaPrima.Descripcion} " & VTexto
                                'PEDIDOS PENDIENTES
                                ElseIf OptPedCliOpcRep.Item(1).Value = True Then
                                    GTituloReporte = "Pedidos Pendientes De Ficha Tecnica " & LblPedCliDes.Caption
                                    GCriteriaReporte = "{DetallepedidosClientes.SaldoPorEntregar} > 0" & " And {CorrelativosMateriaPrima.Descripcion} " & VTexto
                                'PEDIDOS ENTREGADOS
                                ElseIf OptPedCliOpcRep.Item(2).Value = True Then
                                    GTituloReporte = "Pedidos Entregados De Ficha Tecnica " & LblPedCliDes.Caption
                                    GCriteriaReporte = "{DetallepedidosClientes.SaldoPorEntregar} <= 0" & " And {CorrelativosMateriaPrima.Descripcion} " & VTexto
                                'FECHAS DE ENTREGA
                                ElseIf OptPedCliOpcRep.Item(3).Value = True Then
                                    GTituloReporte = "Fechas De Entrega De " & Format(DTPPedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPPedCliFecFin.Value, "dd/mm/yyyy")
                                    GCriteriaReporte = "{DetallepedidosClientes.FechaParaEntregar} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")" & " And {CorrelativosMateriaPrima.Descripcion} " & VTexto
                                End If
                    
                    'No. PEDIDO _________________________________________________________________
                    ElseIf OptPedCli.Item(3).Value = True Then
                                    GTituloReporte = " No. Pedido"
                                    GCriteriaReporte = "{EncabezadoPedidosClientes.Documento} = '" & TxtPedCli.Text & "'"
                    End If
            
                    
                'ELIGE REPORTE DE ACUERDO A LA OPCION
                If OptPedCliTipRep.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "PedidosClientesResumen.rpt"
                    Else
                        GNombreReporte = "PedidosClientesResumenO.rpt"
                    End If
                ElseIf OptPedCliTipRep.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "PedidosClientesDetalle.rpt"
                    Else
                        GNombreReporte = "PedidosClientesDetalleO.rpt"
                    End If
                ElseIf OptPedCliTipRep.Item(2).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "PedidosClientesResumenCuadricula.rpt"
                    Else
                        GNombreReporte = "PedidosClientesResumenCuadriculaO.rpt"
                    End If
                ElseIf OptPedCliTipRep.Item(3).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "PedidosClientesResumenxProducto.rpt"
                    Else
                        GNombreReporte = "PedidosClientesResumenxProductoO.rpt"
                    End If
                End If
End Sub


Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
        Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
                'MATERIA PRIMA EN PEDIDOS DE Descripcion
                If BPedProCodArt = True Then
                        TxtPedPro.Text = DBGridBusqueda.Columns(0).Text
                        TxtPedPro.SetFocus
                'MATERIA PRIMA EN PEDIDOS DE CLIENTES
                ElseIf BPedCliCodArt = True Then
                        TxtPedCli.Text = DBGridBusqueda.Columns(0).Text
                        TxtPedCli.SetFocus
                'Descripcion EN CARPETA DE PEDIDOS DE PROVEEDORES
                ElseIf BPedProCodPro = True Then
                        TxtPedPro.Text = DBGridBusqueda.Columns(0).Text
                        TxtPedPro.SetFocus
                'CLIENTE EN CARPETA DE PEDIDOS DE CLIENTES
                ElseIf BPedCliCodCli = True Then
                        TxtPedCli.Text = DBGridBusqueda.Columns(0).Text
                        TxtPedCli.SetFocus
                'MATERIA PRIMA EN CIERRE PEDIDOS DE Descripcion
                ElseIf BCiePedProCodArt = True Then
                        TxtCiePedPro.Text = DBGridBusqueda.Columns(0).Text
                        TxtCiePedPro.SetFocus
                'MATERIA PRIMA EN CIERRE PEDIDOS DE CLIENTES
                ElseIf BCiePedCliCodArt = True Then
                        TxtCiePedCli.Text = DBGridBusqueda.Columns(0).Text
                        TxtCiePedCli.SetFocus
                'Descripcion EN CARPETA DE CIERRE PEDIDOS DE PROVEEDORES
                ElseIf BCiePedProCodPro = True Then
                        TxtCiePedPro.Text = DBGridBusqueda.Columns(0).Text
                        TxtCiePedPro.SetFocus
                'CLIENTE EN CARPETA DE CIERRE PEDIDOS DE CLIENTES
                ElseIf BCiePedCliCodCli = True Then
                        TxtCiePedCli.Text = DBGridBusqueda.Columns(0).Text
                        TxtCiePedCli.SetFocus
                End If
                        FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        'SI PRECIONA LA TECLA DEL SIGNO '+'
        If KeyAscii = 43 Then
                'MATERIA PRIMA EN PEDIDOS DE Descripcion
                If BPedProCodArt = True Then
                        TxtPedPro.Text = DBGridBusqueda.Columns(0).Text
                        TxtPedPro.SetFocus
                'MATERIA PRIMA EN PEDIDOS DE CLIENTES
                ElseIf BPedCliCodArt = True Then
                        TxtPedCli.Text = DBGridBusqueda.Columns(0).Text
                        TxtPedCli.SetFocus
                'Descripcion EN CARPETA DE PEDIDOS DE PROVEEDORES
                ElseIf BPedProCodPro = True Then
                        TxtPedPro.Text = DBGridBusqueda.Columns(0).Text
                        TxtPedPro.SetFocus
                'CLIENTE EN CARPETA DE PEDIDOS DE CLIENTES
                ElseIf BPedCliCodCli = True Then
                        TxtPedCli.Text = DBGridBusqueda.Columns(0).Text
                        TxtPedCli.SetFocus
                'MATERIA PRIMA EN CIERRE PEDIDOS DE Descripcion
                ElseIf BCiePedProCodArt = True Then
                        TxtCiePedPro.Text = DBGridBusqueda.Columns(0).Text
                        TxtCiePedPro.SetFocus
                'MATERIA PRIMA EN CIERRE PEDIDOS DE CLIENTES
                ElseIf BCiePedCliCodArt = True Then
                        TxtCiePedCli.Text = DBGridBusqueda.Columns(0).Text
                        TxtCiePedCli.SetFocus
                'Descripcion EN CARPETA DE CIERRE PEDIDOS DE PROVEEDORES
                ElseIf BCiePedProCodPro = True Then
                        TxtCiePedPro.Text = DBGridBusqueda.Columns(0).Text
                        TxtCiePedPro.SetFocus
                'CLIENTE EN CARPETA DE CIERRE PEDIDOS DE CLIENTES
                ElseIf BCiePedCliCodCli = True Then
                        TxtCiePedCli.Text = DBGridBusqueda.Columns(0).Text
                        TxtCiePedCli.SetFocus
                End If
                        FrameBusqueda.Visible = False
        
        End If

End Sub

Private Sub Form_Load()
        'FECHAS DE TAB DE PEDIDOS DE Descripcion
        DTPPedProFecIni.Value = Date
        DTPPedProFecFin.Value = Date
                
        'FECHAS DE TAB DE PEDIDOS DE CLIENTE
        DTPPedCliFecIni.Value = Date
        DTPPedCliFecFin.Value = Date
        
        'FECHAS DE TAB DE PEDIDOS DE Descripcion
        DTPCiePedProFecIni = Date
        DTPCiePedProFecFin = Date
        
        'FECHAS DE TAB DE PEDIDOS DE CLIENTE
        DTPCiePedCliFecIni = Date
        DTPCiePedCliFecFin = Date
        
        
End Sub

Private Sub OptCiePedCli_Click(Index As Integer)
        If Index = 1 Then
            TxtCiePedCli.Visible = True
            LblCiePedCliEti.Caption = "Pedido"
            TxtCiePedCli.SetFocus
        ElseIf Index = 2 Then
            TxtCiePedCli.Visible = True
            LblCiePedCliEti.Caption = "Codigo Articulo"
            TxtCiePedCli.SetFocus
        ElseIf Index = 3 Then
            TxtCiePedCli.Visible = True
            LblCiePedCliEti.Caption = "No. Pedido"
            TxtCiePedCli.SetFocus
        Else
            LblCiePedCliEti.Caption = ""
            LblCiePedCliDes.Caption = ""
            TxtCiePedCli.Text = ""
            TxtCiePedCli.Visible = False
        End If
        
        If Index = 1 Then
            LblCiePedCliFecIni.Visible = False
            LblCiePedCliFecFin.Visible = False
            DTPCiePedCliFecIni.Visible = False
            DTPCiePedCliFecFin.Visible = False
        Else
            LblCiePedCliFecIni.Visible = True
            LblCiePedCliFecFin.Visible = True
            DTPCiePedCliFecIni.Visible = True
            DTPCiePedCliFecFin.Visible = True
        End If

End Sub

Private Sub OptCiePedPro_Click(Index As Integer)
        If Index = 1 Then
            TxtCiePedPro.Visible = True
            LblCiePedProEti.Caption = "Pedido"
            TxtCiePedPro.SetFocus
        ElseIf Index = 2 Then
            TxtCiePedPro.Visible = True
            LblCiePedProEti.Caption = "Codigo Articulo"
            TxtCiePedPro.SetFocus
        ElseIf Index = 3 Then
            TxtCiePedPro.Visible = True
            LblCiePedProEti.Caption = "No. Documento"
            TxtCiePedPro.SetFocus
        Else
            LblCiePedProEti.Caption = ""
            LblCiePedProDes.Caption = ""
            TxtCiePedPro.Text = ""
            TxtCiePedPro.Visible = False
        End If
        
        If Index = 1 Then
            LblCiePedProFecIni.Visible = False
            LblCiePedProFecFin.Visible = False
            DTPCiePedProFecIni.Visible = False
            DTPCiePedProFecFin.Visible = False
        Else
            LblCiePedProFecIni.Visible = True
            LblCiePedProFecFin.Visible = True
            DTPCiePedProFecIni.Visible = True
            DTPCiePedProFecFin.Visible = True
        End If

                
End Sub

Private Sub OptPedCli_Click(Index As Integer)
        If Index = 1 Then
            TxtPedCli.Visible = True
            LblPedCliEti.Caption = "Cliente"
            TxtPedCli.SetFocus
        ElseIf Index = 2 Then
            LblPedCliEti.Caption = "Codigo Producto"
            TxtPedCli.Visible = True
            TxtPedCli.SetFocus
        ElseIf Index = 4 Then
            LblPedCliEti.Caption = "Descripcion Producto"
            TxtPedCli.Visible = True
            TxtPedCli.SetFocus
        Else
            LblPedCliEti.Caption = ""
            LblPedProDes.Caption = ""
            TxtPedCli.Text = ""
            TxtPedCli.Visible = False
        End If
        
        'NUMERO DE PEDIDO
        If Index = 3 Then
            FramePedCliOpc.Visible = False
            DTPPedCliFecIni.Visible = False
            DTPPedCliFecFin.Visible = False
            lblpedclifecini.Visible = False
            lblpedclifecFin.Visible = False
            
            LblPedCliEti.Caption = "No. Pedido"
            LblPedCliDes.Caption = ""
            TxtPedCli.Text = ""
            TxtPedCli.Visible = True
            TxtPedCli.SetFocus
        Else
            FramePedCliOpc.Visible = True
            DTPPedCliFecIni.Visible = True
            DTPPedCliFecFin.Visible = True
            lblpedclifecini.Visible = True
            lblpedclifecFin.Visible = True
        End If
        
        

End Sub

Private Sub OptPedCliOpcRep_Click(Index As Integer)
        If (Index = 0) Or (Index = 3) Then
                DTPPedCliFecIni.Visible = True
                DTPPedCliFecFin.Visible = True
                lblpedclifecini.Visible = True
                lblpedclifecFin.Visible = True
        ElseIf (Index = 1 Or Index = 2) Then
                DTPPedCliFecIni.Visible = False
                DTPPedCliFecFin.Visible = False
                lblpedclifecini.Visible = False
                lblpedclifecFin.Visible = False
        End If
               
End Sub

Private Sub OptPedPro_Click(Index As Integer)
        If Index = 1 Then
            TxtPedPro.Visible = True
            LblPedProEti.Caption = "Descripcion"
            TxtPedPro.SetFocus
        ElseIf Index = 2 Then
            TxtPedPro.Visible = True
            LblPedProEti.Caption = "Codigo Producto"
            TxtPedPro.SetFocus
        ElseIf Index = 4 Then
            TxtPedPro.Visible = True
            LblPedProEti.Caption = "Descripcion Producto"
            TxtPedPro.SetFocus
        Else
            LblPedProEti.Caption = ""
            LblPedProDes.Caption = ""
            TxtPedPro.Text = ""
            TxtPedPro.Visible = False
        End If
        
        'NUMERO DE PEDIDO
        If Index = 3 Then
            FramePedProOpc.Visible = False
            DTPPedProFecIni.Visible = False
            DTPPedProFecFin.Visible = False
            lblpedprofecini.Visible = False
            lblpedprofecfin.Visible = False
            
            LblPedProEti.Caption = "No. Pedido"
            LblPedProDes.Caption = ""
            TxtPedPro.Text = ""
            TxtPedPro.Visible = True
            TxtPedPro.SetFocus
        Else
            FramePedProOpc.Visible = True
            DTPPedProFecIni.Visible = True
            DTPPedProFecFin.Visible = True
            lblpedprofecini.Visible = True
            lblpedprofecfin.Visible = True
        End If
        
        
        'OPCIONES DE REPORTE
        If (Index = 0) Or (Index = 3) Then
                DTPPedProFecIni.Visible = True
                DTPPedProFecFin.Visible = True
                lblpedprofecini.Visible = True
                lblpedprofecfin.Visible = True
        'PEDIDOS PENDIENTES
        ElseIf (Index = 1 Or Index = 2) Then
                DTPPedProFecIni.Visible = False
                DTPPedProFecFin.Visible = False
                lblpedprofecini.Visible = False
                lblpedprofecfin.Visible = False
        End If


End Sub


Private Sub OptPedProOpcRep_Click(Index As Integer)
        If (Index = 0) Or (Index = 3) Then
                DTPPedProFecIni.Visible = True
                DTPPedProFecFin.Visible = True
                lblpedprofecini.Visible = True
                lblpedprofecfin.Visible = True
        'PEDIDOS PENDIENTES
        ElseIf (Index = 1 Or Index = 2) Then
                DTPPedProFecIni.Visible = False
                DTPPedProFecFin.Visible = False
                lblpedprofecini.Visible = False
                lblpedprofecfin.Visible = False
        End If
        
        
End Sub

Private Sub TabReportes_Click(PreviousTab As Integer)
        If TabReportes.Tab = 0 Then
            OptPedPro.Item(0).Value = True
        ElseIf TabReportes.Tab = 1 Then
            OptPedCli.Item(0).Value = True
        ElseIf TabReportes.Tab = 2 Then
            OptCiePedPro.Item(0).Value = True
        ElseIf TabReportes.Tab = 3 Then
            OptCiePedCli.Item(0).Value = True
        End If
End Sub

Private Sub TxtBusqueda_Change()
        Set RBusqueda = New ADODB.Recordset
        
        'PROVEEDORES
        If (BPedProCodPro = True Or BCiePedProCodPro = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores Where CodigoProveedor Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores Where UPPER(CodigoProveedor) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
        'CLIENTES
        ElseIf (BPedCliCodCli = True Or BCiePedCliCodCli = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where CodigoCliente Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where UPPER(CodigoCliente) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
            
        'MATERIA PRIMA
        ElseIf (BPedProCodArt = True Or BPedCliCodArt = True Or BCiePedProCodArt = True Or BCiePedCliCodArt = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Descrip Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
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

Private Sub TxtCiePedCli_Change()
    'BUSCA CODIGO DE MATERIA PRIMA
    If (OptCiePedCli.Item(2).Value = True) Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCiePedCli.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCiePedCli.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblCiePedCliDes.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblCiePedCliDes.Caption = ""
                End If
    End If
    
    'BUSCA CLIENTE
  '  If (OptCiePedCli.Item(1).Value = True) Then
  '      Set RBuscaCliente = Db.OpenRecordset("Select Descripcion From Clientes Where CodigoCliente = '" & TxtCiePedCli.Text & "'")
  '              If RBuscaCliente.RecordCount > 0 Then
  '                  LblCiePedCliDes.Caption = RBuscaCliente!Descripcion
  '              Else
  '                  LblCiePedCliDes.Caption = ""
  '              End If
  '  End If

End Sub

Private Sub TxtCiePedCli_DblClick()
                Set RBusqueda = New ADODB.Recordset
                'OPCION POR CLIENTE
                'If OptCiePedCli.Item(1).Value = True Then
                '            BPedProCodPro = False
                '            BPedProCodArt = False
                '            BPedCliCodCli = False
                '            BPedCliCodArt = False
                '            BCiePedProCodPro = False
                '            BCiePedProCodArt = False
                '            BCiePedCliCodCli = True
                '            BCiePedCliCodArt = False
                'OPCION POR MATERIA PRIMA
                If OptCiePedCli.Item(2).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = False
                            BPedCliCodCli = False
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = True
                
            
                            'OPCION DE BODEGA
                            If BCiePedCliCodCli = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes")
                            'OPCION DE MATERIA PRIMA
                            ElseIf BCiePedCliCodArt = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                            End If
                                    
                                    Set DBGridBusqueda.DataSource = RBusqueda
                                    If BPedCliCodCli = True Then
                                        DBGridBusqueda.Columns(1).Width = "4000"
                                    ElseIf BPedCliCodArt = True Then
                                        DBGridBusqueda.Columns(1).Width = "4000"
                                    End If
                                    Columnas
                                    FrameBusqueda.Visible = True
                                    TxtBusqueda.SetFocus
                End If

End Sub

Private Sub TxtCiePedCli_GotFocus()
        TxtCiePedCli.SelStart = 0
        TxtCiePedCli.SelLength = Len(TxtCiePedCli.Text)
End Sub

Private Sub TxtCiePedCli_KeyPress(KeyAscii As Integer)
        
        If KeyAscii = 13 Then
            SendKeys "{TAB}"
        End If
            
        If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                'OPCION POR CLIENTE
                'If OptCiePedCli.Item(1).Value = True Then
                '            BPedProCodPro = False
                '            BPedProCodArt = False
                '            BPedCliCodCli = False
                '            BPedCliCodArt = False
                '            BCiePedProCodPro = False
                '            BCiePedProCodArt = False
                '            BCiePedCliCodCli = True
                '            BCiePedCliCodArt = False
                'OPCION POR MATERIA PRIMA
                If OptCiePedCli.Item(2).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = False
                            BPedCliCodCli = False
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = True
                
            
                            'OPCION DE BODEGA
                            If BCiePedCliCodCli = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes")
                            'OPCION DE MATERIA PRIMA
                            ElseIf BCiePedCliCodArt = True Then
                                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                            End If
                                    
                                    Set DBGridBusqueda.DataSource = RBusqueda
                                    
                                    If BPedCliCodCli = True Then
                                        DBGridBusqueda.Columns(1).Width = "4000"
                                    ElseIf BPedCliCodArt = True Then
                                        DBGridBusqueda.Columns(1).Width = "4000"
                                    End If
                                    Columnas
                                    FrameBusqueda.Visible = True
                                    TxtBusqueda.SetFocus
                End If
        End If
End Sub

Private Sub TxtCiePedPro_Change()
'BUSCA CODIGO DE MATERIA PRIMA
    If (OptCiePedPro.Item(2).Value = True) Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCiePedPro.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCiePedPro.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblCiePedProDes.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblCiePedProDes.Caption = ""
                End If
    End If
    
    'BUSCA Descripcion
   ' If (OptCiePedPro.Item(1).Value = True) Then
   '         Set RBuscaProveedor = Db.OpenRecordset("Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtCiePedPro.Text & "'")
   '             If RBuscaProveedor.RecordCount > 0 Then
   '                 LblCiePedProDes.Caption = RBuscaProveedor!Descripcion
   '             Else
   '                 LblCiePedProDes.Caption = ""
   '             End If
   '
   ' End If
End Sub

Private Sub TxtCiePedPro_DblClick()
                'OPCION POR Descripcion
                'If OptCiePedPro.Item(1).Value = True Then
                '            BPedProCodPro = False
                '            BPedProCodArt = False
                '            BPedCliCodCli = False
                '            BPedCliCodArt = False
                '            BCiePedProCodPro = True
                '            BCiePedProCodArt = False
                '            BCiePedCliCodCli = False
                '            BCiePedCliCodArt = False
                'OPCION POR MATERIA PRIMA
                Set RBusqueda = New ADODB.Recordset
                If OptCiePedPro.Item(2).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = False
                            BPedCliCodCli = False
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = True
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
            
            
                        'OPCION DE BODEGA
                        If BCiePedProCodPro = True Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores")
                        'OPCION DE MATERIA PRIMA
                        ElseIf BCiePedProCodArt = True Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                        End If
                                
                                Set DBGridBusqueda.DataSource = RBusqueda
                                If BPedProCodPro = True Then
                                    DBGridBusqueda.Columns(1).Width = "4000"
                                ElseIf BPedProCodArt = True Then
                                    DBGridBusqueda.Columns(1).Width = "4000"
                                End If
                                Columnas
                                FrameBusqueda.Visible = True
                                TxtBusqueda.SetFocus
                End If

End Sub

Private Sub TxtCiePedPro_GotFocus()
        TxtCiePedPro.SelStart = 0
        TxtCiePedPro.SelLength = Len(TxtCiePedPro.Text)
End Sub

Private Sub TxtCiePedPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{TAB}"
        End If
        

        If KeyAscii = 43 Then
                'OPCION POR Descripcion
                'If OptCiePedPro.Item(1).Value = True Then
                '            BPedProCodPro = False
                '            BPedProCodArt = False
                '            BPedCliCodCli = False
                '            BPedCliCodArt = False
                '            BCiePedProCodPro = True
                '            BCiePedProCodArt = False
                '            BCiePedCliCodCli = False
                '            BCiePedCliCodArt = False
                'OPCION POR MATERIA PRIMA
                Set RBusqueda = New ADODB.Recordset
                
                If OptCiePedPro.Item(2).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = False
                            BPedCliCodCli = False
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = True
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                
            
                        'OPCION DE BODEGA
                        If BCiePedProCodPro = True Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores")
                        'OPCION DE MATERIA PRIMA
                        ElseIf BCiePedProCodArt = True Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                        End If
                                
                                Set DBGridBusqueda.DataSource = RBusqueda
                                If BPedProCodPro = True Then
                                    DBGridBusqueda.Columns(1).Width = "4000"
                                ElseIf BPedProCodArt = True Then
                                    DBGridBusqueda.Columns(1).Width = "4000"
                                End If
                                Columnas
                                FrameBusqueda.Visible = True
                                TxtBusqueda.SetFocus
                End If
        End If

End Sub

Private Sub TxtPedCli_Change()
    'BUSCA CODIGO DE MATERIA PRIMA
    If OptPedCli.Item(2).Value = True Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtPedCli.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtPedCli.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblPedCliDes.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblPedCliDes.Caption = ""
                End If
    End If
    
    'BUSCA CLIENTE
    If OptPedCli.Item(1).Value = True Then
        Set RBuscaCliente = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where CodigoCliente = '" & TxtPedCli.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where UPPER(CodigoCliente) = '" & UCase(TxtPedCli.Text) & "'")
            End If
                If RBuscaCliente.RecordCount > 0 Then
                    LblPedCliDes.Caption = RBuscaCliente!Descripcion
                Else
                    LblPedCliDes.Caption = ""
                End If
    End If

End Sub

Private Sub TxtPedCli_DblClick()
                'OPCION POR CLIENTE
                If OptPedCli.Item(1).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = False
                            BPedCliCodCli = True
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                'OPCION POR MATERIA PRIMA
                ElseIf OptPedCli.Item(2).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = False
                            BPedCliCodCli = False
                            BPedCliCodArt = True
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                End If
            
                Set RBusqueda = New ADODB.Recordset
                'OPCION DE BODEGA EN CARPETA DE TRASLADOS
                If BPedCliCodCli = True Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes")
                'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
                ElseIf BPedCliCodArt = True Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                End If
                        
                        Set DBGridBusqueda.DataSource = RBusqueda
                        If BPedCliCodCli = True Then
                            DBGridBusqueda.Columns(1).Width = "4000"
                        ElseIf BPedCliCodArt = True Then
                            DBGridBusqueda.Columns(1).Width = "4000"
                        End If
                        Columnas
                        FrameBusqueda.Visible = True
                        TxtBusqueda.SetFocus
End Sub

Private Sub TxtPedCli_GotFocus()
            TxtPedCli.SelStart = 0
            TxtPedCli.SelLength = Len(TxtPedCli.Text)
End Sub

Private Sub TxtPedCli_KeyPress(KeyAscii As Integer)

        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then

                'OPCION POR CLIENTE
                If OptPedCli.Item(1).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = False
                            BPedCliCodCli = True
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                'OPCION POR MATERIA PRIMA
                ElseIf OptPedCli.Item(2).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = False
                            BPedCliCodCli = False
                            BPedCliCodArt = True
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                End If
            
                Set RBusqueda = New ADODB.Recordset
                'OPCION DE BODEGA EN CARPETA DE TRASLADOS
                If BPedCliCodCli = True Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes")
                'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
                ElseIf BPedCliCodArt = True Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                End If
                        
                        Set DBGridBusqueda.DataSource = RBusqueda
                        If BPedCliCodCli = True Then
                            DBGridBusqueda.Columns(1).Width = "4000"
                        ElseIf BPedCliCodArt = True Then
                            DBGridBusqueda.Columns(1).Width = "4000"
                        End If
                        Columnas
                        FrameBusqueda.Visible = True
                        TxtBusqueda.SetFocus
        End If

End Sub

Private Sub TxtPedPro_Change()
    'BUSCA CODIGO DE MATERIA PRIMA
    If OptPedPro.Item(2).Value = True Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtPedPro.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtPedPro.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblPedProDes.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblPedProDes.Caption = ""
                End If
    End If
    
    'BUSCA Descripcion
    If OptPedPro.Item(1).Value = True Then
            Set RBuscaProveedor = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtPedPro.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtPedPro.Text) & "'")
                End If
                If RBuscaProveedor.RecordCount > 0 Then
                    LblPedProDes.Caption = RBuscaProveedor!Descripcion
                Else
                    LblPedProDes.Caption = ""
                End If
          
    End If

End Sub

Private Sub TxtPedPro_DblClick()
                'OPCION POR Descripcion
                If OptPedPro.Item(1).Value = True Then
                            BPedProCodPro = True
                            BPedProCodArt = False
                            BPedCliCodCli = False
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                'OPCION POR MATERIA PRIMA
                ElseIf OptPedPro.Item(2).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = True
                            BPedCliCodCli = False
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                End If
                Set RBusqueda = New ADODB.Recordset
                'OPCION DE BODEGA EN CARPETA DE TRASLADOS
                If BPedProCodPro = True Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores")
                'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
                ElseIf BPedProCodArt = True Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                End If
                        
                        Set DBGridBusqueda.DataSource = RBusqueda
                        If BPedProCodPro = True Then
                            DBGridBusqueda.Columns(1).Width = "4000"
                        ElseIf BPedProCodArt = True Then
                            DBGridBusqueda.Columns(1).Width = "4000"
                        End If
                        Columnas
                        FrameBusqueda.Visible = True
                        TxtBusqueda.SetFocus
End Sub

Private Sub TxtPedPro_GotFocus()
        TxtPedPro.SelStart = 0
        TxtPedPro.SelLength = Len(TxtPedPro.Text)
End Sub

Private Sub TxtPedPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
                'OPCION POR Descripcion
                If OptPedPro.Item(1).Value = True Then
                            BPedProCodPro = True
                            BPedProCodArt = False
                            BPedCliCodCli = False
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                'OPCION POR MATERIA PRIMA
                ElseIf OptPedPro.Item(2).Value = True Then
                            BPedProCodPro = False
                            BPedProCodArt = True
                            BPedCliCodCli = False
                            BPedCliCodArt = False
                            BCiePedProCodPro = False
                            BCiePedProCodArt = False
                            BCiePedCliCodCli = False
                            BCiePedCliCodArt = False
                End If
                
                Set RBusqueda = New ADODB.Recordset
                'OPCION DE BODEGA EN CARPETA DE TRASLADOS
                If BPedProCodPro = True Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion From Proveedores")
                'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
                ElseIf BPedProCodArt = True Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                End If
                        
                        Set DBGridBusqueda.DataSource = RBusqueda
                        If BPedProCodPro = True Then
                            DBGridBusqueda.Columns(1).Width = "4000"
                        ElseIf BPedProCodArt = True Then
                            DBGridBusqueda.Columns(1).Width = "4000"
                        End If
                        Columnas
                        FrameBusqueda.Visible = True
                        TxtBusqueda.SetFocus
        End If

End Sub

Sub Columnas()
        DBGridBusqueda.Columns(1).Width = "4000"
End Sub


Public Sub CierrePedidosProveedor()
            VDia = Day(DTPCiePedProFecIni.Value)
            VMes = Month(DTPCiePedProFecIni.Value)
            VAño = Year(DTPCiePedProFecIni.Value)
            VDia2 = Day(DTPCiePedProFecFin.Value)
            VMes2 = Month(DTPCiePedProFecFin.Value)
            VAño2 = Year(DTPCiePedProFecFin.Value)
            
                    'FECHAS
                       
                    If OptCiePedPro.Item(0).Value = True Then
                         GTituloReporte = "Fechas De Despacho De " & Format(DTPCiePedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPCiePedProFecFin.Value, "dd/mm/yyyy")
                         GCriteriaReporte = "{EncabezadoCierrePedidosProve.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    'PEDIDO
                    ElseIf OptCiePedPro.Item(1).Value = True Then
                         GTituloReporte = "Cierres Del Pedido " & TxtCiePedPro.Text & " " & LblCiePedProDes.Caption
                         GCriteriaReporte = "{DetalleCierrePedidosProve.Pedido} Like '" & TxtCiePedPro.Text & "*'"
                    'FECHAS Y CODIGO
                    ElseIf OptCiePedPro.Item(2).Value = True Then
                         GTituloReporte = "Fechas De Despacho De " & Format(DTPCiePedProFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPCiePedProFecFin.Value, "dd/mm/yyyy") & " Del Codigo " & TxtCiePedPro.Text & " " & LblCiePedProDes.Caption
                         GCriteriaReporte = "{EncabezadoCierrePedidosProve.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleCierrePedidosProve.Codigo} Like '" & TxtCiePedPro.Text & "*'"
                    'NUMERO DOCUMENTO
                    ElseIf OptCiePedPro.Item(3).Value = True Then
                         GTituloReporte = "Cierres Del Documento " & TxtCiePedPro.Text & " " & LblCiePedProDes.Caption
                         GCriteriaReporte = "{EncabezadoCierrePedidosProve.NumeroDocumento} Like '" & TxtCiePedPro.Text & "*'"
                    End If
                
                'ELIGE REPORTE DE ACUERDO A LA OPCION
                If OptCiePedProRes.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "CierrePedidosProveedoresResumen.rpt"
                        Else
                            GNombreReporte = "CierrePedidosProveedoresResumenO.rpt"
                        End If
                ElseIf OptCiePedProDet = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "CierrePedidosProveedoresResumen2.rpt"
                        Else
                            GNombreReporte = "CierrePedidosProveedoresResumen2O.rpt"
                        End If
                End If
    
End Sub

Public Sub CierrePedidosCliente()
            VDia = Day(DTPCiePedCliFecIni.Value)
            VMes = Month(DTPCiePedCliFecIni.Value)
            VAño = Year(DTPCiePedCliFecIni.Value)
            VDia2 = Day(DTPCiePedCliFecFin.Value)
            VMes2 = Month(DTPCiePedCliFecFin.Value)
            VAño2 = Year(DTPCiePedCliFecFin.Value)
            
                    'FECHAS
                    If OptCiePedCli.Item(0).Value = True Then
                         GTituloReporte = "Fechas De Despacho De " & Format(DTPCiePedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPCiePedCliFecFin.Value, "dd/mm/yyyy")
                         GCriteriaReporte = "{EncabezadoCierrePedidosCliente.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    'PEDIDO
                    ElseIf OptCiePedCli.Item(1).Value = True Then
                         GTituloReporte = "Cierres Del Pedido " & TxtCiePedCli.Text & " " & LblCiePedCliDes.Caption
                         GCriteriaReporte = "{DetalleCierrePedidosCliente.Pedido} Like '" & TxtCiePedCli.Text & "*'"
                    'FECHAS Y CODIGO
                    ElseIf OptCiePedCli.Item(2).Value = True Then
                         GTituloReporte = "Fechas De Despacho De " & Format(DTPCiePedCliFecIni.Value, "dd/mm/yyyy") & " A " & Format(DTPCiePedCliFecFin.Value, "dd/mm/yyyy") & " Del Codigo " & TxtCiePedCli.Text & " " & LblCiePedCliDes.Caption
                         GCriteriaReporte = "{EncabezadoCierrePedidosCliente.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {DetalleCierrePedidosCliente.Codigo} Like '" & TxtCiePedCli.Text & "*'"
                    'NUMERO PEDIDO
                    ElseIf OptCiePedCli.Item(3).Value = True Then
                         GTituloReporte = "Cierres Del Documento " & TxtCiePedCli.Text & " " & LblCiePedCliDes.Caption
                         GCriteriaReporte = "{EncabezadoCierrePedidosCliente.NumeroDocumento} Like '" & TxtCiePedCli.Text & "*'"
                    End If
                
                'ELIGE REPORTE DE ACUERDO A LA OPCION
                If OptCiePedCliRes.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "CierrePedidosClientesResumen.rpt"
                    Else
                        GNombreReporte = "CierrePedidosClientesResumenO.rpt"
                    End If
                ElseIf OptCiePedCliDet = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "CierrePedidosClientesResumen2.rpt"
                    Else
                        GNombreReporte = "CierrePedidosClientesResumen2O.rpt"
                    End If
                End If
    
End Sub
