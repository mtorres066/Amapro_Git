VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InventarioSalidas 
   BackColor       =   &H00FF8080&
   Caption         =   "Salidas De Inventario"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "InventarioSalidas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
      Caption         =   "Busqueda De Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6735
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   11880
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
         Left            =   11040
         Picture         =   "InventarioSalidas.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1440
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   4935
      End
   End
   Begin TabDlg.SSTab TabDespachos 
      Height          =   8055
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14208
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   882
      BackColor       =   16744576
      TabCaption(0)   =   "Encabezado"
      TabPicture(0)   =   "InventarioSalidas.frx":293C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameEncabezado"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "InventarioSalidas.frx":2D8E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameDetalle"
      Tab(1).Control(1)=   "TxtCueTar"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "DBGridDetalle"
      Tab(1).ControlCount=   3
      Begin MSDataGridLib.DataGrid DBGridDetalle 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   92
         Top             =   2400
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
            DataField       =   "FechaProduccion"
            Caption         =   "FechaProduccion"
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
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "FichaTecnica"
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
         BeginProperty Column04 
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
         BeginProperty Column05 
            DataField       =   "Batch"
            Caption         =   "Batch"
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
            DataField       =   "Calidad"
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
         BeginProperty Column07 
            DataField       =   "Bodega"
            Caption         =   "Bodega"
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
            DataField       =   "Cantidad"
            Caption         =   "Cantidad"
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column01 
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1725.165
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   720
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtCueTar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
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
         Left            =   -64200
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   7440
         Width           =   855
      End
      Begin VB.Frame FrameDetalle 
         Caption         =   "Detalle Salidas"
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
         Height          =   7335
         Left            =   -74880
         TabIndex        =   56
         Top             =   600
         Width           =   11685
         Begin VB.Frame FrameDetalleCompras 
            Enabled         =   0   'False
            Height          =   1455
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   11415
            Begin VB.TextBox TxtCal 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2760
               TabIndex        =   64
               Top             =   1080
               Width           =   1275
            End
            Begin VB.TextBox TxtBarra 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   35
               TabIndex        =   58
               Top             =   360
               Width           =   2535
            End
            Begin VB.TextBox TxtDesPro 
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
               Height          =   285
               Left            =   3360
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   74
               TabStop         =   0   'False
               Top             =   720
               Width           =   4815
            End
            Begin VB.TextBox TxtCodPro 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   1440
               MaxLength       =   15
               TabIndex        =   61
               Top             =   720
               Width           =   1815
            End
            Begin VB.TextBox TxtDocDet 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   7800
               TabIndex        =   73
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.TextBox TxtBod 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               MaxLength       =   3
               TabIndex        =   66
               Top             =   1080
               Width           =   495
            End
            Begin VB.TextBox TxtTar 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   8880
               TabIndex        =   62
               Top             =   720
               Width           =   495
            End
            Begin VB.TextBox TxtLin 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   4080
               MaxLength       =   2
               TabIndex        =   60
               Top             =   360
               Width           =   495
            End
            Begin VB.TextBox TxtBat 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   4800
               TabIndex        =   65
               Top             =   1080
               Width           =   1155
            End
            Begin MSMask.MaskEdBox MskFecPro 
               Height          =   285
               Left            =   2760
               TabIndex        =   59
               Top             =   360
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   8438015
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TxtCanPro 
               Height          =   285
               Left            =   10320
               TabIndex        =   63
               Top             =   720
               Width           =   975
               _ExtentX        =   1720
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,###,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Codigo De Barra"
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
               TabIndex        =   86
               Top             =   120
               Width           =   1575
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
               Index           =   0
               Left            =   120
               TabIndex        =   85
               Top             =   720
               Width           =   1230
            End
            Begin VB.Label Label2 
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
               Left            =   4680
               TabIndex        =   84
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label3 
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
               Index           =   0
               Left            =   9480
               TabIndex        =   83
               Top             =   720
               Width           =   855
            End
            Begin VB.Label Label6 
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
               Left            =   8280
               TabIndex        =   82
               Top             =   720
               Width           =   585
            End
            Begin VB.Label Label6 
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
               Index           =   8
               Left            =   2760
               TabIndex        =   81
               Top             =   120
               Width           =   540
            End
            Begin VB.Label Label6 
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
               Index           =   9
               Left            =   4080
               TabIndex        =   80
               Top             =   120
               Width           =   480
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Batch"
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
               Index           =   10
               Left            =   4080
               TabIndex        =   79
               Top             =   1080
               Width           =   510
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Bodega"
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
               Left            =   6000
               TabIndex        =   78
               Top             =   1080
               Width           =   660
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Calidad"
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
               Index           =   21
               Left            =   2040
               TabIndex        =   77
               Top             =   1080
               Width           =   645
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
               Left            =   4680
               TabIndex        =   76
               Top             =   360
               Width           =   6615
            End
            Begin VB.Label LblBod 
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
               Left            =   7320
               TabIndex        =   75
               Top             =   1080
               Width           =   3975
            End
         End
         Begin VB.CommandButton CmdAgregar2 
            Caption         =   "A&gregar"
            Height          =   480
            Left            =   120
            Picture         =   "InventarioSalidas.frx":30A8
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   6720
            Visible         =   0   'False
            Width           =   1600
         End
         Begin VB.CommandButton CmdGrabar2 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   480
            Left            =   3480
            Picture         =   "InventarioSalidas.frx":35DA
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   6720
            Visible         =   0   'False
            Width           =   1600
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
            Height          =   480
            Left            =   8520
            Picture         =   "InventarioSalidas.frx":3B0C
            TabIndex        =   72
            Top             =   6720
            Visible         =   0   'False
            Width           =   1600
         End
         Begin VB.CommandButton CmdCancelar2 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   480
            Left            =   5160
            Picture         =   "InventarioSalidas.frx":403E
            Style           =   1  'Graphical
            TabIndex        =   70
            Top             =   6720
            Visible         =   0   'False
            Width           =   1600
         End
         Begin VB.CommandButton CmdBorrar2 
            Caption         =   "B&orrar"
            Height          =   480
            Left            =   6840
            Picture         =   "InventarioSalidas.frx":4570
            Style           =   1  'Graphical
            TabIndex        =   71
            Top             =   6720
            Visible         =   0   'False
            Width           =   1600
         End
         Begin VB.CommandButton CmdEditar2 
            Caption         =   "Editar"
            Height          =   480
            Left            =   1800
            Picture         =   "InventarioSalidas.frx":4AA2
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   6720
            Visible         =   0   'False
            Width           =   1600
         End
      End
      Begin VB.Frame FrameEncabezado 
         Caption         =   "Encabezado De Salidas"
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
         Height          =   6855
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   11655
         Begin VB.CommandButton CmdImprimirCedulas 
            Caption         =   "&Imprimir Cedulas"
            Height          =   720
            Left            =   8640
            Picture         =   "InventarioSalidas.frx":4FD4
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   6000
            Width           =   1000
         End
         Begin VB.CommandButton CmdBotones2 
            BackColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   1
            Left            =   120
            MouseIcon       =   "InventarioSalidas.frx":5506
            Picture         =   "InventarioSalidas.frx":5948
            Style           =   1  'Graphical
            TabIndex        =   91
            ToolTipText     =   "Primer Registro"
            Top             =   6120
            Width           =   375
         End
         Begin VB.CommandButton CmdBotones2 
            BackColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   2
            Left            =   480
            MouseIcon       =   "InventarioSalidas.frx":5E7A
            Picture         =   "InventarioSalidas.frx":62BC
            Style           =   1  'Graphical
            TabIndex        =   90
            ToolTipText     =   "Registro Anterior"
            Top             =   6120
            Width           =   375
         End
         Begin VB.CommandButton CmdBotones2 
            BackColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   3
            Left            =   10800
            MouseIcon       =   "InventarioSalidas.frx":67EE
            Picture         =   "InventarioSalidas.frx":6C30
            Style           =   1  'Graphical
            TabIndex        =   89
            ToolTipText     =   "Siguiente Registro"
            Top             =   6120
            Width           =   375
         End
         Begin VB.CommandButton CmdBotones2 
            BackColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   4
            Left            =   11160
            MouseIcon       =   "InventarioSalidas.frx":7162
            Picture         =   "InventarioSalidas.frx":75A4
            Style           =   1  'Graphical
            TabIndex        =   88
            ToolTipText     =   "Ultimo Registro"
            Top             =   6120
            Width           =   375
         End
         Begin VB.Frame FrameCompras 
            Enabled         =   0   'False
            Height          =   5655
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Width           =   11415
            Begin VB.CheckBox ChkBusTar 
               Caption         =   "Buscar Tarimas"
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
               Height          =   195
               Left            =   1560
               TabIndex        =   5
               Top             =   2400
               Width           =   1935
            End
            Begin VB.TextBox TxtDocIng 
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
               ForeColor       =   &H00FF0000&
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   34
               TabStop         =   0   'False
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtLib 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   33
               TabStop         =   0   'False
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtReq 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   32
               TabStop         =   0   'False
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TxtBatch 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
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
               Left            =   1560
               TabIndex        =   6
               Top             =   2760
               Width           =   1215
            End
            Begin VB.TextBox TxtEstado 
               Appearance      =   0  'Flat
               BackColor       =   &H00FF8080&
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   31
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   6
               Left            =   1560
               MaxLength       =   50
               TabIndex        =   13
               Top             =   5280
               Width           =   6855
            End
            Begin VB.TextBox TxtCli 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   3
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox TxtTra 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   4
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox TxtNumDoc 
               Appearance      =   0  'Flat
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
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   1
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   8
               Top             =   3480
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   9
               Top             =   3840
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   3
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   10
               Top             =   4200
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   4
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   11
               Top             =   4560
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   5
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   12
               Top             =   4920
               Width           =   1215
            End
            Begin VB.TextBox TxtLinea 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
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
               Left            =   1560
               MaxLength       =   2
               TabIndex        =   7
               Top             =   3120
               Width           =   1215
            End
            Begin VB.TextBox TxtTipDoc 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   2
               Top             =   1320
               Width           =   1215
            End
            Begin MSMask.MaskEdBox MskFec 
               Height          =   285
               Left            =   1560
               TabIndex        =   0
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
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
               Index           =   0
               Left            =   120
               TabIndex        =   55
               Top             =   600
               Width           =   540
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Transaccion"
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
               Left            =   120
               TabIndex        =   54
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label6 
               Caption         =   "Cliente"
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
               Top             =   1680
               Width           =   975
            End
            Begin VB.Label Label6 
               Caption         =   "No. Batch"
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
               Index           =   13
               Left            =   120
               TabIndex        =   52
               Top             =   2760
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Clasificacion"
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
               Index           =   12
               Left            =   8520
               TabIndex        =   51
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Requerido"
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
               Left            =   8760
               TabIndex        =   50
               Top             =   600
               Width           =   885
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Liberado"
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
               Left            =   8880
               TabIndex        =   49
               Top             =   960
               Width           =   750
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
               Index           =   4
               Left            =   120
               TabIndex        =   48
               Top             =   5280
               Width           =   1275
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
               Left            =   2880
               TabIndex        =   47
               Top             =   1680
               Width           =   5535
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Transportista"
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
               Left            =   120
               TabIndex        =   46
               Top             =   2040
               Width           =   1125
            End
            Begin VB.Label LblTra 
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
               TabIndex        =   45
               Top             =   2040
               Width           =   5535
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
               Index           =   11
               Left            =   120
               TabIndex        =   44
               Top             =   1320
               Width           =   1410
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "# Documento"
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
               Height          =   195
               Index           =   14
               Left            =   120
               TabIndex        =   43
               Top             =   960
               Width           =   1155
            End
            Begin VB.Label Label6 
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
               Index           =   15
               Left            =   120
               TabIndex        =   42
               Top             =   3120
               Width           =   480
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Cargado Por"
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
               Index           =   16
               Left            =   120
               TabIndex        =   41
               Top             =   3480
               Width           =   1065
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Entregado Por"
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
               Index           =   17
               Left            =   120
               TabIndex        =   40
               Top             =   3840
               Width           =   1230
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Conductor"
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
               Index           =   18
               Left            =   120
               TabIndex        =   39
               Top             =   4200
               Width           =   885
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Placas Camion"
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
               Index           =   19
               Left            =   120
               TabIndex        =   38
               Top             =   4560
               Width           =   1260
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Placas Furgon"
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
               Index           =   20
               Left            =   120
               TabIndex        =   37
               Top             =   4920
               Width           =   1230
            End
            Begin VB.Label LblDocumento 
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
               TabIndex        =   36
               Top             =   1320
               Width           =   5535
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
               TabIndex        =   35
               Top             =   3120
               Width           =   5535
            End
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar"
            Height          =   720
            Left            =   960
            Picture         =   "InventarioSalidas.frx":7AD6
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   6000
            Width           =   1125
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "&Grabar"
            Enabled         =   0   'False
            Height          =   720
            Left            =   3240
            Picture         =   "InventarioSalidas.frx":7E53
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   6000
            Width           =   1000
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   720
            Left            =   4320
            Picture         =   "InventarioSalidas.frx":83AF
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   6000
            Width           =   1000
         End
         Begin VB.CommandButton CmdBorrar 
            Caption         =   "&Borrar"
            Height          =   720
            Left            =   5400
            Picture         =   "InventarioSalidas.frx":88E6
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   6000
            Width           =   1000
         End
         Begin VB.CommandButton CmdSalida 
            Appearance      =   0  'Flat
            Height          =   720
            Left            =   9720
            Picture         =   "InventarioSalidas.frx":8EAE
            Style           =   1  'Graphical
            TabIndex        =   21
            ToolTipText     =   "Salida"
            Top             =   6000
            Width           =   1000
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "B&uscar Transaccion"
            Height          =   720
            Left            =   6480
            Picture         =   "InventarioSalidas.frx":93C9
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   6000
            Width           =   1000
         End
         Begin VB.CommandButton CmdEditar 
            Caption         =   "&Editar"
            Height          =   720
            Left            =   2160
            Picture         =   "InventarioSalidas.frx":9851
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   6000
            Width           =   1000
         End
         Begin VB.CommandButton CmdImprimir 
            Caption         =   "&Imprimir"
            Height          =   720
            Left            =   7560
            Picture         =   "InventarioSalidas.frx":9C28
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   6000
            Width           =   1000
         End
      End
   End
End
Attribute VB_Name = "InventarioSalidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As Long
Dim VCantidad As Long
Dim VCodigoProducto As String
Dim VCantidadProducto As Long
Dim VBodega As String
Dim VBatch As Double
Dim VClasificacion As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean
Dim BCliente As Boolean
Dim BProducto As Boolean
Dim BDocumento As Boolean
Dim BEditar As Boolean
Dim BBodegaDetalle As Boolean
Dim BTransportistas As Boolean
Dim BEditarEncabezado As Boolean
Dim BEditarDetalle As Boolean
Dim BTarima As Boolean
Dim VLinea As String

Dim RBuscaProducto As New ADODB.Recordset
Dim RMaximo As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset
Dim RBuscaDetalle As New ADODB.Recordset
Dim RBuscaEncabezado As New ADODB.Recordset
Dim RBuscaEntradasInventario As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RCuentaTarimas As New ADODB.Recordset
Dim RBuscaCliente As New ADODB.Recordset
Dim RBuscaTransportista As New ADODB.Recordset
Dim RBuscaSalidasMP As New ADODB.Recordset
Dim RReporteSalidasMP As New ADODB.Recordset
Dim RBuscaBarra As New ADODB.Recordset
Dim RBuscaTarimas As New ADODB.Recordset

Dim RBusqueda As New ADODB.Recordset
Dim REncabezado As New ADODB.Recordset
Dim RDetalle As New ADODB.Recordset
Dim VTexto As String
Dim VTexto2 As String

Sub Botones1()
    If Bandera = True Then
         FrameCompras.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdSalida.Enabled = False
         CmdImprimir.Enabled = False
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         
    Else
         FrameCompras.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdSalida.Enabled = True
         CmdImprimir.Enabled = True
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         
         
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleCompras.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
         CmdEditar2.Enabled = False
    Else
         FrameDetalleCompras.Enabled = False
         CmdAgregar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
         CmdImprimir.Enabled = True
         CmdBuscar.Enabled = True
         CmdEditar2.Enabled = True
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


Private Sub Check1_Click()

End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub CmdAgregar2_Click()
On Error Resume Next
    
    Bandera2 = True
    Botones2
    Limpia_CamposDetalle
    DbGridDetalle.Enabled = False
    TxtDocDet.Text = VDocumento
    
    'ASIGNA LA BODEGA INGRESADA EN EL ENCABEZADO
    TxtBod.Text = VBodega
    TxtBat.Text = VBatch
    TxtBarra.SetFocus
    TxtDesPro.Text = ""
    TxtCal.Text = "A"
    MskFecPro.Enabled = True
    TxtLin.Enabled = True
    TxtCodPro.Enabled = True
    TxtTar.Enabled = True
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            If GBorrar = True Then
                    'NO HACE NADA PORQUE SI TIENE ACCESO A BORRAR
            ElseIf TxtEstado.Text = "LIBERADO" Then
                    'VERIFICA SI YA FUE LIBERADA LA ENTRADA
                    MsgBox "Este Documento No Se Puede BORRAR Porque Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
                    Exit Sub
            End If
            
            VDocumento = TxtDocIng.Text
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                                REncabezado.Delete
                            
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    If Err <> 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                        Err.Clear
                                    End If
                                Else 'ORACLE
                                    'SI HAY ERRORES
                                    If Err = -2147217873 Then
                                        MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                        Err.Clear
                                    ElseIf Err <> -2147217873 And Err <> 0 Then
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
                                            Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & TxtDocIng.Text)
                                            Else 'ORACLE
                                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & TxtDocIng.Text)
                                            End If
                                                Llena_CamposDetalle
                                                Set DbGridDetalle.DataSource = RDetalle
                
            End If
            
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
                        
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                            MousePointer = 11
                        
                        'BORRA EL REGISTRO
                        RDetalle.Delete
                    
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147217873 Then
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                                Exit Sub
                            End If
                        End If
                        
                    
                    'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RDetalle.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RDetalle.MoveLast
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
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & TxtDocIng.Text)
                        Else 'ORACLE
                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & TxtDocIng.Text)
                        End If
                         Llena_CamposDetalle
                         Set DbGridDetalle.DataSource = RDetalle
    
MousePointer = 0


End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
    mensaje = InputBox("Transaccion a Buscar")
    If IsNumeric(mensaje) Then
        
                REncabezado.MoveFirst
                REncabezado.Find " Documento = " & mensaje
                
                                                
                Llena_CamposEncabezado
                
                        Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & TxtDocIng.Text)
                                Else 'ORACLE
                                        Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & TxtDocIng.Text)
                                End If
                                        Llena_CamposDetalle
                                        Set DbGridDetalle.DataSource = RDetalle
    Else
            MsgBox "La Transaccion Debe Ser Numerica", vbOKOnly + vbInformation, "Informacion"
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
    MskFecPro.Enabled = True
    TxtLin.Enabled = True
    TxtCodPro.Enabled = True
    TxtTar.Enabled = True
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
    'VERIFICA SI YA FUE LIBERADA LA ENTRADA
    If GEditar = True Then
        'NO HACE NADA PORQUE SI TIENE ACCESO A EDITAR
    ElseIf TxtEstado.Text = "LIBERADO" Then
        MsgBox "Esta Documento No Se Puede EDITAR Porque Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
        Exit Sub
    End If
        
    BEditarEncabezado = True
    BEditar = True
    Bandera = True
    Botones1
    MskFec.SetFocus
    'GRABA EL USUARIO QUE ESTA EDITANDO
    TxtReq.Text = GUsuario
    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False
End Sub


Private Sub CmdEditar2_Click()
    On Error Resume Next
    DbGridDetalle.Enabled = False
    BEditarDetalle = True
    Bandera2 = True
    Botones2
    VDocumento = TxtDocDet.Text
    TxtCodPro.SetFocus
    
    MskFecPro.Enabled = False
    TxtLin.Enabled = False
    TxtCodPro.Enabled = False
    TxtTar.Enabled = False

End Sub

Private Sub CmdGrabar2_Click()
On Error Resume Next
    
    TxtCodPro.Text = UCase(TxtCodPro.Text)
    
    'GUARDA VARIABLES
    VCantidad = TxtCanPro.Text
    VCodigoProducto = TxtCodPro.Text
        
    'REVISAMOS DATOS
    If Not IsNumeric(TxtCanPro.Text) Then
       MsgBox "Cantidad De Producto Incorrecta", vbOKOnly + vbCritical, "Error"
       TxtCanPro.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS EL BATCH DE DETALLE
    If Not IsNumeric(TxtBat.Text) Then
       MsgBox "Numero De Bath Incorrecto", vbOKOnly + vbCritical, "Error"
       TxtBat.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS LA TARIMA
    If Not IsNumeric(TxtTar.Text) Then
       MsgBox "Numero De Tarima Incorrecto", vbOKOnly + vbCritical, "Error"
       TxtBat.SetFocus
       Exit Sub
    End If
    
    
    'REVISAMOS DATOS
    If Not IsDate(MskFecPro.Text) Then
       MsgBox "Fecha De Produccion Incorrecta", vbOKOnly + vbCritical, "Error"
       MskFecPro.SetFocus
       Exit Sub
    End If
    
        
        
        'REVISA SI LA TARIMA EXISTE EN LA ENTRADAS DE PRODUCTO TERMINADO
        Set RBuscaEntradasInventario = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaEntradasInventario, "Select * From DetalleEntradasInventario Where FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTar.Text & " And Linea = '" & TxtLin.Text & "' And FichaTecnica = '" & TxtCodPro.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaEntradasInventario, "Select * From DetalleEntradasInventario Where FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTar.Text & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtCodPro.Text) & "'")
                End If
    
                If RBuscaEntradasInventario.RecordCount > 0 Then
                    'SI LA ENCUENTRA NO HACE NADA
                Else
                        MsgBox "Bulto/Tarima No Existe En Inventario", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                End If
                    
                            If BEditarDetalle = False Then
                                    VTexto = TxtDocDet.Text & ", " ' DOCUMENTO
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                            VTexto = VTexto & "#" & Format(MskFecPro.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                                    Else 'ORACLE
                                            VTexto = VTexto & "To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & ", '"  'FECHA
                                    End If
                                    VTexto = VTexto & TxtLin.Text & "', '" 'LINEA
                                    VTexto = VTexto & TxtCodPro.Text & "', " 'FICHA TECNICA
                                    VTexto = VTexto & TxtTar.Text & ", " 'TARIMA
                                    VTexto = VTexto & TxtBat.Text & ", '" 'BATCH
                                    VTexto = VTexto & TxtCal.Text & "', '" 'CALIDAD
                                    VTexto = VTexto & TxtBod.Text & "', " 'BODEGA
                                    VTexto = VTexto & TxtCanPro.Text 'CANTIDAD
                                    
                                    Conexion.Execute "Insert Into DetalleSalidasInventario Values(" & VTexto & ")"
                            Else 'SI ESTA EDITANDO
                                    
                                    'If GOrigenDeDatos = "AmaproAccess" Then
                                    '    VTexto = "Fecha = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "#, " 'FECHA
                                    'Else 'ORACLE
                                    '    VTexto = "Fecha = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                                    'End If
                                    'VTexto = VTexto & "Linea = '" & TxtLin.Text & "', "  'LINEA
                                    'VTexto = VTexto & "FichaTecnica = '" & TxtCodPro.Text & "', " 'FICHA
                                    'VTexto = VTexto & "Tarima = " & TxtTar.Text & ", " 'TARIMA
                                    VTexto = "Batch = " & TxtBat.Text & ", " 'BATCH
                                    VTexto = VTexto & "Calidad = '" & TxtCal.Text & "', " 'CALIDAD
                                    VTexto = VTexto & "Bodega = '" & TxtBod.Text & "', " 'BODEGA
                                    VTexto = VTexto & "Cantidad = " & TxtCanPro.Text 'CANTIDAD
                                    
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        VTexto = VTexto & " Where Documento = " & VDocumento & " And FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Linea = '" & TxtLin.Text & "' And FichaTecnica = '" & TxtCodPro.Text & "' And Tarima = " & TxtTar.Text
                                    Else
                                        VTexto = VTexto & " Where Documento = " & VDocumento & " And FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtCodPro.Text) & "' And Tarima = " & TxtTar.Text
                                    End If
                                    
                                    Conexion.Execute "Update DetalleSalidasInventario Set " & VTexto
                            End If
                                        
                                    'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                        'iI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                Err.Clear
                                                Exit Sub
                                        End If
                                    Else 'ORACLE
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err = -2147217873 Then
                                                MsgBox "Tarima/Bulto Ya Existe En Este Documento De Salida", vbOKOnly + vbInformation, "Informacion"
                                                Err.Clear
                                                Exit Sub
                                        ElseIf Err <> -2147217873 And Err <> 0 Then
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                Err.Clear
                                                Exit Sub
                                        End If
                                    End If
                                    
                                    'CUENTA LAS TARIMAS QUE HAN SALIDO
                                    Set RCuentaTarimas = New ADODB.Recordset
                                        Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From DetalleSalidasInventario Where Documento = " & TxtDocIng.Text)
                                        If RCuentaTarimas.RecordCount > 0 Then
                                            TxtCueTar.Text = RCuentaTarimas(0)
                                        Else
                                            TxtCueTar.Text = 0
                                        End If
                
            Bandera2 = False
            Botones2
            RDetalle.Requery
            RDetalle.MoveLast
            Llena_CamposDetalle
            DbGridDetalle.Enabled = True
            TxtDesPro.Text = ""
            CmdAgregar2.SetFocus
            
            MskFecPro.Enabled = True
            TxtLin.Enabled = True
            TxtCodPro.Enabled = True
            TxtTar.Enabled = True
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
    
    BEditar = False
    Bandera = True
    Botones1
    BEditarEncabezado = False
    FrameDetalle.Visible = False
    DbGridDetalle.Visible = False
    Limpia_CamposEncabezado
    TxtReq.Text = GUsuario
    MskFec.Text = Format(Date, "dd/mm/yyyy")
    MskFec.SetFocus
    TxtEstado.Text = "NO LIBERADA"
    TxtCueTar.Text = ""
    
    '
        
    
End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next

MousePointer = 11
    'SI ESTA AGREGANDO UN REGISTRO
    If BEditar = False Then
            'BUSCA SI YA EXISTE EL NUMERO DE DOCUMENTO PARA ESTE TIPO DE DOCUMENTO
            Set RBuscaDocumento = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaDocumento, "Select * From EncabezadoSalidasInventario Where TipoDeDocumento = '" & TxtTipDoc.Text & "' And NumeroDocumento = '" & TxtNumDoc.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaDocumento, "Select * From EncabezadoSalidasInventario Where UPPER(TipoDeDocumento) = '" & UCase(TxtTipDoc.Text) & "' And UPPER(NumeroDocumento) = '" & UCase(TxtNumDoc.Text) & "'")
                End If
                    If RBuscaDocumento.RecordCount > 0 Then
                        MsgBox "Numero Documento Para Este Tipo De Documento Ya Existe", vbOKOnly + vbInformation, "Informacion"
                        TxtTipDoc.SetFocus
                        Exit Sub
                    End If
    End If


    
    'REVISA LA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
        MskFec.SetFocus
        Exit Sub
    End If
        
    
    'REVISA LINEA
    If TxtLinea.Text = "" Then
            MsgBox "Linea No Puede Estar Vacia", vbOKOnly + vbInformation, "Informacion"
            TxtLinea.SetFocus
            Exit Sub
    End If
    
    'REVISA EL BATCH
    If Not IsNumeric(TxtBatch.Text) Then
            MsgBox "El Batch Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
            TxtBatch.SetFocus
            Exit Sub
    End If
    
    'REVISA TRANSPORTISTA
    If TxtTra.Text = "" Then
            MsgBox "Codigo Transportista No Puede Estar Vacia", vbOKOnly + vbInformation, "Informacion"
            TxtTra.SetFocus
            Exit Sub
    End If
    
    


    VDocumento = TxtDocIng.Text
    VBatch = TxtBatch.Text
    VLinea = TxtLinea.Text
              
    MskFec.Text = Format(MskFec.Text, "dd/mm/yyyy")

    
               
                If BEditarEncabezado = False Then
                'BUSCA EL DOCUMENTO MAXIMO Y LE AGREGA UNO MAS
                            Set RMaximo = New ADODB.Recordset
                                Call Abrir_Recordset(RMaximo, "Select max(Documento) from EncabezadoSalidasInventario")
                                If RMaximo.RecordCount > 0 Then
                                    If IsNull(RMaximo(0)) Then
                                        VDocumento = "1"
                                    Else
                                        VDocumento = Val(RMaximo(0)) + 1
                                    End If
                                End If
                                
                            VTexto = VDocumento & ", " 'DOCUMENTO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = VTexto & "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                            Else 'ORACLE
                                VTexto = VTexto & "To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", '"  'FECHA
                            End If
                            VTexto = VTexto & TxtCli.Text & "', " 'CLIENTE
                            VTexto = VTexto & TxtBatch & ", '" 'BATCH
                            VTexto = VTexto & TxtLinea.Text & "', '" 'LINEA
                            VTexto = VTexto & TxtTra.Text & "', '" 'TRANSPORTISTA
                            VTexto = VTexto & TxtTipDoc.Text & "', '" 'TIPO DE DOCUMENTO
                            VTexto = VTexto & TxtNumDoc.Text & "', '" 'NUMERO DE DOCUMENTO
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', '" 'CARGADO POR
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', '" 'ENTREGADO POR
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', '" 'CONDUCTOR
                            VTexto = VTexto & TxtTexto.Item(4).Text & "', '" 'PLACAS CAMION
                            VTexto = VTexto & TxtTexto.Item(5).Text & "', '" 'PLACAS FURGON
                            VTexto = VTexto & TxtTexto.Item(6).Text & "', '" 'OBSERVACIONES
                            VTexto = VTexto & TxtReq.Text & "', '" 'REQUERIDO
                            VTexto = VTexto & TxtLib.Text & "', '" 'LIBERADO
                            VTexto = VTexto & TxtEstado.Text & "'" 'ESTADO
                            
                            Conexion.Execute "Insert Into EncabezadoSalidasInventario Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = "Fecha = #" & Format(MskFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = "Fecha = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "Cliente = '" & UCase(TxtCli.Text) & "', " 'CLIENTE
                            VTexto = VTexto & "Batch = " & UCase(TxtBatch.Text) & ", " 'BATCH
                            VTexto = VTexto & "Linea = '" & TxtLinea.Text & "', " 'LINEA
                            VTexto = VTexto & "CodigoTransportista = '" & TxtTra.Text & "', " 'TRANSPORTISTA
                            VTexto = VTexto & "TipoDeDocumento = '" & TxtTipDoc.Text & "', " 'TIPO DE DOCUMENTO
                            VTexto = VTexto & "NumeroDocumento = '" & TxtNumDoc.Text & "', " 'Numero Documento
                            VTexto = VTexto & "CargadoPor = '" & TxtTexto.Item(1).Text & "', " 'Cargado Por
                            VTexto = VTexto & "EntregadoPor = '" & TxtTexto.Item(2).Text & "', " 'Entregado Por
                            VTexto = VTexto & "Conductor = '" & TxtTexto.Item(3).Text & "', " 'Conductor Por
                            VTexto = VTexto & "PlacasCamion = '" & TxtTexto.Item(4).Text & "', " 'Placas Camion Por
                            VTexto = VTexto & "PlacasFurgon = '" & TxtTexto.Item(5).Text & "', " 'Placas Furgon
                            VTexto = VTexto & "Observaciones = '" & TxtTexto.Item(6).Text & "', " 'Observaciones
                            VTexto = VTexto & "Requerido = '" & TxtReq.Text & "', " 'Requerido
                            VTexto = VTexto & "Liberado = '" & TxtLib.Text & "', " 'Liberado
                            VTexto = VTexto & "Estado = '" & TxtEstado.Text & "'" 'Estado
                            
                            VTexto = VTexto & " Where Documento = " & VDocumento 'DOCUMENTO
                            
                            Conexion.Execute "UPDATE EncabezadoSalidasInventario SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Transaccion Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            MousePointer = 0
                            MskFec.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            MousePointer = 0
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Transaccion Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            MousePointer = 0
                            MskFec.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            MousePointer = 0
                            Exit Sub
                        End If
                    End If
                           
    
                    'CAMBIA BOTONES
                    Bandera = False
                    Botones1
                    
    
                    If VBatch = 0 Then
                        'PARA QUE NO JALEN EL BATCH 0, QUE ES EL DE TODAS LAS MATERIAS PRIMAS
                    Else
                    
                        'SI QUIERE BUSCAR TARIMAS
                        If ChkBusTar.Value = 1 Then
                            
                            'BUSCAMOS SI CON ESTE DOCUMENTO YA FUERON SELECCIONADAS ALGUNAS TARIMAS
                            'SI YA HAY TARIMAS YA NO AGREGA MAS
                            Set RBuscaDocumento = New ADODB.Recordset
                                Call Abrir_Recordset(RBuscaDocumento, "Select Documento From DetalleSalidasInventario Where Documento = " & VDocumento)
                                
                                'SI ENCUENTRA DATOS
                                If RBuscaDocumento.RecordCount > 0 Then
                                'NO SELECCIONA NINGUNA TARIMA
                                        mensaje = MsgBox("Ya hay tarimas con este documento, Quiere Agregar Tarimas De Este Batch " & TxtBat.Text & "  Y Linea " & TxtLin.Text, vbYesNo + vbInformation, "Informacion")
                                        'SI CONTESTA QUE SI AGREGA LAS TARIMAS, Y ES POSIBLE REPETIRLAS¡
                                        If mensaje = vbYes Then
                                                
                                                'SELECCIONA TODAS LAS TARIMAS DE ENTRADAS DE PRODUCTO TERMINADO DE ACUERDO AL BATCH
                                                'Y QUE EL SALDO SEA MAYOR QUE CERO Y LA LINEA A LA QUE PERTENECEN
                                                Set RBuscaEntradasInventario = New ADODB.Recordset
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Call Abrir_Recordset(RBuscaEntradasInventario, "Select FichaTecnica, Tarima, Linea, FechaProduccion, Batch, Calidad, Bodega, Saldo From DetalleEntradasInventario Where Batch = " & VBatch & " And Saldo > 0 And Linea = '" & VLinea & "'")
                                                    Else 'ORACLE
                                                        Call Abrir_Recordset(RBuscaEntradasInventario, "Select FichaTecnica, Tarima, Linea, FechaProduccion, Batch, Calidad, Bodega, Saldo From DetalleEntradasInventario Where Batch = " & VBatch & " And Saldo > 0 And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                                    End If
                                                        
                                                        'INICIA LA TRANSACCION
                                                        Conexion.BeginTrans
                                                        
                                                        'CREA UN CICLO CON LOS DATOS DE PRODUCCION DEL BATCH
                                                        Do Until RBuscaEntradasInventario.EOF
                                                                            VTexto2 = VDocumento & ", "
                                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                                VTexto2 = VTexto2 & "#" & Format(RBuscaEntradasInventario!FechaProduccion, "mm/dd/yyyy") & "#, '" 'FECHA
                                                                            Else 'ORACLE
                                                                                VTexto2 = VTexto2 & "To_Date('" & RBuscaEntradasInventario!FechaProduccion & "', 'dd/mm/yyyy')" & ", '"  'FECHA
                                                                            End If
                                                                            VTexto2 = VTexto2 & RBuscaEntradasInventario!Linea & "', '"
                                                                            VTexto2 = VTexto2 & RBuscaEntradasInventario!FichaTecnica & "', "
                                                                            VTexto2 = VTexto2 & RBuscaEntradasInventario!Tarima & ", "
                                                                            VTexto2 = VTexto2 & VBatch & ", '"
                                                                            VTexto2 = VTexto2 & RBuscaEntradasInventario!Calidad & "', '"
                                                                            VTexto2 = VTexto2 & RBuscaEntradasInventario!Bodega & "', "
                                                                            VTexto2 = VTexto2 & RBuscaEntradasInventario!Saldo
                                                                        
                                                                            Conexion.Execute "Insert Into DetalleSalidasInventario Values(" & VTexto2 & ")"
                                                                            
                                                                            If Err = -2147217873 Then
                                                                                
                                                                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                                                                Conexion.RollbackTrans
                                                                                MousePointer = 0
                                                                                MsgBox "Error, No Se Terminaron De Agregar Todos Los Bultos/Tarimas Del Batch y Linea " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                                                Exit Sub
                                                                            End If
                                                            'SE MUEVE AL SIGUIENTE REGISTRO
                                                            RBuscaEntradasInventario.MoveNext
                                                        Loop
                                                        'TERMINA LA TRANSACCION
                                                        Conexion.CommitTrans
                                                        
                                                        
                                        End If
                                Else
                                                       
                                                'SELECCIONA TODAS LAS TARIMAS DE ENTRADAS DE PRODUCTO TERMINADO DE ACUERDO AL BATCH
                                                'Y QUE EL SALDO SEA MAYOR QUE CERO Y LA LINEA A LA QUE PERTENECEN
                                                Set RBuscaEntradasInventario = New ADODB.Recordset
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Call Abrir_Recordset(RBuscaEntradasInventario, "Select FichaTecnica, Tarima, Linea, FechaProduccion, Batch, Calidad, Bodega, Saldo From DetalleEntradasInventario Where Batch = " & VBatch & " And Saldo > 0 And Linea = '" & VLinea & "'")
                                                    Else
                                                        Call Abrir_Recordset(RBuscaEntradasInventario, "Select FichaTecnica, Tarima, Linea, FechaProduccion, Batch, Calidad, Bodega, Saldo From DetalleEntradasInventario Where Batch = " & VBatch & " And Saldo > 0 And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                                    End If
                                                    
                                                            If RBuscaEntradasInventario.RecordCount > 0 Then
                                                                    Conexion.BeginTrans
                                                                    
                                                                    'CREA UN CICLO CON LOS DATOS DE PRODUCCION DEL BATCH
                                                                    Do Until RBuscaEntradasInventario.EOF
                                                                            'AGREGA LAS TARIMAS DE ESE BATCH
                                                                                        VTexto2 = VDocumento & ", "
                                                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                                                            VTexto2 = VTexto2 & "#" & Format(RBuscaEntradasInventario!FechaProduccion, "mm/dd/yyyy") & "#, '" 'FECHA
                                                                                        Else 'ORACLE
                                                                                            VTexto2 = VTexto2 & "To_Date('" & RBuscaEntradasInventario!FechaProduccion & "', 'dd/mm/yyyy')" & ", '"  'FECHA
                                                                                        End If
                                                                                        VTexto2 = VTexto2 & RBuscaEntradasInventario!Linea & "', '"
                                                                                        VTexto2 = VTexto2 & RBuscaEntradasInventario!FichaTecnica & "', "
                                                                                        VTexto2 = VTexto2 & RBuscaEntradasInventario!Tarima & ", "
                                                                                        VTexto2 = VTexto2 & VBatch & ", '"
                                                                                        VTexto2 = VTexto2 & RBuscaEntradasInventario!Calidad & "', '"
                                                                                        VTexto2 = VTexto2 & RBuscaEntradasInventario!Bodega & "', "
                                                                                        VTexto2 = VTexto2 & RBuscaEntradasInventario!Saldo
                                                                                    
                                                                                        Conexion.Execute "Insert Into DetalleSalidasInventario Values(" & VTexto2 & ")"
                                                                                            
                                                                                        If Err = -2147217873 Then
                                                                                            'SI YA EXISTE LA TARIMA EN ESTE DOCUMENTO NO HACE NADA
                                                                                        ElseIf Err <> -2147217873 And Err <> 0 Then
                                                                                            Conexion.RollbackTrans
                                                                                            MousePointer = 0
                                                                                            MsgBox "Error, No Se Terminaron De Agregar Todos Los Bultos/Tarimas Del Batch y Linea " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                                                                            Exit Sub
                                                                                        End If
                                                                                         
                                                                            
                                                                        'SE MUEVE AL SIGUIENTE REGISTRO
                                                                        RBuscaEntradasInventario.MoveNext
                                                                    Loop
                                                                    
                                                                    Conexion.CommitTrans
                                                                    
                                                            End If
                                End If 'SI ENCUENTRA DATOS
                                
                            End If 'TERMINA EL DE BUSCAR TARIMAS
                        
                        End If 'TERMINA SI BATCH ES CERO
        
                        REncabezado.Requery
                        REncabezado.MoveFirst
                        REncabezado.Find ("Documento = " & VDocumento)
                    
                                    Llena_CamposEncabezado
                    
                                                         
                                    Set RDetalle = New ADODB.Recordset
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & VDocumento)
                                            Else 'ORACLE
                                                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & VDocumento)
                                            End If
                                                Llena_CamposDetalle
                                                Set DbGridDetalle.DataSource = RDetalle

    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Visible = True
    FrameDetalle.Enabled = True
    FrameEncabezado.Enabled = False
    
    DbGridDetalle.Visible = True
    DbGridDetalle.AllowDelete = True
    DbGridDetalle.AllowUpdate = True
    
    'ESCONDE EL DATA
    CmdBotones2.Item(1).Visible = False
    CmdBotones2.Item(2).Visible = False
    CmdBotones2.Item(3).Visible = False
    CmdBotones2.Item(4).Visible = False
    
        
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = False
    BotonesVisiblesEncabezado
    
    TabDespachos.Tab = 1
        
    CmdAgregar2.SetFocus
    
MousePointer = 0
    
End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11
                    'MUESTRA EL REPORTE
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioSalidas.rpt"
                    Else
                        GNombreReporte = "InventarioSalidasO.rpt"
                    End If
                    GCriteriaReporte = "{EncabezadoSalidasInventario.Documento} = " & TxtDocIng.Text
                    FrmReporte.Show
                   
MousePointer = 0

End Sub

Private Sub CmdImprimirCedulas_Click()
On Error Resume Next
        
        MousePointer = 11
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "CedulaMateriaPrimaSalidas.rpt"
                Else
                    GNombreReporte = "CedulaMateriaPrimaSalidasO.rpt"
                End If
                GCriteriaReporte = "{EncabezadoSalidasInventario.Documento} = " & TxtDocIng
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
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    CmdBotones2.Item(1).Visible = True
    CmdBotones2.Item(2).Visible = True
    CmdBotones2.Item(3).Visible = True
    CmdBotones2.Item(4).Visible = True
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = False
    'FrameDetalle.Visible = False
    FrameEncabezado.Enabled = True
                
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesVisiblesEncabezado
    
    DbGridDetalle.AllowDelete = False
    DbGridDetalle.AllowUpdate = False
    
    TabDespachos.Tab = 0
    
    

    
    If Err <> 0 Then
        MsgBox Err.Description
    End If
    
    
                
End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub


Private Sub Command2_Click()

End Sub

Private Sub DBGridBusqueda_DblClick()
    'BODEGA
    If BCliente = True Then
        TxtCli.Text = DBGridBusqueda.Columns(0)
        TxtCli.SetFocus
    'BODEGA DETALLE
    ElseIf BBodegaDetalle = True Then
        TxtBod.Text = DBGridBusqueda.Columns(0)
        TxtBod.SetFocus
    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        TxtCodPro.Text = DBGridBusqueda.Columns(0)
        TxtCodPro.SetFocus
    'TRANSPORTISTAS
    ElseIf BTransportistas = True Then
        TxtTra.Text = DBGridBusqueda.Columns(0)
        TxtTra.SetFocus
    'TIPO DE DOCUMENTO
    ElseIf BDocumento = True Then
        TxtTipDoc.Text = DBGridBusqueda.Columns(0)
        TxtTipDoc.SetFocus
    'TARIMA
    ElseIf BTarima = True Then
        MskFecPro.Text = DBGridBusqueda.Columns(0)
        TxtLin.Text = DBGridBusqueda.Columns(1)
        TxtCodPro.Text = DBGridBusqueda.Columns(2)
        TxtTar.Text = DBGridBusqueda.Columns(3)
        TxtBod.Text = DBGridBusqueda.Columns(4)
        TxtCanPro.Text = DBGridBusqueda.Columns(6)
        TxtCal.Text = DBGridBusqueda.Columns(7)
        TxtBat.Text = DBGridBusqueda.Columns(8)
        TxtTar.SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
If KeyAscii = 43 Then
    'BODEGA
    If BCliente = True Then
        TxtCli.Text = DBGridBusqueda.Columns(0)
        TxtCli.SetFocus
    'BODEGA DETALLE
    ElseIf BBodegaDetalle = True Then
        TxtBod.Text = DBGridBusqueda.Columns(0)
        TxtBod.SetFocus
    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        TxtCodPro.Text = DBGridBusqueda.Columns(0)
        TxtCodPro.SetFocus
    'TRANSPORTISTAS
    ElseIf BTransportistas = True Then
        TxtTra.Text = DBGridBusqueda.Columns(0)
        TxtTra.SetFocus
    'TIPO DE DOCUMENTO
    ElseIf BDocumento = True Then
        TxtTipDoc.Text = DBGridBusqueda.Columns(0)
        TxtTipDoc.SetFocus
    'TARIMA
    ElseIf BTarima = True Then
        MskFecPro.Text = DBGridBusqueda.Columns(0)
        TxtLin.Text = DBGridBusqueda.Columns(1)
        TxtCodPro.Text = DBGridBusqueda.Columns(2)
        TxtTar.Text = DBGridBusqueda.Columns(3)
        TxtBod.Text = DBGridBusqueda.Columns(4)
        TxtCanPro.Text = DBGridBusqueda.Columns(6)
        TxtCal.Text = DBGridBusqueda.Columns(7)
        TxtBat.Text = DBGridBusqueda.Columns(8)
        TxtTar.SetFocus
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
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoSalidasInventario Where Fecha >= #" & Format((Date - 730), "mm/dd/yyyy") & "# And Fecha <= #" & Format(Date, "mm/dd/yyyy") & "#")
                        Else 'ORACLE
                            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoSalidasInventario Where Fecha >= To_Date('" & (Date - 730) & "', 'dd/mm/yyyy') And Fecha <= To_Date('" & Date & "', 'dd/mm/yyyy')")
                        End If
                
                REncabezado.MoveLast
                
                Llena_CamposEncabezado
                
    
                Set RDetalle = New ADODB.Recordset
                
                        If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & TxtDocIng.Text)
                        Else 'ORACLE
                                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.Cantidad From DetalleSalidasInventario D Where D.Documento = " & TxtDocIng.Text)
                        End If
                            Llena_CamposDetalle
                            Set DbGridDetalle.DataSource = RDetalle
                
                
                If Err <> 0 Then
                    MsgBox Err.Description
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
Private Sub MskFecPro_GotFocus()
    MskFecPro.SelStart = 0
    MskFecPro.SelLength = Len(MskFecPro.Text)
End Sub
Private Sub MskFecPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TabDespachos_Click(PreviousTab As Integer)
                        If TabDespachos.Tab = 1 Then
                                    'CUENTA LAS TARIMAS QUE HAN SALIDO
                                    Set RCuentaTarimas = New ADODB.Recordset
                                        Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From DetalleSalidasInventario Where Documento = " & TxtDocIng.Text)
                                        If RCuentaTarimas.RecordCount > 0 Then
                                            TxtCueTar.Text = RCuentaTarimas(0)
                                        Else
                                            TxtCueTar.Text = 0
                                        End If
                        End If
End Sub

Private Sub TxtBarra_GotFocus()
        TxtBarra.SelStart = 0
        TxtBarra.SelLength = Len(TxtBarra.Text)
End Sub

Private Sub TxtBarra_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtBarra_LostFocus()
    If TxtBarra.Text <> "" Then
        'BUSCA LA BARRA
        Set RBuscaBarra = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBarra, "Select * From DetalleEntradasInventario Where Barra = '" & TxtBarra.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBarra, "Select * From DetalleEntradasInventario Where UPPER(Barra) = '" & UCase(TxtBarra.Text) & "'")
            End If
            
            If RBuscaBarra.RecordCount > 0 Then
                'SI ENCUENTRA LA BARRA JALA TODOS LOS DATOS
                MskFecPro.Text = RBuscaBarra!FechaProduccion
                TxtLin.Text = RBuscaBarra!Linea
                TxtCodPro.Text = RBuscaBarra!FichaTecnica
                TxtTar.Text = RBuscaBarra!Tarima
                TxtBat.Text = RBuscaBarra!Batch
                TxtCal.Text = RBuscaBarra!Calidad
                TxtBod.Text = RBuscaBarra!Bodega
                TxtCanPro.Text = RBuscaBarra!Saldo
                'EJECUTA EL BOTON DE GRABAR
                Call CmdGrabar2_Click
                'EJECUTA EL BOTON DE AGREGAR
                Call CmdAgregar2_Click
                TxtBarra.Text = ""
            Else
                MskFecPro.Text = ""
                TxtLin.Text = ""
                TxtCodPro.Text = ""
                TxtTar.Text = ""
                TxtBat.Text = ""
                TxtCal.Text = ""
                TxtBod.Text = ""
                TxtCanPro.Text = ""
                MsgBox "Codigo De Barra No Existe En Inventario", vbOKOnly + vbInformation, "Informacion"
                TxtBarra.Text = ""
                TxtBarra.SetFocus
            End If
    End If
End Sub

Private Sub Txtbat_GotFocus()
    TxtBat.SelStart = 0
    TxtBat.SelLength = Len(TxtBat.Text)
End Sub
Private Sub Txtbat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub TxtBatch_GotFocus()
    TxtBatch.SelStart = 0
    TxtBatch.SelLength = Len(TxtBatch.Text)
End Sub
Private Sub TxtBatch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub

Private Sub TxtBod_Change()
        Set RBuscaBodega = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBod.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBod.Text) & "'")
            End If
            If RBuscaBodega.RecordCount > 0 Then
                LblBod.Caption = RBuscaBodega!Descripcion
            Else
                LblBod.Caption = ""
            End If
End Sub

Private Sub TxtBod_DblClick()
            Set RBusqueda = New ADODB.Recordset
            BCliente = False
            BProducto = False
            BBodegaDetalle = True
            BTransportistas = False
            BDocumento = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Order by CodigoBodega")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBod_GotFocus()
            TxtBod.SelStart = 0
            TxtBod.SelLength = Len(TxtBod.Text)
End Sub

Private Sub TxtBod_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
               SendKeys "{tab}"
            End If
            
            If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                BCliente = False
                BProducto = False
                BBodegaDetalle = True
                BTransportistas = False
                BDocumento = False
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Order by CodigoBodega")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
            End If
End Sub

Private Sub TxtCal_GotFocus()
        TxtCal.SelStart = 0
        TxtCal.SelLength = Len(TxtCal.Text)
End Sub

Private Sub TxtCal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtCli_Change()
        Set RBuscaCliente = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where CodigoCliente = '" & TxtCli.Text & "'")
        Else
            Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where UPPER(CodigoCliente) = '" & UCase(TxtCli.Text) & "'")
        End If
            If RBuscaCliente.RecordCount > 0 Then
                LblCli.Caption = RBuscaCliente!Descripcion
            Else
                LblCli.Caption = ""
            End If
End Sub
Private Sub TxtCli_DblClick()
        BCliente = True
        BProducto = False
        BBodegaDetalle = False
        BTransportistas = False
        BDocumento = False
        BTarima = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Order by CodigoCliente")
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
        BCliente = True
        BProducto = False
        BBodegaDetalle = False
        BTransportistas = False
        BDocumento = False
        BTarima = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Order by CodigoCliente")
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "4000"
    End If
End Sub
Private Sub Txtbuscar_Change()
    Set RBusqueda = New ADODB.Recordset
    'BODEGA
    If BBodegaDetalle = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion from BodegasInventario Where CodigoBodega Like '%" & TxtBuscar.Text & "%' Order by CodigoBodega")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion from BodegasInventario Where UPPER(CodigoBodega) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoBodega")
                End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion from BodegasInventario Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by CodigoBodega")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion from BodegasInventario Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoBodega")
                End If
            
        End If

    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Esp_Tec Like '%" & TxtBuscar.Text & "%' And Activa = -1 Order by Esp_Tec")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscar.Text) & "%' And Activa = -1 Order by Esp_Tec")
                End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Descrip Like '%" & TxtBuscar.Text & "%' And Activa = -1 Order by Esp_Tec")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' And Activa = -1 Order by Esp_Tec")
                End If
        End If
    'CLIENTES
    ElseIf BCliente = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where CodigoCliente Like '%" & TxtBuscar.Text & "%' Order by CodigoCliente")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where UPPER(CodigoCliente) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoCliente")
                End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by CodigoCliente")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoCliente")
                End If
            
        End If
    'TRANSPORTISTAS
    ElseIf BTransportistas = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Transportistas Where Codigo Like '%" & TxtBuscar.Text & "%' Order by Codigo")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Codigo")
                End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by Codigo")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Codigo")
                End If
            
        End If
    End If
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "4000"
End Sub
Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub
Private Sub TxtCanPro_GotFocus()
    TxtCanPro.SelStart = 0
    TxtCanPro.SelLength = Len(TxtCanPro.Text)
End Sub
Private Sub TxtCanPro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub
Private Sub TxtCodPro_Change()
            Set RBuscaProducto = New ADODB.Recordset
                 If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProducto, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCodPro.Text & "'")
                 Else
                    Call Abrir_Recordset(RBuscaProducto, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodPro.Text) & "'")
                 End If
                 If RBuscaProducto.RecordCount > 0 Then
                        TxtDesPro.Text = RBuscaProducto(0)
                 Else
                        TxtDesPro.Text = ""
                 End If
End Sub

Private Sub TxtCodPro_DblClick()
            BCliente = False
            BProducto = True
            BBodegaDetalle = False
            BTransportistas = False
            BTarima = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Activa = -1 Order by Esp_Tec")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
End Sub

Private Sub TxtCodPro_KeyPress(KeyAscii As Integer)
            'SI PRECIONA ENTER
            If KeyAscii = 13 Then
               SendKeys "{tab}"
            End If
            'SI PRECIONA LA TECLA DE SIGNO +
            If KeyAscii = 43 Then
               BCliente = False
               BProducto = True
               BBodegaDetalle = False
               BTransportistas = False
               BTarima = False
               FrameBuscar.Visible = True
               TxtBuscar.SetFocus
               Set RBusqueda = New ADODB.Recordset
               Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Activa = -1 Order by Esp_Tec")
               Set DBGridBusqueda.DataSource = RBusqueda
               DBGridBusqueda.Columns(1).Width = "4000"
            End If
End Sub


Private Sub TxtDesPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtDocIng_GotFocus()
        TxtDocIng.SelStart = 0
        TxtDocIng.SelLength = Len(TxtDocIng.Text)
End Sub

Private Sub TxtDocing_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtLin_Change()
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLin.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
            End If
                
            If RBuscaLinea.RecordCount > 0 Then
                LblLin.Caption = RBuscaLinea!Descrip
            Else
                LblLin.Caption = ""
            End If
End Sub

Private Sub TxtLin_GotFocus()
        TxtLin.SelStart = 0
        TxtLin.SelLength = Len(TxtLin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub


Private Sub TxtLinea_Change()
        
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLinea.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
            End If
        
        
            If RBuscaLinea.RecordCount > 0 Then
                LblLinea.Caption = RBuscaLinea!Descrip
            Else
                LblLinea.Caption = ""
            End If

End Sub

Private Sub TxtLinea_GotFocus()
        TxtLinea.SelStart = 0
        TxtLinea.SelLength = Len(TxtLinea.Text)
End Sub

Private Sub TxtLinea_KeyPress(KeyAscii As Integer)
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

Private Sub TxtTar_DblClick()
        'REVISA SI LA TARIMA EXISTE EN LA ENTRADAS DE PRODUCTO TERMINADO
        Set RBuscaTarimas = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTarimas, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Bodega, B.Descripcion, D.Saldo, D.Calidad, D.Batch From DetalleEntradasInventario D, BodegasInventario B Where D.FichaTecnica = '" & TxtCodPro.Text & "' And D.Saldo > 0 And D.Bodega = B.CodigoBodega Order By D.FechaProduccion, D.Tarima")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaTarimas, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Bodega, B.Descripcion, D.Saldo, D.Calidad, D.Batch From DetalleEntradasInventario D, BodegasInventario B Where UPPER(D.FichaTecnica) = '" & UCase(TxtCodPro.Text) & "' And D.Saldo > 0 And D.Bodega = B.CodigoBodega Order By D.FechaProduccion, D.Tarima")
                End If
    
                If RBuscaTarimas.RecordCount > 0 Then
                        BCliente = False
                        BProducto = False
                        BBodegaDetalle = False
                        BTransportistas = False
                        BTarima = True
                        FrameBuscar.Visible = True
                        TxtBuscar.SetFocus
                        Set DBGridBusqueda.DataSource = RBuscaTarimas
                        DBGridBusqueda.Columns(0).Width = "1000"
                        DBGridBusqueda.Columns(1).Width = "500"
                        DBGridBusqueda.Columns(2).Width = "1500"
                        DBGridBusqueda.Columns(3).Width = "500"
                        DBGridBusqueda.Columns(4).Width = "500"
                        DBGridBusqueda.Columns(5).Width = "2500"
                        DBGridBusqueda.Columns(6).Width = "1200"
                        DBGridBusqueda.Columns(6).NumberFormat = "#,###,##0.00"
                        DBGridBusqueda.Columns(7).Width = "300"
                        DBGridBusqueda.Columns(8).Width = "800"
                Else
                    
                End If

End Sub

Private Sub Txttar_GotFocus()
        TxtTar.SelStart = 0
        TxtTar.SelLength = Len(TxtTar.Text)
End Sub

Private Sub Txttar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                'REVISA SI LA TARIMA EXISTE EN LA ENTRADAS DE PRODUCTO TERMINADO
                Set RBuscaTarimas = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaTarimas, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Bodega, B.Descripcion, D.Saldo, D.Calidad, D.Batch From DetalleEntradasInventario D, BodegasInventario B Where D.FichaTecnica = '" & TxtCodPro.Text & "' And D.Saldo > 0 And D.Bodega = B.CodigoBodega Order By D.FechaProduccion, D.Tarima")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaTarimas, "Select D.FechaProduccion, D.Linea, D.FichaTecnica, D.Tarima, D.Bodega, B.Descripcion, D.Saldo, D.Calidad, D.Batch From DetalleEntradasInventario D, BodegasInventario B Where UPPER(D.FichaTecnica) = '" & UCase(TxtCodPro.Text) & "' And D.Saldo > 0 And D.Bodega = B.CodigoBodega Order By D.FechaProduccion, D.Tarima")
                        End If
            
                        If RBuscaTarimas.RecordCount > 0 Then
                                BCliente = False
                                BProducto = False
                                BBodegaDetalle = False
                                BTransportistas = False
                                BTarima = True
                                FrameBuscar.Visible = True
                                TxtBuscar.SetFocus
                                Set DBGridBusqueda.DataSource = RBuscaTarimas
                                DBGridBusqueda.Columns(0).Width = "1000"
                                DBGridBusqueda.Columns(1).Width = "500"
                                DBGridBusqueda.Columns(2).Width = "1500"
                                DBGridBusqueda.Columns(3).Width = "500"
                                DBGridBusqueda.Columns(4).Width = "500"
                                DBGridBusqueda.Columns(5).Width = "2500"
                                DBGridBusqueda.Columns(6).Width = "1200"
                                DBGridBusqueda.Columns(6).NumberFormat = "#,###,##0.00"
                                DBGridBusqueda.Columns(7).Width = "300"
                                DBGridBusqueda.Columns(8).Width = "800"
                        Else
                            
                        End If

        End If

End Sub

Public Sub BotonesVisiblesEncabezado()
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

Private Sub TxtTar_LostFocus()

        If IsNumeric(TxtTar.Text) Then
        
            If MskFecPro.Text = "" Then
                    Set RBuscaTarima = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaTarima, "Select FechaProduccion From DetalleEntradasInventario Where Tarima = " & TxtTar.Text & " And FichaTecnica = '" & TxtCodPro.Text & "' And Linea = '" & TxtLin.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaTarima, "Select FechaProduccion From DetalleEntradasInventario Where Tarima = " & TxtTar.Text & " And UPPER(FichaTecnica) = '" & UCase(TxtCodPro.Text) & "' And Linea = '" & TxtLin.Text & "'")
                            End If
                        
                        If RBuscaTarima.RecordCount > 0 Then
                                MskFecPro.Text = RBuscaTarima!FechaProduccion
                        Else
                            MsgBox "Ficha Tecnica Con Este Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
            End If
            
            If TxtTar.Text = 0 Then
            
            Else
                'REVISA SI LA TARIMA EXISTE EN LA ENTRADAS DE PRODUCTO TERMINADO
                Set RBuscaEntradasInventario = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaEntradasInventario, "Select Saldo, Bodega From DetalleEntradasInventario Where FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTar.Text & " And Linea = '" & TxtLin.Text & "' And FichaTecnica = '" & TxtCodPro.Text & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaEntradasInventario, "Select Saldo, Bodega From DetalleEntradasInventario Where FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTar.Text & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtCodPro.Text) & "'")
                        End If
            
                        If RBuscaEntradasInventario.RecordCount > 0 Then
                            TxtCanPro.Text = RBuscaEntradasInventario!Saldo
                            TxtBod.Text = RBuscaEntradasInventario!Bodega
                        Else
                            TxtCanPro.Text = 0
                            TxtBod.Text = ""
                            MsgBox "Bulto o Tarima No Exite", vbOKOnly + vbInformation, "Informacion"
                        End If
            End If
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

End Sub

Private Sub TxtTipDoc_Change()
        Set RBuscaDocumento = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaDocumento, "Select Descripcion From Documentos Where CodigoDocumento = '" & TxtTipDoc.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaDocumento, "Select Descripcion From Documentos Where UPPER(CodigoDocumento) = '" & UCase(TxtTipDoc.Text) & "'")
            End If
                If RBuscaDocumento.RecordCount > 0 Then
                    LblDocumento.Caption = RBuscaDocumento!Descripcion
                Else
                    LblDocumento.Caption = ""
                End If
End Sub

Private Sub TxtTipDoc_DblClick()
        BCliente = False
        BProducto = False
        BBodegaDetalle = False
        BTransportistas = False
        BDocumento = True
        BTarima = False
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos")
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "4000"
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
                    BCliente = False
                    BProducto = False
                    BBodegaDetalle = False
                    BTransportistas = False
                    BDocumento = True
                    BTarima = False
                    FrameBuscar.Visible = True
                    TxtBuscar.SetFocus
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
            End If

End Sub

Private Sub TxtTra_Change()
            Set RBuscaTransportista = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTransportista, "Select Descripcion From Transportistas Where Codigo = '" & TxtTra.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaTransportista, "Select Descripcion From Transportistas Where UPPER(Codigo) = '" & UCase(TxtTra.Text) & "'")
                End If
                If RBuscaTransportista.RecordCount > 0 Then
                    LblTra.Caption = RBuscaTransportista!Descripcion
                Else
                    LblTra.Caption = ""
                End If
End Sub

Private Sub TxtTra_DblClick()
            BCliente = False
            BProducto = False
            BBodegaDetalle = False
            BTransportistas = True
            BDocumento = False
            BTarima = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtTra_GotFocus()
        TxtTra.SelStart = 0
        TxtTra.SelLength = Len(TxtTra.Text)
End Sub

Private Sub TxtTra_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BCliente = False
            BProducto = False
            BBodegaDetalle = False
            BTransportistas = True
            BDocumento = False
            BTarima = False
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas")
            Set DBGridBusqueda.DataSource = RBusqueda
            DBGridBusqueda.Columns(1).Width = "4000"
        End If

End Sub

Public Sub Llena_CamposEncabezado()
On Error Resume Next
            If REncabezado.RecordCount > 0 Then
                If IsNull(REncabezado!Documento) Then
                    TxtDocIng.Text = ""
                Else
                    TxtDocIng.Text = REncabezado!Documento
                End If
                If IsNull(REncabezado!fecha) Then
                    MskFec.Text = ""
                Else
                    MskFec.Text = REncabezado!fecha
                End If
                If IsNull(REncabezado!Cliente) Then
                    TxtCli.Text = ""
                Else
                    TxtCli.Text = REncabezado!Cliente
                End If
                If IsNull(REncabezado!Batch) Then
                    TxtBatch.Text = 0
                Else
                    TxtBatch.Text = REncabezado!Batch
                End If
                If IsNull(REncabezado!Linea) Then
                    TxtLinea.Text = ""
                Else
                    TxtLinea.Text = REncabezado!Linea
                End If
                If IsNull(REncabezado!CodigoTransportista) Then
                    TxtTra.Text = ""
                Else
                    TxtTra.Text = REncabezado!CodigoTransportista
                End If
                If IsNull(REncabezado!TipoDeDocumento) Then
                    TxtTipDoc.Text = ""
                Else
                    TxtTipDoc.Text = REncabezado!TipoDeDocumento
                End If
                If IsNull(REncabezado!NumeroDocumento) Then
                    TxtNumDoc.Text = ""
                Else
                    TxtNumDoc.Text = REncabezado!NumeroDocumento
                End If
                If IsNull(REncabezado!CargadoPor) Then
                    TxtTexto.Item(1).Text = ""
                Else
                    TxtTexto.Item(1).Text = REncabezado!CargadoPor
                End If
                If IsNull(REncabezado!EntregadoPor) Then
                    TxtTexto.Item(2).Text = ""
                Else
                    TxtTexto.Item(2).Text = REncabezado!EntregadoPor
                End If
                If IsNull(REncabezado!Conductor) Then
                    TxtTexto.Item(3).Text = ""
                Else
                    TxtTexto.Item(3).Text = REncabezado!Conductor
                End If
                If IsNull(REncabezado!PlacasCamion) Then
                    TxtTexto.Item(4).Text = ""
                Else
                    TxtTexto.Item(4).Text = REncabezado!PlacasCamion
                End If
                If IsNull(REncabezado!PlacasFurgon) Then
                    TxtTexto.Item(5).Text = ""
                Else
                    TxtTexto.Item(5).Text = REncabezado!PlacasFurgon
                End If
                If IsNull(REncabezado!Observaciones) Then
                    TxtTexto.Item(6).Text = ""
                Else
                    TxtTexto.Item(6).Text = REncabezado!Observaciones
                End If
                If IsNull(REncabezado!Requerido) Then
                    TxtReq.Text = ""
                Else
                    TxtReq.Text = REncabezado!Requerido
                End If
                If IsNull(REncabezado!Liberado) Then
                    TxtLib.Text = ""
                Else
                    TxtLib.Text = REncabezado!Liberado
                End If
                If IsNull(REncabezado!Estado) Then
                    TxtEstado.Text = ""
                Else
                    TxtEstado.Text = REncabezado!Estado
                End If
                
            Else
                TxtDocIng.Text = ""
                MskFec.Text = ""
                TxtCli.Text = ""
                TxtBatch.Text = 0
                TxtLinea.Text = ""
                TxtTra.Text = ""
                TxtTipDoc.Text = ""
                TxtNumDoc.Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(2).Text = ""
                TxtTexto.Item(3).Text = ""
                TxtTexto.Item(4).Text = ""
                TxtTexto.Item(5).Text = ""
                TxtTexto.Item(6).Text = ""
                TxtReq.Text = ""
                TxtLib.Text = ""
                TxtEstado.Text = ""
                
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
                If IsNull(RDetalle!FechaProduccion) Then
                    MskFecPro.Text = ""
                Else
                    MskFecPro.Text = RDetalle!FechaProduccion
                End If
                If IsNull(RDetalle!Linea) Then
                    TxtLin.Text = ""
                Else
                    TxtLin.Text = RDetalle!Linea
                End If
                If IsNull(RDetalle!FichaTecnica) Then
                    TxtCodPro.Text = ""
                Else
                    TxtCodPro.Text = RDetalle!FichaTecnica
                End If
                If IsNull(RDetalle!Tarima) Then
                    TxtTar.Text = 0
                Else
                    TxtTar.Text = RDetalle!Tarima
                End If
                If IsNull(RDetalle!Batch) Then
                    TxtBat.Text = 0
                Else
                    TxtBat.Text = RDetalle!Batch
                End If
                If IsNull(RDetalle!Calidad) Then
                    TxtCal.Text = ""
                Else
                    TxtCal.Text = RDetalle!Calidad
                End If
                If IsNull(RDetalle!Bodega) Then
                    TxtBod.Text = ""
                Else
                    TxtBod.Text = RDetalle!Bodega
                End If
                If IsNull(RDetalle!Cantidad) Then
                    TxtCanPro.Text = ""
                Else
                    TxtCanPro.Text = RDetalle!Cantidad
                End If
            Else
                TxtDocDet.Text = ""
                MskFecPro.Text = ""
                TxtLin.Text = ""
                TxtCodPro.Text = ""
                TxtTar.Text = 0
                TxtBat.Text = ""
                TxtCal.Text = ""
                TxtBod.Text = ""
                TxtCanPro.Text = 0
            End If
            
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Limpia_CamposEncabezado()
                TxtDocIng.Text = "0"
                MskFec.Text = ""
                TxtCli.Text = ""
                TxtBatch.Text = 0
                TxtLinea.Text = ""
                TxtTra.Text = ""
                TxtTipDoc.Text = ""
                TxtNumDoc.Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(2).Text = ""
                TxtTexto.Item(3).Text = ""
                TxtTexto.Item(4).Text = ""
                TxtTexto.Item(5).Text = ""
                TxtTexto.Item(6).Text = ""
                TxtReq.Text = ""
                TxtLib.Text = ""
                TxtEstado.Text = ""
End Sub

Public Sub Limpia_CamposDetalle()
                TxtDocDet.Text = ""
                MskFecPro.Text = ""
                TxtLin.Text = ""
                TxtCodPro.Text = ""
                TxtTar.Text = 0
                TxtBat.Text = 0
                TxtCal.Text = ""
                TxtBod.Text = ""
                TxtCanPro.Text = 0
End Sub




