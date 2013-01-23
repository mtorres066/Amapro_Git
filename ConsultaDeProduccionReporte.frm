VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ConsultaDeProduccionReporte 
   BackColor       =   &H000080FF&
   Caption         =   "Consulta De Reporte De Produccion"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "ConsultaDeProduccionReporte.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdGenera 
      Height          =   615
      Left            =   10560
      Picture         =   "ConsultaDeProduccionReporte.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   0
      Width           =   615
   End
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
      Height          =   8175
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6975
         Left            =   120
         TabIndex        =   17
         Top             =   1080
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   12303
         _Version        =   393216
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
         Height          =   615
         Left            =   7440
         Picture         =   "ConsultaDeProduccionReporte.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   16
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   15
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      Caption         =   "Tipo De Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   1812
      Begin VB.OptionButton OptTodos 
         BackColor       =   &H000080FF&
         Caption         =   "Todos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptLinea 
         BackColor       =   &H000080FF&
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
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptGrupo 
         BackColor       =   &H000080FF&
         Caption         =   "Grupo De Linea"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.TextBox TxtLinea 
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
      Height          =   285
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   6
      ToolTipText     =   "doble click o signo '+' para ayuda"
      Top             =   480
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.CommandButton CmdSalida 
      Height          =   615
      Left            =   11160
      Picture         =   "ConsultaDeProduccionReporte.frx":25FE
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   0
      Width           =   615
   End
   Begin MSComCtl2.DTPicker DtpFecFin 
      Height          =   255
      Left            =   7200
      TabIndex        =   1
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   50266115
      CurrentDate     =   37248
   End
   Begin MSComCtl2.DTPicker DtpFecIni 
      Height          =   255
      Left            =   5400
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   50266115
      CurrentDate     =   37248
   End
   Begin TabDlg.SSTab TabGeneral 
      Height          =   7452
      Left            =   0
      TabIndex        =   19
      Top             =   840
      Width           =   11892
      _ExtentX        =   20981
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      BackColor       =   33023
      TabCaption(0)   =   "Produccion"
      TabPicture(0)   =   "ConsultaDeProduccionReporte.frx":2B19
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "DbGridGerencia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "DbGridParos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "DbGridLineas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DbGridMes"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Ordenes Abiertas"
      TabPicture(1)   =   "ConsultaDeProduccionReporte.frx":2E33
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TabProduccion"
      Tab(1).ControlCount=   1
      Begin MSDataGridLib.DataGrid DbGridMes 
         Height          =   2775
         Left            =   5640
         TabIndex        =   23
         ToolTipText     =   "click en encabezado de columna para indexar"
         Top             =   4560
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8438015
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
            Weight          =   700
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
         Caption         =   "Produccion Por Mes y Año"
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "Expr1001"
            Caption         =   "Año"
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
            DataField       =   "Expr1000"
            Caption         =   "Mes"
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
         BeginProperty Column04 
            DataField       =   "Expr1004"
            Caption         =   "PC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "Expr1005"
            Caption         =   "PNC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Expr1006"
            Caption         =   "Desp."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
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
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   345.26
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   329.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1920.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   780.095
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
               ColumnWidth     =   780.095
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DbGridLineas 
         Height          =   2775
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "click en encabezado de columna para indexar"
         Top             =   4560
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   4895
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12632319
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
            Weight          =   700
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
         Caption         =   "Produccion Por Linea"
         ColumnCount     =   5
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "Expr1002"
            Caption         =   "PC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Expr1003"
            Caption         =   "PNC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Expr1004"
            Caption         =   "Desp."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
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
               ColumnWidth     =   569.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1695.118
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   764.787
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid DbGridParos 
         Height          =   3735
         Left            =   7920
         TabIndex        =   21
         ToolTipText     =   "click en encabezado de columna para indexar"
         Top             =   720
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
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
            Weight          =   700
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
         Caption         =   "Horas De Paro"
         ColumnCount     =   4
         BeginProperty Column00 
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
         BeginProperty Column01 
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
         BeginProperty Column02 
            DataField       =   "Tipo"
            Caption         =   "Tipo"
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
            DataField       =   "Expr1003"
            Caption         =   "Horas"
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
               ColumnWidth     =   555.024
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   1739.906
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   195.024
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   689.953
            EndProperty
         EndProperty
      End
      Begin TabDlg.SSTab TabProduccion 
         Height          =   6612
         Left            =   -74880
         TabIndex        =   20
         Top             =   720
         Width           =   11652
         _ExtentX        =   20558
         _ExtentY        =   11668
         _Version        =   393216
         Tab             =   2
         TabHeight       =   706
         BackColor       =   16777215
         ForeColor       =   16711680
         TabCaption(0)   =   "Resumen Ordenes Abiertas"
         TabPicture(0)   =   "ConsultaDeProduccionReporte.frx":4B3D
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "DataOrden"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Detalle Ordenes Abiertas"
         TabPicture(1)   =   "ConsultaDeProduccionReporte.frx":4B59
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "DbGridOrden"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "Inventario"
         TabPicture(2)   =   "ConsultaDeProduccionReporte.frx":4B75
         Tab(2).ControlEnabled=   -1  'True
         Tab(2).Control(0)=   "DBGridInvProTer"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).ControlCount=   1
         Begin MSDataGridLib.DataGrid DataOrden 
            Height          =   6015
            Left            =   -74880
            TabIndex        =   4
            ToolTipText     =   "Click en encabezado de columna para indexar"
            Top             =   480
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10610
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12640511
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
            ColumnCount     =   8
            BeginProperty Column00 
               DataField       =   "Documento"
               Caption         =   "Orden"
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
               DataField       =   "FechaApertura"
               Caption         =   "Fecha Apertura"
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
               DataField       =   "FechaEntrega"
               Caption         =   "Fecha Entrega"
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
               DataField       =   "Expr1005"
               Caption         =   "Requerido"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4106
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column06 
               DataField       =   "Expr1006"
               Caption         =   "Entregado"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4106
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column07 
               DataField       =   "Expr1007"
               Caption         =   "Saldo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
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
                  ColumnWidth     =   1094.74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1184.882
               EndProperty
               BeginProperty Column02 
                  ColumnWidth     =   3734.929
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   1065.26
               EndProperty
               BeginProperty Column04 
                  ColumnWidth     =   1080
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
                  ColumnWidth     =   840.189
               EndProperty
               BeginProperty Column06 
                  Alignment       =   1
                  ColumnWidth     =   854.929
               EndProperty
               BeginProperty Column07 
                  Alignment       =   1
                  ColumnWidth     =   929.764
               EndProperty
            EndProperty
         End
         Begin MSDataGridLib.DataGrid DBGridInvProTer 
            Height          =   6015
            Left            =   120
            TabIndex        =   26
            ToolTipText     =   "Click en encabezado de columna para indexar"
            Top             =   480
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10610
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   12632319
            HeadLines       =   1
            RowHeight       =   15
            TabAcrossSplits =   -1  'True
            TabAction       =   2
            WrapCellPointer =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
         Begin MSDataGridLib.DataGrid DbGridOrden 
            Height          =   6015
            Left            =   -74880
            TabIndex        =   25
            ToolTipText     =   "Click en encabezado de columna para indexar"
            Top             =   480
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   10610
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   8438015
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
               Weight          =   700
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "Descrip"
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
            BeginProperty Column01 
               DataField       =   "Descripcion"
               Caption         =   "Pasada"
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
            BeginProperty Column03 
               DataField       =   "Requerido"
               Caption         =   "Requerido"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4106
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column04 
               DataField       =   "Entregado"
               Caption         =   "Entregado"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   4106
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column05 
               DataField       =   "Saldo"
               Caption         =   "Saldo"
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "#,##0"
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
                  ColumnWidth     =   2894.74
               EndProperty
               BeginProperty Column01 
                  ColumnWidth     =   1679.811
               EndProperty
               BeginProperty Column02 
               EndProperty
               BeginProperty Column03 
                  Alignment       =   1
               EndProperty
               BeginProperty Column04 
                  Alignment       =   1
               EndProperty
               BeginProperty Column05 
                  Alignment       =   1
               EndProperty
            EndProperty
         End
      End
      Begin MSDataGridLib.DataGrid DbGridGerencia 
         Height          =   3735
         Left            =   120
         TabIndex        =   24
         ToolTipText     =   "click en encabezado de columna para indexar"
         Top             =   720
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   6588
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   16761024
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
            Weight          =   700
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
         Caption         =   "Produccion Por Orden De Produccion"
         ColumnCount     =   5
         BeginProperty Column00 
            DataField       =   "Documento"
            Caption         =   "Orden"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "Descrip"
            Caption         =   "Producto"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "Expr1002"
            Caption         =   "PC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "Expr1003"
            Caption         =   "PNC"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "Expr1004"
            Caption         =   "Desp."
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2789.858
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnWidth     =   975.118
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   824.882
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   959.811
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label LblLinea 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
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
      Left            =   6600
      TabIndex        =   8
      Top             =   480
      Width           =   3735
   End
   Begin VB.Label LblDescripcion 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
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
      Left            =   4440
      TabIndex        =   7
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Al"
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
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Del"
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
      Left            =   5040
      TabIndex        =   2
      Top             =   120
      Width           =   300
   End
End
Attribute VB_Name = "ConsultaDeProduccionReporte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RTotal As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim ROrdenesResumen As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RProduccionxOrden As New ADODB.Recordset
Dim RParos As New ADODB.Recordset
Dim RProduccionxLinea As New ADODB.Recordset
Dim RProduccionxMes As New ADODB.Recordset

Dim RDetalleorden As New ADODB.Recordset
Dim RInventario As New ADODB.Recordset

Dim BLinea As Boolean
Dim BGrupo As Boolean
Dim Cont As Integer
Dim VTotalFilas As Integer
Dim VOrden As String


Private Sub CmdGenera_Click()
On Error Resume Next
MousePointer = 11

'_______________________________________________________________________________________________________________________
            'GRID DE FICHA TECNICA
                Set RProduccionxOrden = New ADODB.Recordset
                If OptTodos.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxOrden, "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos EP, DetalleProduccionPorOrden P, FichaTecnica F, EncabezadoOrdenProduccion EO Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Documento = P.Documento And P.Orden = EO.Documento And EO.FichaTecnica = F.ESP_TEC Group By EO.Documento, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxOrden, "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos EP, DetalleProduccionPorOrden P, FichaTecnica F, EncabezadoOrdenProduccion EO Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And EP.Documento = P.Documento And UPPER(P.Orden) = UPPER(EO.Documento) And UPPER(EO.FichaTecnica) = UPPER(F.ESP_TEC) Group By EO.Documento, F.Descrip")
                    End If
                ElseIf OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxOrden, "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos EP, DetalleProduccionPorOrden P, FichaTecnica F, EncabezadoOrdenProduccion EO, Lineas L Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And EP.Documento = P.Documento And P.Orden = EO.Documento And EO.FichaTecnica = F.ESP_TEC Group By EO.Documento, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxOrden, "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos EP, DetalleProduccionPorOrden P, FichaTecnica F, EncabezadoOrdenProduccion EO, Lineas L Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) = '" & UCase(TxtLinea.Text) & "' And EP.Documento = P.Documento And UPPER(P.Orden) = UPPER(EO.Documento) And UPPER(EO.FichaTecnica) = UPPER(F.ESP_TEC) Group By EO.Documento, F.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxOrden, "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos EP, DetalleProduccionPorOrden P, FichaTecnica F, EncabezadoOrdenProduccion EO, Lineas L Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' And EP.Documento = P.Documento And P.Orden = EO.Documento And EO.FichaTecnica = F.ESP_TEC Group By EO.Documento, F.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxOrden, "SELECT EO.Documento, F.Descrip, Sum(P.ProductoConforme), Sum(P.ProductoNoConforme), Sum(P.Desperdicio) From EncabezadoCapturaParos EP, DetalleProduccionPorOrden P, FichaTecnica F, EncabezadoOrdenProduccion EO, Lineas L Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And EP.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' And EP.Documento = P.Documento And P.Orden = EO.Documento And EO.FichaTecnica = F.ESP_TEC Group By EO.Documento, F.Descrip")
                    End If
                End If
            
                Set DbGridGerencia.DataSource = RProduccionxOrden
                
                If Err <> 0 Then
                    MsgBox Err.Number & Err.Description, vbOKOnly + vbCritical, "Error"
                    Err.Clear
                End If

'_______________________________________________________________________________________________________________________
            'EL GRID DE LINEAS
                Set RProduccionxLinea = New ADODB.Recordset
                
                If OptTodos.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxLinea, "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea Group By EP.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxLinea, "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & "  and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) Group By EP.Linea, L.Descrip")
                    End If
                ElseIf OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxLinea, "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By EP.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxLinea, "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) = '" & UCase(TxtLinea.Text) & "' Group By EP.Linea, L.Descrip")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxLinea, "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By EP.Linea, L.Descrip")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxLinea, "SELECT EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And UPPER(L.Linea) = '" & UCase(TxtLinea.Text) & "' Group By EP.Linea, L.Descrip")
                    End If
                End If
            
                Set DbGridLineas.DataSource = RProduccionxLinea
                
                If Err <> 0 Then
                    MsgBox Err.Number & Err.Description, vbOKOnly + vbCritical, "Error"
                    Err.Clear
                End If
                               
            
'_______________________________________________________________________________________________________________________
            'EL GRID DE MES
            
                Set RProduccionxMes = New ADODB.Recordset
                If OptTodos.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxMes, "SELECT month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea Group By month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip Order By Year(EP.Fecha)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxMes, "SELECT TO_Char(EP.Fecha,'YYYY'), To_Char(EP.Fecha,'MM'), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) Group By To_Char(EP.Fecha,'YYYY'), to_char(EP.Fecha,'MM'), EP.Linea, L.Descrip Order By To_Char(EP.Fecha,'YYYY'), To_Char(EP.Fecha,'MM')")
                    End If
                ElseIf OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxMes, "SELECT month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' Group By month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip Order By Year(EP.Fecha)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxMes, "SELECT TO_Char(EP.Fecha,'YYYY'), To_Char(EP.Fecha,'MM'), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) = '" & UCase(TxtLinea.Text) & "' Group By To_Char(EP.Fecha,'YYYY'), To_Char(EP.Fecha,'MM'), EP.Linea, L.Descrip Order By To_Char(EP.Fecha,'YYYY'), To_Char(EP.Fecha,'MM') ")
                    End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccionxMes, "SELECT month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Linea = '" & TxtLinea.Text & "' Group By month(EP.Fecha), Year(EP.Fecha), EP.Linea, L.Descrip Order By Year(EP.Fecha)")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccionxMes, "SELECT TO_Char(EP.Fecha,'YYYY'), To_Char(EP.Fecha,'MM'), EP.Linea, L.Descrip, sum(EP.ProductoConforme), Sum(EP.ProductoNoConforme), Sum(EP.Desperdicio) From EncabezadoCapturaParos EP, Lineas L Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And UPPER(L.Linea) = '" & UCase(TxtLinea.Text) & "' Group By To_Char(EP.Fecha,'YYYY'), To_Char(EP.Fecha,'MM'), EP.Linea, L.Descrip Order By To_Char(EP.Fecha,'YYYY'), To_Char(EP.Fecha,'MM')")
                    End If
                End If

                Set DbgridMes.DataSource = RProduccionxMes
                
                If Err <> 0 Then
                    MsgBox Err.Number & Err.Description, vbOKOnly + vbCritical, "Error"
                    Err.Clear
                End If
                
    '_______________________________________________________________________________________________________________________
            'PAROS
                Set RParos = New ADODB.Recordset
                If OptTodos.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo")
                    Else 'ORACLE
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) Group By EP.Linea, L.Descrip, P.Tipo")
                    End If
                ElseIf OptGrupo.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = L.Linea And L.Grupo = '" & TxtLinea.Text & "' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo")
                    Else 'ORACLE
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = UPPER(L.Linea) And UPPER(L.Grupo) = '" & UCase(TxtLinea.Text) & "' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) Group By EP.Linea, L.Descrip, P.Tipo")
                        End If
                ElseIf OptLinea.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos as EP, DetalleCapturaParos As DP, Lineas as L, Paros as P Where EP.Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "#  and EP.Fecha <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And EP.Linea = '" & TxtLinea.Text & "' And EP.Linea = L.Linea And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo Group By EP.Linea, L.Descrip, P.Tipo")
                    Else 'ORACLE
                        Call Abrir_Recordset(RParos, "SELECT EP.Linea, L.Descrip, P.Tipo, Sum(DP.Minutos/60) From EncabezadoCapturaParos EP, DetalleCapturaParos DP, Lineas L, Paros P Where EP.Fecha >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " and EP.Fecha <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = '" & UCase(TxtLinea.Text) & "' And EP.Linea = L.Linea And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) Group By EP.Linea, L.Descrip, P.Tipo")
                    End If
                End If
            
            Set DbGridParos.DataSource = RParos
            
            If Err <> 0 Then
                    MsgBox Err.Number & Err.Description, vbOKOnly + vbCritical, "Error"
                    Err.Clear
                End If
            
MousePointer = 0
        
End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False

End Sub

Private Sub CmdSalida_Click()
            Unload Me
End Sub

Private Sub DataOrden_DblClick()
On Error Resume Next

            'ASIGNAMOS A UNA VARIABLE LA ORDEN DE LA COLUMNA 1
            
            VOrden = DataOrden.Columns(0).Text
            
            DbGridOrden.Caption = "ORDEN " & Space(5) & VOrden
            DBGridInvProTer.Caption = "ORDEN: " & Space(5) & VOrden
                
            'EL GRID DE ORDEN DETALLE ABIERTAS
                        Set RDetalleorden = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDetalleorden, "Select L.Descrip, P.Descripcion, DO.Observaciones, DO.Requerido, DO.Entregado, DO.Saldo From DetalleOrdenProduccion DO, Lineas L, Pasadas P Where DO.Documento = '" & VOrden & "' And DO.Linea = L.Linea And DO.Pasada = P.Codigo")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDetalleorden, "Select L.Descrip, P.Descripcion, DO.Observaciones, DO.Requerido, DO.Entregado, DO.Saldo From DetalleOrdenProduccion DO, Lineas L, Pasadas P Where UPPER(DO.Documento) = '" & UCase(VOrden) & "' And UPPER(DO.Linea) = UPPER(L.Linea) And UPPER(DO.Pasada) = UPPER(P.Codigo)")
                            End If
                        Set DbGridOrden.DataSource = RDetalleorden
                        TabProduccion.Tab = 1
                        
           'INVENTARIO PRODUCTO TERMINADO
                        Set RInventario = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RInventario, "Select DE.Bodega, B.Descripcion, DE.FichaTecnica, F.Descrip, Count(DE.Saldo), Sum(DE.Saldo), Sum(DE.Saldo * F.PesoxUnidad / 1000) From DetalleEntradasInventario DE, Bodegasinventario B, FichaTecnica F Where DE.OrdenProduccion = '" & VOrden & "' And DE.Saldo > 0 And DE.Bodega = B.CodigoBodega And DE.FichaTecnica = F.Esp_Tec Group By DE.Bodega, B.Descripcion, DE.FichaTecnica, F.Descrip")
                            Else 'ORACLE
                                Call Abrir_Recordset(RInventario, "Select DE.Bodega, B.Descripcion, DE.FichaTecnica, F.Descrip, Count(DE.Saldo), Sum(DE.Saldo), Sum(DE.Saldo * F.PesoxUnidad / 1000) From DetalleEntradasInventario DE, BodegasInventario B, FichaTecnica F Where UPPER(DE.OrdenProduccion) = '" & UCase(VOrden) & "' And DE.Saldo > 0 And UPPER(DE.Bodega) = UPPER(B.CodigoBodega) And UPPER(DE.FichaTecnica) = UPPER(F.Esp_Tec) Group By DE.Bodega, B.Descripcion, DE.FichaTecnica, F.Descrip")
                            End If
                            Set DBGridInvProTer.DataSource = RInventario
                            
                            DBGridInvProTer.Columns(0).Width = "400"
                            DBGridInvProTer.Columns(1).Width = "2000"
                            DBGridInvProTer.Columns(2).Width = "1200"
                            DBGridInvProTer.Columns(3).Width = "3000"
                            DBGridInvProTer.Columns(4).Width = "1000"
                            DBGridInvProTer.Columns(5).Width = "1000"
                            DBGridInvProTer.Columns(6).Width = "1000"
                            
                            DBGridInvProTer.Columns(0).Caption = "Bodega"
                            DBGridInvProTer.Columns(1).Caption = "Descripcion"
                            DBGridInvProTer.Columns(2).Caption = "Ficha Tecnica"
                            DBGridInvProTer.Columns(3).Caption = "Descripcion"
                            DBGridInvProTer.Columns(4).Caption = "Tarimas/Bultos"
                            
                            DBGridInvProTer.Columns(5).Caption = "Unidades"
                            DBGridInvProTer.Columns(6).Caption = "Kilos"
                            
                            DBGridInvProTer.Columns(4).NumberFormat = "#,###,##0"
                            DBGridInvProTer.Columns(5).NumberFormat = "#,###,##0"
                            DBGridInvProTer.Columns(6).NumberFormat = "#,###,##0.00"
                            
                            DBGridInvProTer.Columns(4).Alignment = dbgRight
                            DBGridInvProTer.Columns(5).Alignment = dbgRight
                            DBGridInvProTer.Columns(6).Alignment = dbgRight

End Sub

Private Sub DataOrden_HeadClick(ByVal ColIndex As Integer)
        ROrdenesResumen.Sort = ROrdenesResumen.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_DblClick()
            If BLinea = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(0)
            ElseIf BGrupo = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(2)
            End If
            FrameBusqueda.Visible = False
            TxtLinea.SetFocus
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                If BLinea = True Then
                    TxtLinea.Text = DBGridBusqueda.Columns(0)
                ElseIf BGrupo = True Then
                    TxtLinea.Text = DBGridBusqueda.Columns(2)
                End If
                FrameBusqueda.Visible = False
                TxtLinea.SetFocus
            End If
End Sub

Private Sub DbGridGerencia_HeadClick(ByVal ColIndex As Integer)
        RProduccionxOrden.Sort = RProduccionxOrden.Fields(ColIndex).Name
End Sub

Private Sub DBGridInvProTer_HeadClick(ByVal ColIndex As Integer)
        RInventario.Sort = RInventario.Fields(ColIndex).Name
End Sub

Private Sub DbGridLineas_HeadClick(ByVal ColIndex As Integer)
        RProduccionxLinea.Sort = RProduccionxLinea.Fields(ColIndex).Name
End Sub

Private Sub DbGridMes_HeadClick(ByVal ColIndex As Integer)
        RProduccionxMes.Sort = RProduccionxMes.Fields(ColIndex).Name
End Sub

Private Sub DbGridOrden_HeadClick(ByVal ColIndex As Integer)
        RDetalleorden.Sort = RDetalleorden.Fields(ColIndex).Name
End Sub

Private Sub dbgridparos_HeadClick(ByVal ColIndex As Integer)
        RParos.Sort = RParos.Fields(ColIndex).Name
End Sub

Private Sub Form_Load()
            DtpFecIni.Value = Date
            DTPFecFin.Value = Date

End Sub

Private Sub OptGrupo_Click()
            LblDescripcion.Caption = "Grupo"
            TxtLinea.Visible = True
            TxtLinea.SetFocus
End Sub

Private Sub OptLinea_Click()
            LblDescripcion.Caption = "Linea"
            TxtLinea.Visible = True
            TxtLinea.SetFocus
End Sub

Private Sub OptTodos_Click()
            LblDescripcion.Caption = ""
            TxtLinea.Visible = False
End Sub

Private Sub TabGeneral_Click(PreviousTab As Integer)
    
        If TabGeneral.Tab = 1 Then
        
            
            '_______________________________________________________________________________________________________________________
            'EL GRID DE ORDEN RESUMEN
            Set ROrdenesResumen = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(ROrdenesResumen, "Select EO.Documento, EO.FichaTecnica, F.Descrip, EO.FechaApertura, EO.FechaEntrega, Sum(DO.Requerido), Sum(DO.Entregado), Sum(DO.Saldo) From EncabezadoOrdenProduccion EO, DetalleOrdenProduccion DO, FichaTecnica F Where EO.Documento = DO.Documento And EO.FichaTecnica = F.Esp_Tec And EO.Estado = 'ABIERTA' Group By EO.Documento, EO.FichaTecnica, F.Descrip, EO.FechaApertura, EO.FechaEntrega Order by EO.Documento")
                Else 'ORACLE
                    Call Abrir_Recordset(ROrdenesResumen, "Select EO.Documento, EO.FichaTecnica, F.Descrip, EO.FechaApertura, EO.FechaEntrega, Sum(DO.Requerido), Sum(DO.Entregado), Sum(DO.Saldo) From EncabezadoOrdenProduccion EO, DetalleOrdenProduccion DO, FichaTecnica F Where UPPER(EO.Documento) = UPPER(DO.Documento) And UPPER(EO.FichaTecnica) = UPPER(F.Esp_Tec) And UPPER(EO.Estado) = 'ABIERTA' Group By EO.Documento, EO.FichaTecnica, F.Descrip, EO.FechaApertura, EO.FechaEntrega Order by EO.Documento")
                End If
                    Set DataOrden.DataSource = ROrdenesResumen
        End If
                
                

End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas where Descrip Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
                
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas where Linea Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
                
            End If
                    
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtLinea_Change()
        If OptLinea.Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLinea.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                End If
                    If RBuscaLinea.RecordCount > 0 Then
                        LblLinea.Caption = RBuscaLinea!Descrip
                    Else
                        LblLinea.Caption = ""
                    End If
        End If
            
End Sub

Private Sub TxtLinea_DblClick()
            Set RBusqueda = New ADODB.Recordset
            If OptGrupo.Value = True Then
                    BGrupo = True
                    BLinea = False
            ElseIf OptLinea.Value = True Then
                    BGrupo = False
                    BLinea = True
            End If
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
End Sub

Private Sub TxtLinea_GotFocus()
        TxtLinea.SelStart = 0
        TxtLinea.SelLength = Len(TxtLinea.Text)
End Sub

Private Sub TxtLinea_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            Set RBusqueda = New ADODB.Recordset
            If OptGrupo.Value = True Then
                    BGrupo = True
                    BLinea = False
            ElseIf OptLinea.Value = True Then
                    BGrupo = False
                    BLinea = True
            End If
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If
End Sub
