VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Proveedores 
   BackColor       =   &H00FF8080&
   Caption         =   "Proveedores"
   ClientHeight    =   7905
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "Proveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8415
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
      Height          =   7815
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   8412
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   6615
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   11668
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
         Left            =   7440
         Picture         =   "Proveedores.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   27
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4092
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   2160
      MouseIcon       =   "Proveedores.frx":293C
      Picture         =   "Proveedores.frx":2D7E
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7920
      MouseIcon       =   "Proveedores.frx":32B0
      Picture         =   "Proveedores.frx":36F2
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Ultimo Registro"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7560
      MouseIcon       =   "Proveedores.frx":3C24
      Picture         =   "Proveedores.frx":4066
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Siguiente Registro"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   600
      MouseIcon       =   "Proveedores.frx":4598
      Picture         =   "Proveedores.frx":49DA
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Registro Anterior"
      Top             =   7080
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   240
      MouseIcon       =   "Proveedores.frx":4F0C
      Picture         =   "Proveedores.frx":534E
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Primer Registro"
      Top             =   7080
      Width           =   375
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   6735
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   11880
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "Proveedores.frx":5880
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBodegas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "Proveedores.frx":5B9A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "Proveedores.frx":5FEC
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameOpciones"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5895
         Left            =   -74880
         TabIndex        =   30
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10398
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
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "CodigoProveedor"
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
         BeginProperty Column01 
            DataField       =   "Descripcion"
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
            DataField       =   "DiasCredito"
            Caption         =   "Dias Credito"
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
            DataField       =   "Direccion"
            Caption         =   "Direccion"
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
            DataField       =   "Telefono"
            Caption         =   "Telefono"
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
            DataField       =   "Fax"
            Caption         =   "Fax"
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
            DataField       =   "Nit"
            Caption         =   "Nit"
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
            DataField       =   "Encargado"
            Caption         =   "Encargado"
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
            DataField       =   "Grupo"
            Caption         =   "Grupo"
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
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3660.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   404.787
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   929.764
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   794.835
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   915.024
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   629.858
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   945.071
            EndProperty
         EndProperty
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
         Height          =   5655
         Left            =   -74880
         TabIndex        =   20
         Top             =   840
         Width           =   8085
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Datos"
            Height          =   732
            Index           =   0
            Left            =   6120
            Picture         =   "Proveedores.frx":643E
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   3120
            Width           =   1812
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Todos"
            Height          =   732
            Index           =   1
            Left            =   6120
            Picture         =   "Proveedores.frx":8138
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   3960
            Width           =   1812
         End
         Begin VB.TextBox TxtBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   6120
            TabIndex        =   35
            ToolTipText     =   " "
            Top             =   2640
            Width           =   1845
         End
         Begin VB.Label Lbletiqueta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Codigo Proveedor"
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
            Left            =   4485
            TabIndex        =   38
            Top             =   2640
            Width           =   1530
         End
      End
      Begin VB.Frame FrameBodegas 
         Caption         =   "Datos Del Proveedor"
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
         Height          =   4095
         Left            =   120
         TabIndex        =   17
         Top             =   1680
         Width           =   8115
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   9
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   8
            Top             =   3240
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   6
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   6
            Top             =   2520
            Width           =   5055
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   5
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2160
            Width           =   5055
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   4
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1800
            Width           =   5055
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   3
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   3
            Top             =   1440
            Width           =   5055
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   2
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1080
            Width           =   5055
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   7
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   7
            Top             =   2880
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   1
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            Top             =   720
            Width           =   5055
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   0
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   0
            Top             =   360
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Index           =   8
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1692
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Dias Credito"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   44
            Top             =   3240
            Width           =   855
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nit"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   43
            Top             =   2160
            Width           =   195
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Encargado"
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   42
            Top             =   2520
            Width           =   780
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Grupo"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   41
            Top             =   2880
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   40
            Top             =   3600
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   1800
            Width           =   255
         End
         Begin VB.Label LblGrupo 
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
            TabIndex        =   23
            Top             =   2880
            Width           =   5055
         End
         Begin VB.Label Label2 
            Caption         =   "Direccion"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Telefono"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   21
            Top             =   1440
            Width           =   630
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   18
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6480
      MouseIcon       =   "Proveedores.frx":8442
      Picture         =   "Proveedores.frx":8884
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6960
      Width           =   1000
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5400
      MouseIcon       =   "Proveedores.frx":A8F6
      Picture         =   "Proveedores.frx":AD38
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6960
      Width           =   1000
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4320
      MouseIcon       =   "Proveedores.frx":B26A
      Picture         =   "Proveedores.frx":B6AC
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6960
      Width           =   1000
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3240
      MouseIcon       =   "Proveedores.frx":BBDE
      Picture         =   "Proveedores.frx":C020
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6960
      Width           =   1000
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   1080
      MouseIcon       =   "Proveedores.frx":C552
      Picture         =   "Proveedores.frx":C994
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6960
      Width           =   1000
   End
End
Attribute VB_Name = "Proveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim BEditar As Boolean
Dim VTexto As String
Dim Vllave As String

Dim RProveedores As New ADODB.Recordset
Dim RBuscaGrupo As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Sub botones()
    If Bandera = True Then
         FrameBodegas.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtTexto.Item(0).SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False

         FrameOpciones.Visible = False
         DataGrid1.Visible = False
    Else
         FrameBodegas.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True

         FrameOpciones.Visible = True
         DataGrid1.Visible = True
    End If
End Sub



Private Sub CmdAgregar_Click()
On Error Resume Next
        Bandera = True
        botones
        Limpia_Campos
        TxtTexto.Item(0).Enabled = True
        TxtTexto.Item(0).SetFocus
        TxtTexto.Item(8).Text = GUsuario
        BEditar = False
        
        
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RProveedores.Delete
                        
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
                        RProveedores.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RProveedores.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If

End Sub



Private Sub CmdBotones2_Click(Index As Integer)
On Error Resume Next
MousePointer = 11
    If Index = 1 Then
        RProveedores.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RProveedores.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RProveedores.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RProveedores.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RProveedores.BOF Then
        RProveedores.MoveFirst
    ElseIf RProveedores.EOF Then
        RProveedores.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    
    'INICIALIZAMOS EL RECORDSET
        Set RProveedores = New ADODB.Recordset
        
    If Index = 0 Then
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RProveedores, "Select * from Proveedores where CodigoProveedor like '" & TxtBuscar.Text & "%'")
        Else
            Call Abrir_Recordset(RProveedores, "Select * from Proveedores where Upper(CodigoProveedor) like '" & UCase(TxtBuscar.Text) & "%'")
        End If
    ElseIf Index = 1 Then
            Call Abrir_Recordset(RProveedores, "Select * from Proveedores")
    End If
        Set DataGrid1.DataSource = RProveedores
        TabBodegas.Tab = 1

End Sub

Private Sub CmdCancelar_Click()
            Bandera = False
            botones
            Llena_Campos
            'HABILITA LA LLAVE
            TxtTexto.Item(0).Enabled = True
End Sub

Private Sub CmdEditar_Click()
                Bandera = True
                botones
                'DESABILITA LA LLAVE
                TxtTexto.Item(0).Enabled = False
                TxtTexto.Item(1).SetFocus
                TxtTexto.Item(8).Text = GUsuario
                BEditar = True

End Sub

Private Sub CmdGrabar_Click()
On Error Resume Next
                                                            
                    If Not IsNumeric(TxtTexto.Item(9).Text) Then
                        MsgBox "Dias De Credito Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        TxtTexto.Item(9).SetFocus
                        Exit Sub
                    End If
                    
                    If BEditar = False Then
                            VTexto = "'" & TxtTexto.Item(0).Text & "', " ' CODIGO
                            VTexto = VTexto & "'" & TxtTexto.Item(1).Text & "', " ' DESCRIPCION
                            VTexto = VTexto & "'" & TxtTexto.Item(7).Text & "', " ' USUARIO
                            VTexto = VTexto & "'" & TxtTexto.Item(3).Text & "', " ' USUARIO
                            VTexto = VTexto & "'" & TxtTexto.Item(4).Text & "', " ' USUARIO
                            VTexto = VTexto & "'" & TxtTexto.Item(5).Text & "', " ' USUARIO
                            VTexto = VTexto & "'" & TxtTexto.Item(6).Text & "', " ' USUARIO
                            VTexto = VTexto & "'" & TxtTexto.Item(7).Text & "', " ' USUARIO
                            VTexto = VTexto & "'" & TxtTexto.Item(8).Text & "', " '
                            VTexto = VTexto & TxtTexto.Item(9).Text 'DIAS CREDITO
                            
                            Conexion.Execute "Insert Into Proveedores Values(" & VTexto & ")"
                    Else
                            'VTexto = "'" & TxtTexto.Item(0).Text & "', " ' CODIGO
                            VTexto = "Descripcion = '" & TxtTexto.Item(1).Text & "', " ' DESCRIPCION
                            VTexto = VTexto & "Direccion = '" & TxtTexto.Item(7).Text & "', " ' DIRECCION
                            VTexto = VTexto & "Telefono = '" & TxtTexto.Item(3).Text & "', " ' TELEFONO
                            VTexto = VTexto & "Fax = '" & TxtTexto.Item(4).Text & "', " ' FAX
                            VTexto = VTexto & "Nit = '" & TxtTexto.Item(5).Text & "', " ' NIT
                            VTexto = VTexto & "Encargado = '" & TxtTexto.Item(6).Text & "', " ' ENCARGADO
                            VTexto = VTexto & "Grupo = '" & TxtTexto.Item(7).Text & "', " ' GRUPO
                            VTexto = VTexto & "Usuario = '" & TxtTexto.Item(8).Text & "', " ' USUARIO
                            VTexto = VTexto & "DiasCredito = " & TxtTexto.Item(9).Text 'DIAS CREDITO
                            
                            VTexto = VTexto & " Where CodigoProveedor = '" & TxtTexto.Item(0).Text & "'" 'LLAVE
                            
                            Conexion.Execute "Update Proveedores Set " & VTexto
                    End If
                    
                   'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Codigo Proveedor Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo Proveedor Ya Existe", vbOKOnly + vbInformation, "Informacion"
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
                        CmdAgregar.SetFocus
                        'HABILITA LA LLAVE
                        TxtTexto.Item(0).Enabled = True
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RProveedores.Requery
                        RProveedores.MoveLast
                        Llena_Campos
      

End Sub

Private Sub CmdSale_Click()
        FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DataGrid1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
                RProveedores.Sort = RProveedores.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If

    
End Sub

Private Sub DBGridBusqueda_DblClick()
            TxtTexto.Item(7).Text = DBGridBusqueda.Columns(0).Text
            TxtTexto.Item(7).SetFocus
            FrameBusqueda.Visible = False

End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                TxtTexto.Item(7).Text = DBGridBusqueda.Columns(0).Text
                TxtTexto.Item(7).SetFocus
                FrameBusqueda.Visible = False
            End If
End Sub

Private Sub Form_Load()
        Set RProveedores = New ADODB.Recordset
        Call Abrir_Recordset(RProveedores, "Select * From Proveedores")
        Set DataGrid1.DataSource = RProveedores
        Llena_Campos
    
        'If GEditar = True Then
        '        DataGrid1.AllowUpdate = True
        'Else
        '        DataGrid1.AllowUpdate = False
        'End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
        RProveedores.Close
        RBuscaGrupo.Close
        RBusqueda.Close
        
        Set RProveedores = Nothing
        Set RBuscaGrupo = Nothing
        Set RBusqueda = Nothing
        
        If Err <> 0 Then
        End If
        
End Sub



Private Sub TabBodegas_Click(PreviousTab As Integer)
    If TabBodegas.Tab = 0 Then
        CmdBorrar.Enabled = True
            If CmdGrabar.Enabled = False Then
                Llena_Campos
            End If
    Else
        CmdBorrar.Enabled = False
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
            Set RProveedores = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProveedores, "Select CodigoProveedor, Descripcion From Proveedores where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProveedores, "Select CodigoProveedor, Descripcion From Proveedores where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProveedores, "Select CodigoProveedor, Descripcion From Proveedores where CodigoProveedor Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProveedores, "Select CodigoProveedor, Descripcion From Proveedores where UPPER(CodigoProveedor) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
                    
                    Set DBGridBusqueda.DataSource = RProveedores
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub




Public Sub Llena_Campos()
On Error Resume Next

    If RProveedores.RecordCount > 0 Then
        
                'NUMERO DOCUMEENTO
                    If IsNull(RProveedores!CodigoProveedor) Then
                        TxtTexto.Item(0).Text = ""
                    Else
                        TxtTexto.Item(0).Text = RProveedores!CodigoProveedor
                    End If
                'TIPO DE DOCUMENTO
                    If IsNull(RProveedores!Descripcion) Then
                        TxtTexto.Item(1).Text = ""
                    Else
                        TxtTexto.Item(1).Text = RProveedores!Descripcion
                    End If
                'PEDIDO
                    If IsNull(RProveedores!Direccion) Then
                        TxtTexto.Item(7).Text = ""
                    Else
                        TxtTexto.Item(7).Text = RProveedores!Direccion
                    End If
                'FICHA TECNICA
                    If IsNull(RProveedores!Telefono) Then
                        TxtTexto.Item(3).Text = ""
                    Else
                        TxtTexto.Item(3).Text = RProveedores!Telefono
                    End If
                    
                'NUMERO DOCUMEENTO
                    If IsNull(RProveedores!Fax) Then
                        TxtTexto.Item(4).Text = ""
                    Else
                        TxtTexto.Item(4).Text = RProveedores!Fax
                    End If
                'TIPO DE DOCUMENTO
                    If IsNull(RProveedores!Nit) Then
                        TxtTexto.Item(5).Text = ""
                    Else
                        TxtTexto.Item(5).Text = RProveedores!Nit
                    End If
                'PEDIDO
                    If IsNull(RProveedores!Encargado) Then
                        TxtTexto.Item(6).Text = ""
                    Else
                        TxtTexto.Item(6).Text = RProveedores!Encargado
                    End If
                'FICHA TECNICA
                    If IsNull(RProveedores!Grupo) Then
                        TxtTexto.Item(7).Text = ""
                    Else
                        TxtTexto.Item(7).Text = RProveedores!Grupo
                    End If
                'USUARIO
                    If IsNull(RProveedores!Usuario) Then
                        TxtTexto.Item(8).Text = ""
                    Else
                        TxtTexto.Item(8).Text = RProveedores!Usuario
                    End If
                'DIAS CREDITO
                    If IsNull(RProveedores!DiasCredito) Then
                        TxtTexto.Item(9).Text = ""
                    Else
                        TxtTexto.Item(9).Text = RProveedores!DiasCredito
                    End If
    Else
                    Limpia_Campos
    
    End If
                    
    
        If Err <> 0 Then
            'MsgBox Err.Description
        End If
End Sub

Public Sub Limpia_Campos()
        TxtTexto.Item(0).Text = ""
        TxtTexto.Item(1).Text = ""
        TxtTexto.Item(7).Text = ""
        TxtTexto.Item(3).Text = ""
        TxtTexto.Item(4).Text = ""
        TxtTexto.Item(5).Text = ""
        TxtTexto.Item(6).Text = ""
        TxtTexto.Item(7).Text = ""
        TxtTexto.Item(8).Text = ""
        TxtTexto.Item(9).Text = "0"
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
        If Index = 7 Then
            Set RBuscaGrupo = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From ProveedoresGrupos Where Codigo = '" & TxtTexto.Item(7).Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From ProveedoresGrupos Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(7).Text) & "'")
                End If
                If RBuscaGrupo.RecordCount > 0 Then
                    LblGrupo.Caption = RBuscaGrupo!Descripcion
                Else
                    LblGrupo.Caption = ""
                End If
        
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        If Index = 7 Then
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaGrupo = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaGrupo, "Select Codigo, Descripcion From ProveedoresGrupos")
            
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBusqueda.DataSource = RBuscaGrupo
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
            If Index = 7 Then
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaGrupo = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaGrupo, "Select Codigo, Descripcion From ProveedoresGrupos")
            
            
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBusqueda.DataSource = RBuscaGrupo
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
            End If
        End If
End Sub
