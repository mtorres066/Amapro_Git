VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form BodegasInventario 
   BackColor       =   &H00008000&
   Caption         =   "Bodegas De Inventario"
   ClientHeight    =   4710
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   Icon            =   "BodegasInventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
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
      Height          =   4695
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   8412
      Begin MSDataGridLib.DataGrid DBGridBusqueda 
         Height          =   3495
         Left            =   120
         TabIndex        =   32
         Top             =   1080
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   6165
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
         Picture         =   "BodegasInventario.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   31
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
         TabIndex        =   30
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   7800
      MouseIcon       =   "BodegasInventario.frx":24B4
      Picture         =   "BodegasInventario.frx":28F6
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Ultimo Registro"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   7440
      MouseIcon       =   "BodegasInventario.frx":2E28
      Picture         =   "BodegasInventario.frx":326A
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Siguiente Registro"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "BodegasInventario.frx":379C
      Picture         =   "BodegasInventario.frx":3BDE
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Registro Anterior"
      Top             =   3960
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "BodegasInventario.frx":4110
      Picture         =   "BodegasInventario.frx":4552
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Primer Registro"
      Top             =   3960
      Width           =   375
   End
   Begin TabDlg.SSTab TabBodegas 
      Height          =   3732
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   8412
      _ExtentX        =   14843
      _ExtentY        =   6588
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "BodegasInventario.frx":4A84
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameBodegas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "BodegasInventario.frx":4D9E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGrid1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "BodegasInventario.frx":51F0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameOpciones"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   34
         Top             =   720
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   5106
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
         ColumnCount     =   9
         BeginProperty Column00 
            DataField       =   "CodigoBodega"
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
         BeginProperty Column03 
            DataField       =   "Telefono"
            Caption         =   "Telefono"
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
         BeginProperty Column05 
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
         BeginProperty Column06 
            DataField       =   "EsBodegaDeProceso"
            Caption         =   "EsBodegaDeProceso"
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
            DataField       =   "EsBodegaDeNoConforme"
            Caption         =   "EsBodegaDeNoConforme"
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
               ColumnWidth     =   1005.165
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3119.811
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
            EndProperty
            BeginProperty Column04 
            EndProperty
            BeginProperty Column05 
            EndProperty
            BeginProperty Column06 
            EndProperty
            BeginProperty Column07 
            EndProperty
            BeginProperty Column08 
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
         Height          =   2655
         Left            =   -74880
         TabIndex        =   21
         Top             =   840
         Width           =   8085
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Datos"
            Height          =   732
            Index           =   0
            Left            =   6120
            Picture         =   "BodegasInventario.frx":5642
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   960
            Width           =   1812
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Todos"
            Height          =   732
            Index           =   1
            Left            =   6120
            Picture         =   "BodegasInventario.frx":733C
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1800
            Width           =   1812
         End
         Begin VB.TextBox TxtBuscar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Left            =   4920
            TabIndex        =   39
            ToolTipText     =   " "
            Top             =   480
            Width           =   1845
         End
         Begin VB.OptionButton OptCodigo 
            Caption         =   "&Codigo"
            Height          =   225
            Left            =   120
            TabIndex        =   15
            ToolTipText     =   " "
            Top             =   300
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptNombre 
            Caption         =   "&Descripcion"
            Height          =   195
            Left            =   1320
            TabIndex        =   16
            ToolTipText     =   " "
            Top             =   300
            Width           =   1340
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
            Left            =   3600
            TabIndex        =   40
            Top             =   480
            Width           =   1215
         End
      End
      Begin VB.Frame FrameBodegas 
         Caption         =   "Datos de Bodega"
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
         Height          =   2892
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   8115
         Begin VB.CheckBox Check2 
            Caption         =   "Es Bodega No Conforme"
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
            Height          =   315
            Left            =   5400
            TabIndex        =   7
            Top             =   2520
            Width           =   2532
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Es Bodega De Proceso"
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
            Height          =   315
            Left            =   2880
            TabIndex        =   6
            Top             =   2520
            Width           =   2412
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   288
            Index           =   5
            Left            =   1080
            MaxLength       =   10
            TabIndex        =   5
            Top             =   2160
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            Height          =   285
            Index           =   6
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1692
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1800
            Width           =   6855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   3
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   3
            Top             =   1440
            Width           =   3495
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   2
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   2
            Top             =   1080
            Width           =   6855
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   0
            Left            =   1080
            MaxLength       =   3
            TabIndex        =   0
            ToolTipText     =   " "
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            Height          =   285
            Index           =   1
            Left            =   1080
            MaxLength       =   50
            TabIndex        =   1
            ToolTipText     =   " "
            Top             =   720
            Width           =   6855
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
            Height          =   252
            Left            =   2880
            TabIndex        =   27
            Top             =   2160
            Width           =   5052
         End
         Begin VB.Label Label2 
            Caption         =   "Grupo"
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   26
            Top             =   2160
            Width           =   732
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   192
            Index           =   4
            Left            =   120
            TabIndex        =   25
            Top             =   2520
            Width           =   540
         End
         Begin VB.Label Label2 
            Caption         =   "Encargado"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   24
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Telefono"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   23
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Direccion"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   22
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Codigo"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Descripcion"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   975
         End
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   6360
      MouseIcon       =   "BodegasInventario.frx":7646
      Picture         =   "BodegasInventario.frx":7A88
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   5280
      MouseIcon       =   "BodegasInventario.frx":9AFA
      Picture         =   "BodegasInventario.frx":9F3C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4200
      MouseIcon       =   "BodegasInventario.frx":A46E
      Picture         =   "BodegasInventario.frx":A8B0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3120
      MouseIcon       =   "BodegasInventario.frx":ADE2
      Picture         =   "BodegasInventario.frx":B224
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   2040
      MouseIcon       =   "BodegasInventario.frx":B756
      Picture         =   "BodegasInventario.frx":BB98
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3840
      Width           =   1000
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   960
      MouseIcon       =   "BodegasInventario.frx":C0CA
      Picture         =   "BodegasInventario.frx":C50C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   1000
   End
End
Attribute VB_Name = "BodegasInventario"
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

Dim RBodegasMateriaPrima As New ADODB.Recordset
Dim RBuscaGrupo As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Sub botones()
    If Bandera = True Then
         FrameBodegas.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         Txttexto.Item(0).SetFocus
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
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
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

Private Sub Check1_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub Check2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        Bandera = True
        botones
        Limpia_Campos
        
        Txttexto.Item(0).Enabled = True
        Txttexto.Item(0).SetFocus
        Txttexto.Item(6).Text = GUsuario
        BEditar = False
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RBodegasMateriaPrima.Delete
                        
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
                        RBodegasMateriaPrima.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RBodegasMateriaPrima.MoveLast
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
        RBodegasMateriaPrima.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RBodegasMateriaPrima.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RBodegasMateriaPrima.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RBodegasMateriaPrima.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RBodegasMateriaPrima.BOF Then
        RBodegasMateriaPrima.MoveFirst
    ElseIf RBodegasMateriaPrima.EOF Then
        RBodegasMateriaPrima.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    
    'INICIALIZAMOS EL RECORDSET
        Set RBodegasMateriaPrima = New ADODB.Recordset
        
    If Index = 0 Then
        If OptCodigo.Value = True Then
            Call Abrir_Recordset(RBodegasMateriaPrima, "Select * from BodegasInventario where CodigoBodega like '" & TxtBuscar.Text & "%'")
        ElseIf OptNombre.Value = True Then
            Call Abrir_Recordset(RBodegasMateriaPrima, "Select * from BodegasInventario where Descripcion like '" & TxtBuscar.Text & "%'")
        End If
    ElseIf Index = 1 Then
            Call Abrir_Recordset(RBodegasMateriaPrima, "Select * from BodegasInventario")
    End If
        Set DataGrid1.DataSource = RBodegasMateriaPrima
        TabBodegas.Tab = 1

End Sub

Private Sub CmdCancelar_Click()
            Bandera = False
            botones
            Llena_Campos
            Txttexto.Item(0).Enabled = True
End Sub

Private Sub CmdEditar_Click()

        Bandera = True
        botones
        Txttexto.Item(0).Enabled = False
        Txttexto.Item(1).SetFocus
        Txttexto.Item(6).Text = GUsuario
        BEditar = True
        
End Sub

Private Sub CmdGrabar_Click()
On Error Resume Next
                    Set RBuscaGrupo = New ADODB.Recordset
                        Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From BodegasInventarioGrupos Where Codigo = '" & Txttexto.Item(5).Text & "'")
                            If RBuscaGrupo.RecordCount > 0 Then
                            
                            Else
                                MsgBox "Grupo De Bodega No Existe", vbOKOnly + vbInformation, "Informacion"
                                Txttexto.Item(5).SetFocus
                                Exit Sub
                            End If


                    'AGREGAR
                    If BEditar = False Then
                            VTexto = "Values('" & Txttexto.Item(0).Text & "', '" ' CODIGO
                            VTexto = VTexto & Txttexto.Item(1).Text & "', '" 'DESCRIPCION
                            VTexto = VTexto & Txttexto.Item(2).Text & "', '" 'DIRECCION
                            VTexto = VTexto & Txttexto.Item(3).Text & "', '" 'TELEFONO
                            VTexto = VTexto & Txttexto.Item(4).Text & "', '" 'ENCARGADO
                            VTexto = VTexto & Txttexto.Item(5).Text & "', " 'GRUPO
                            If Check1.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'ES BODEGA DE PROCESO
                            Else
                                VTexto = VTexto & "0" & ", " 'ES BODEGA DE PROCESO
                            End If
                            If Check2.Value = "1" Then
                                VTexto = VTexto & "-1" & ", '" 'ES BODEGA DE NO CONFORME
                            Else
                                VTexto = VTexto & "0" & ", '" 'ES BODEGA DE NO CONFORME
                            End If
                            
                            VTexto = VTexto & Txttexto.Item(6).Text & "')" 'USUARIO
                            
                            Conexion.Execute "Insert Into BodegasInventario " & VTexto
                    'EDITAR
                    Else
                            VTexto = "Descripcion = '" & Txttexto.Item(1).Text & "', " 'DESCRIPCION
                            VTexto = VTexto & "Direccion = '" & Txttexto.Item(2).Text & "', " 'DIRECCION
                            VTexto = VTexto & "Telefono = '" & Txttexto.Item(3).Text & "', " 'TELEFONO
                            VTexto = VTexto & "Encargado = '" & Txttexto.Item(4).Text & "', " 'ENCARGADO
                            VTexto = VTexto & "Grupo = '" & Txttexto.Item(5).Text & "', " 'GRUPO
                            If Check1.Value = "1" Then
                                VTexto = VTexto & "EsBodegaDeProceso = -1" & ", " 'ES BODEGA DE PROCESO
                            Else
                                VTexto = VTexto & "EsBodegaDeProceso = 0" & ", " 'ES BODEGA DE PROCESO
                            End If
                            If Check2.Value = "1" Then
                                VTexto = VTexto & "EsBodegaDeNoConforme = -1" & ", " 'ES BODEGA DE NO CONFORME
                            Else
                                VTexto = VTexto & "EsBodegaDeNoConforme = 0" & ", " 'ES BODEGA DE NO CONFORME
                            End If
                            VTexto = VTexto & "usuario = '" & Txttexto.Item(6).Text & "' " ' USUARIO
                            VTexto = VTexto & "Where CodigoBodega = '" & Txttexto.Item(0).Text & "'"
                        
                            Conexion.Execute "UPDATE BodegasInventario SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            MsgBox "Codigo De Bodega Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            Txttexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Codigo De Bodega Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            Txttexto.Item(0).SetFocus
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
                        Txttexto.Item(0).Enabled = True
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RBodegasMateriaPrima.Requery
                        RBodegasMateriaPrima.MoveLast
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
                RBodegasMateriaPrima.Sort = RBodegasMateriaPrima.Fields(ColIndex).Name
            
            If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                Err.Clear
            End If

    
End Sub

Private Sub DBGridBusqueda_DblClick()
            Txttexto.Item(5).Text = DBGridBusqueda.Columns(0).Text
            Txttexto.Item(5).SetFocus
            FrameBusqueda.Visible = False

End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                Txttexto.Item(5).Text = DBGridBusqueda.Columns(0).Text
                Txttexto.Item(5).SetFocus
                FrameBusqueda.Visible = False
            End If
End Sub

Private Sub Form_Load()
        Set RBodegasMateriaPrima = New ADODB.Recordset
        Call Abrir_Recordset(RBodegasMateriaPrima, "Select * From BodegasInventario")
        Set DataGrid1.DataSource = RBodegasMateriaPrima
        Llena_Campos
    
        If GEditar = True Then
                DataGrid1.AllowUpdate = True
        Else
                DataGrid1.AllowUpdate = False
        End If

End Sub


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
        
        RBodegasMateriaPrima.Close
        RBuscaGrupo.Close
        RBusqueda.Close
        
        Set RBodegasMateriaPrima = Nothing
        Set RBuscaGrupo = Nothing
        Set RBusqueda = Nothing
        
        If Err <> 0 Then
        End If
        
End Sub

Private Sub OptCodigo_Click()
        Lbletiqueta.Caption = "Codigo"
End Sub


Private Sub OptNombre_Click()
        Lbletiqueta.Caption = "Descripcion"
End Sub



Private Sub TabBodegas_Click(PreviousTab As Integer)
    If TabBodegas.Tab = 0 Then
        If CmdGrabar.Enabled = False Then
            Llena_Campos
        End If
        CmdBorrar.Enabled = True
    Else
        CmdBorrar.Enabled = False
    End If

End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From BodegasInventarioGrupos where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From BodegasInventarioGrupos where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
                
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From BodegasInventarioGrupos where Codigo Like '%" & TxtBusqueda.Text & "%'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From BodegasInventarioGrupos where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
                    
                    Set DBGridBusqueda.DataSource = RBusqueda
                    
                    DBGridBusqueda.Columns(1).Width = "4000"

End Sub




Public Sub Llena_Campos()
On Error Resume Next
        
        Txttexto.Item(0).Text = RBodegasMateriaPrima!CodigoBodega
            If IsNull(RBodegasMateriaPrima!Descripcion) Then
                Txttexto.Item(1).Text = ""
            Else
                Txttexto.Item(1).Text = RBodegasMateriaPrima!Descripcion
            End If
        Txttexto.Item(2).Text = RBodegasMateriaPrima!Direccion
            If IsNull(RBodegasMateriaPrima!Direccion) Then
                Txttexto.Item(2).Text = ""
            Else
                Txttexto.Item(2).Text = RBodegasMateriaPrima!Direccion
            End If
        Txttexto.Item(3).Text = RBodegasMateriaPrima!Telefono
            If IsNull(RBodegasMateriaPrima!Telefono) Then
                Txttexto.Item(3).Text = ""
            Else
                Txttexto.Item(3).Text = RBodegasMateriaPrima!Telefono
            End If
        Txttexto.Item(4).Text = RBodegasMateriaPrima!Encargado
            If IsNull(RBodegasMateriaPrima!Encargado) Then
                Txttexto.Item(4).Text = ""
            Else
                Txttexto.Item(4).Text = RBodegasMateriaPrima!Encargado
            End If
        Txttexto.Item(5).Text = RBodegasMateriaPrima!Grupo
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RBodegasMateriaPrima!EsBodegaDeProceso = "Verdadero" Then
                    Check1.Value = "1"
                Else
                    Check1.Value = "0"
                End If
        Else
                If RBodegasMateriaPrima!EsBodegaDeProceso = "-1" Then
                    Check1.Value = "1"
                Else
                    Check1.Value = "0"
                End If
        End If
                
        If GOrigenDeDatos = "AmaproAccess" Then
                If RBodegasMateriaPrima!EsBodegadeNoConforme = "Verdadero" Then
                    Check2.Value = "1"
                Else
                    Check2.Value = "0"
                End If
        Else
                If RBodegasMateriaPrima!EsBodegadeNoConforme = "-1" Then
                    Check2.Value = "1"
                Else
                    Check2.Value = "0"
                End If
        End If
        
        Txttexto.Item(6).Text = RBodegasMateriaPrima!Usuario
        
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        Txttexto.Item(0).Text = ""
        Txttexto.Item(1).Text = ""
        Txttexto.Item(2).Text = ""
        Txttexto.Item(3).Text = ""
        Txttexto.Item(4).Text = ""
        Txttexto.Item(5).Text = ""
        Txttexto.Item(6).Text = ""
        Check1.Value = 0
        Check2.Value = 0
        
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
        If Index = 5 Then
            Set RBuscaGrupo = New ADODB.Recordset
            Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From BodegasInventarioGrupos Where Codigo = '" & Txttexto.Item(5).Text & "'")
                If RBuscaGrupo.RecordCount > 0 Then
                    LblGrupo.Caption = RBuscaGrupo!Descripcion
                Else
                    LblGrupo.Caption = ""
                End If
        
        End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
        If Index = 5 Then
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaGrupo = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaGrupo, "Select Codigo, Descripcion From BodegasInventarioGrupos")
            
            
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBusqueda.DataSource = RBuscaGrupo
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        End If
End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
            Txttexto.Item(Index).SelStart = 0
            Txttexto.Item(Index).SelLength = Len(Txttexto.Item(Index).Text)
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            If Index = 5 Then
                'INICIALIZAMOS EL RECORDSET
                Set RBuscaGrupo = New ADODB.Recordset
                'ABRIMOS EL RECORDSET
                Call Abrir_Recordset(RBuscaGrupo, "Select Codigo, Descripcion From BodegasInventarioGrupos")
            
            
                'LLENAMOS EL GRID CON EL RECORDSET
                Set DBGridBusqueda.DataSource = RBuscaGrupo
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
            End If
        End If
End Sub
