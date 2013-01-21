VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CapturaProduccionLiberada 
   BackColor       =   &H000000FF&
   Caption         =   "Captura Produccion Liberada"
   ClientHeight    =   7995
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11775
   Icon            =   "CapturaProduccionLiberada.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   7815
      Left            =   0
      TabIndex        =   39
      Top             =   120
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid Dbgridconsultas 
         Height          =   6495
         Left            =   120
         TabIndex        =   43
         Top             =   1080
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11456
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
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   42
         Top             =   720
         Width           =   3975
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1680
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   10680
         Picture         =   "CapturaProduccionLiberada.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   240
         Width           =   735
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
         Left            =   240
         TabIndex        =   63
         Top             =   720
         Width           =   1455
      End
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   585
      Index           =   1
      Left            =   120
      MouseIcon       =   "CapturaProduccionLiberada.frx":237C
      Picture         =   "CapturaProduccionLiberada.frx":27BE
      Style           =   1  'Graphical
      TabIndex        =   84
      ToolTipText     =   "Primer Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   585
      Index           =   2
      Left            =   480
      MouseIcon       =   "CapturaProduccionLiberada.frx":2CF0
      Picture         =   "CapturaProduccionLiberada.frx":3132
      Style           =   1  'Graphical
      TabIndex        =   83
      ToolTipText     =   "Registro Anterior"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   585
      Index           =   3
      Left            =   10920
      MouseIcon       =   "CapturaProduccionLiberada.frx":3664
      Picture         =   "CapturaProduccionLiberada.frx":3AA6
      Style           =   1  'Graphical
      TabIndex        =   82
      ToolTipText     =   "Siguiente Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   585
      Index           =   4
      Left            =   11280
      MouseIcon       =   "CapturaProduccionLiberada.frx":3FD8
      Picture         =   "CapturaProduccionLiberada.frx":441A
      Style           =   1  'Graphical
      TabIndex        =   81
      ToolTipText     =   "Ultimo Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   1212
      Left            =   3000
      ScaleHeight     =   79
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   359
      TabIndex        =   79
      Top             =   3840
      Visible         =   0   'False
      Width           =   5415
   End
   Begin VB.CommandButton CmdBotones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Etiquetas Caja"
      Height          =   705
      Index           =   9
      Left            =   7920
      MouseIcon       =   "CapturaProduccionLiberada.frx":494C
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton CmdBotones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Boleta No Conforme"
      Height          =   705
      Index           =   7
      Left            =   6720
      MouseIcon       =   "CapturaProduccionLiberada.frx":4D8E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "Tarimas Revisadas"
      Height          =   705
      Index           =   6
      Left            =   9120
      MouseIcon       =   "CapturaProduccionLiberada.frx":51D0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7200
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Im&primir Boleta"
      Height          =   705
      Index           =   5
      Left            =   5520
      MouseIcon       =   "CapturaProduccionLiberada.frx":5612
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   705
      Index           =   8
      Left            =   10080
      MouseIcon       =   "CapturaProduccionLiberada.frx":5A54
      Picture         =   "CapturaProduccionLiberada.frx":5E96
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   7200
      Width           =   800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   705
      Index           =   4
      Left            =   4680
      MouseIcon       =   "CapturaProduccionLiberada.frx":63B1
      Picture         =   "CapturaProduccionLiberada.frx":67F3
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7200
      Width           =   800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   705
      Index           =   3
      Left            =   3840
      MouseIcon       =   "CapturaProduccionLiberada.frx":6DBB
      Picture         =   "CapturaProduccionLiberada.frx":71FD
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7200
      Width           =   800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   705
      Index           =   2
      Left            =   2880
      MouseIcon       =   "CapturaProduccionLiberada.frx":7734
      Picture         =   "CapturaProduccionLiberada.frx":7B76
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7200
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   705
      Index           =   1
      Left            =   1920
      MouseIcon       =   "CapturaProduccionLiberada.frx":80D2
      Picture         =   "CapturaProduccionLiberada.frx":8514
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7200
      Width           =   915
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   705
      Index           =   0
      Left            =   960
      MouseIcon       =   "CapturaProduccionLiberada.frx":88EB
      Picture         =   "CapturaProduccionLiberada.frx":8D2D
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7200
      Width           =   915
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   255
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CapturaProduccionLiberada.frx":90AA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameProduccion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TxtDatos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Vista General "
      TabPicture(1)   =   "CapturaProduccionLiberada.frx":93C4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridProduccion"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda O Seleccion De Datos"
      TabPicture(2)   =   "CapturaProduccionLiberada.frx":9816
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LblEtiqueta(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "LblEtiqueta(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "LblEtiqueta(3)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "LblBuscar2"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "FrameBuscar"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "CmdActualizar"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "CmdBuscar"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "TxtBuscar"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "DtpFecIni"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "DtpFecFin"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "TxtBuscar2"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin MSDataGridLib.DataGrid DbGridProduccion 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   80
         ToolTipText     =   "para seleccionar haga click de el lado izquiero de la fila"
         Top             =   720
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   10821
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
         ColumnCount     =   15
         BeginProperty Column00 
            DataField       =   "Fec_Prd"
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
            DataField       =   "Hor_Prd"
            Caption         =   "Hora"
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
            DataField       =   "Esp_Tec"
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
            DataField       =   "Envases"
            Caption         =   "Envases"
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
         BeginProperty Column08 
            DataField       =   "Turno"
            Caption         =   "Turno"
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
            DataField       =   "Cod_Emp"
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
         BeginProperty Column10 
            DataField       =   "NoMP9301"
            Caption         =   "No.Mp9301"
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
            DataField       =   "ColorMp9301"
            Caption         =   "Color MP9301"
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
         BeginProperty Column13 
            DataField       =   "Orden"
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
         BeginProperty Column14 
            DataField       =   "Barra"
            Caption         =   "Barra"
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
               ColumnWidth     =   1094.74
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   480.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   329.953
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   659.906
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   315.213
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   329.953
            EndProperty
            BeginProperty Column09 
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column10 
               ColumnWidth     =   734.74
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   870.236
            EndProperty
            BeginProperty Column12 
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column13 
               ColumnWidth     =   1154.835
            EndProperty
            BeginProperty Column14 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox TxtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3420
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   67
         Top             =   2040
         Width           =   8415
      End
      Begin VB.TextBox TxtBuscar2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -65520
         TabIndex        =   36
         Top             =   4440
         Visible         =   0   'False
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   375
         Left            =   -65520
         TabIndex        =   34
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37213
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   375
         Left            =   -68280
         TabIndex        =   33
         Top             =   3000
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   52035587
         CurrentDate     =   37213
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -65520
         TabIndex        =   35
         ToolTipText     =   " "
         Top             =   3960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Datos"
         Height          =   855
         Left            =   -66720
         Picture         =   "CapturaProduccionLiberada.frx":9C68
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   4920
         Width           =   3015
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Seleccionar Todos Datos"
         Height          =   825
         Left            =   -66720
         Picture         =   "CapturaProduccionLiberada.frx":A0AA
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   5880
         Width           =   3045
      End
      Begin VB.Frame FrameBuscar 
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
         TabIndex        =   26
         Top             =   840
         Width           =   3255
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Ficha Tecnica"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000006&
            Height          =   285
            Index           =   2
            Left            =   120
            Picture         =   "CapturaProduccionLiberada.frx":A3B4
            TabIndex        =   29
            Top             =   1080
            Width           =   2475
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Orden"
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
            Index           =   3
            Left            =   120
            Picture         =   "CapturaProduccionLiberada.frx":AC7E
            TabIndex        =   30
            Top             =   1440
            Width           =   915
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "# Identificacion y Color"
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
            Index           =   5
            Left            =   120
            Picture         =   "CapturaProduccionLiberada.frx":AF88
            TabIndex        =   32
            Top             =   2160
            Width           =   2475
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas"
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
            Index           =   0
            Left            =   120
            Picture         =   "CapturaProduccionLiberada.frx":DA82
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   1035
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Fechas Y Linea"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   285
            Index           =   1
            Left            =   120
            Picture         =   "CapturaProduccionLiberada.frx":DD8C
            TabIndex        =   28
            Top             =   720
            Width           =   2355
         End
         Begin VB.OptionButton OptOpcion 
            Caption         =   "Batch Y Linea"
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
            Height          =   285
            Index           =   4
            Left            =   120
            Picture         =   "CapturaProduccionLiberada.frx":E096
            TabIndex        =   31
            Top             =   1800
            Width           =   1635
         End
      End
      Begin VB.Frame FrameProduccion 
         Caption         =   "Datos Generales Captura de Produccion"
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
         Height          =   6135
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   11295
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   4
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   35
            TabIndex        =   77
            TabStop         =   0   'False
            Top             =   5760
            Width           =   4515
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   3
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   8
            ToolTipText     =   "agrupacion de 16 tarimas"
            Top             =   2400
            Width           =   1395
         End
         Begin VB.TextBox TxtObservaciones 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   15
            ToolTipText     =   "Maximo 50 Digitos"
            Top             =   5400
            Width           =   9855
         End
         Begin VB.TextBox TxtTur 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   10
            Top             =   3120
            Width           =   435
         End
         Begin MSMask.MaskEdBox MskHor 
            Height          =   285
            Left            =   1320
            TabIndex        =   5
            Top             =   1320
            Width           =   1395
            _ExtentX        =   2461
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   25
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   11
            ToolTipText     =   "No. Hoja De Identificacion"
            Top             =   3480
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
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
            Index           =   16
            Left            =   1320
            MaxLength       =   20
            TabIndex        =   13
            Top             =   4200
            Width           =   1335
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            Height          =   285
            Index           =   8
            Left            =   1320
            TabIndex        =   9
            Top             =   2760
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Index           =   7
            Left            =   1320
            TabIndex        =   7
            ToolTipText     =   "agrupacion de 16 tarimas"
            Top             =   2040
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
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
            Index           =   6
            Left            =   1320
            TabIndex        =   6
            Top             =   1680
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   2
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   4
            Top             =   960
            Width           =   1395
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   1
            Left            =   1320
            MaxLength       =   15
            TabIndex        =   3
            Top             =   600
            Width           =   1400
         End
         Begin VB.TextBox TxtTexto 
            Appearance      =   0  'Flat
            BackColor       =   &H008080FF&
            Height          =   285
            Index           =   0
            Left            =   1320
            MaxLength       =   2
            TabIndex        =   2
            Top             =   240
            Width           =   1400
         End
         Begin VB.ComboBox CboColor 
            Appearance      =   0  'Flat
            Height          =   315
            ItemData        =   "CapturaProduccionLiberada.frx":E4D8
            Left            =   1320
            List            =   "CapturaProduccionLiberada.frx":E4F1
            TabIndex        =   12
            Text            =   "BLANCA"
            Top             =   3840
            Width           =   1335
         End
         Begin VB.ComboBox CboCal 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "CapturaProduccionLiberada.frx":E52E
            Left            =   1320
            List            =   "CapturaProduccionLiberada.frx":E53E
            TabIndex        =   14
            Text            =   "A"
            ToolTipText     =   "Calidad De Tarima"
            Top             =   4560
            Width           =   615
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Barra"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   78
            Top             =   5760
            Width           =   1005
         End
         Begin VB.Label LblEmpaque 
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
            Left            =   9240
            TabIndex        =   76
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label LblUltimoBatch2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   5400
            TabIndex        =   75
            Top             =   4920
            Width           =   1215
         End
         Begin VB.Label LblBatch2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
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
            Height          =   375
            Left            =   4080
            TabIndex        =   74
            Top             =   4920
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Ultimo Batch"
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
            Height          =   435
            Index           =   6
            Left            =   4800
            TabIndex        =   73
            Top             =   4920
            Width           =   615
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tarimas Batch Produccion"
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
            Height          =   435
            Index           =   5
            Left            =   2760
            TabIndex        =   72
            Top             =   4920
            Width           =   1320
         End
         Begin VB.Label lblLabels 
            Caption         =   "Ultimo Batch Produccion Lib"
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
            Height          =   435
            Index           =   1
            Left            =   8760
            TabIndex        =   71
            Top             =   4920
            Width           =   1335
         End
         Begin VB.Label lblLabels 
            Caption         =   "Tarimas Batch Produccion Lib."
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
            Height          =   435
            Index           =   4
            Left            =   6720
            TabIndex        =   70
            Top             =   4920
            Width           =   1395
         End
         Begin VB.Label lblLabels 
            Caption         =   "Orden"
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
            TabIndex        =   69
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label LblUltimoBatch 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   10200
            TabIndex        =   68
            Top             =   4920
            Width           =   975
         End
         Begin VB.Label lblLabels 
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"CapturaProduccionLiberada.frx":E54E
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   2760
            TabIndex        =   66
            Top             =   960
            Width           =   8415
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   3
            Left            =   120
            TabIndex        =   65
            Top             =   5400
            Width           =   1140
         End
         Begin VB.Label LblFichaTecnica 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   60
            Top             =   600
            Width           =   6375
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
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   2760
            TabIndex        =   59
            Top             =   240
            Width           =   8415
         End
         Begin VB.Label lblLabels 
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
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   57
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblLabels 
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
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   19
            Left            =   120
            TabIndex        =   56
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lblLabels 
            Caption         =   "Hora"
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   20
            Left            =   120
            TabIndex        =   55
            Top             =   1320
            Width           =   855
         End
         Begin VB.Label lblLabels 
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
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   21
            Left            =   120
            TabIndex        =   54
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label lblLabels 
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
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   53
            Top             =   1680
            Width           =   615
         End
         Begin VB.Label lblLabels 
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
            ForeColor       =   &H00FF0000&
            Height          =   255
            Index           =   29
            Left            =   120
            TabIndex        =   52
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label lblLabels 
            Caption         =   "Cantidad"
            Height          =   255
            Index           =   30
            Left            =   120
            TabIndex        =   51
            Top             =   2760
            Width           =   855
         End
         Begin VB.Label lblLabels 
            Caption         =   "Calidad"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   31
            Left            =   120
            TabIndex        =   50
            Top             =   4560
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "Turno"
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   49
            Top             =   3120
            Width           =   975
         End
         Begin VB.Label lblLabels 
            Caption         =   "# Identificacion"
            Height          =   255
            Index           =   7
            Left            =   120
            TabIndex        =   48
            Top             =   3480
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Color"
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   47
            Top             =   3840
            Width           =   1095
         End
         Begin VB.Label lblLabels 
            Caption         =   "Usuario"
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   46
            Top             =   4200
            Width           =   855
         End
         Begin VB.Label LblBatch 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   8160
            TabIndex        =   45
            Top             =   4920
            Width           =   495
         End
      End
      Begin VB.Label LblBuscar2 
         Alignment       =   1  'Right Justify
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
         Left            =   -66120
         TabIndex        =   64
         Top             =   4440
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label LblEtiqueta 
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   -66240
         TabIndex        =   62
         Top             =   3120
         Width           =   510
      End
      Begin VB.Label LblEtiqueta 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   -69240
         TabIndex        =   61
         Top             =   3120
         Width           =   795
      End
      Begin VB.Label LblEtiqueta 
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
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   -69000
         TabIndex        =   58
         Top             =   3960
         Width           =   3375
      End
   End
End
Attribute VB_Name = "CapturaProduccionLiberada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim VTipo As String
Dim VUltimaFecha As Date
Dim VFechaActual As Date
Dim VUnidadesxCaja As Integer
Dim Cont As Integer

Dim RLineas As New ADODB.Recordset
Dim RBuscaUnidadesxCaja As New ADODB.Recordset
Dim RBuscaProduccion As New ADODB.Recordset
Dim RBuscaEnvases As New ADODB.Recordset
Dim RReporteIdentificacionInterno As New ADODB.Recordset
Dim RBuscaUltimaFicha As New ADODB.Recordset
Dim RBuscaObservaciones As New ADODB.Recordset
Dim RBuscaUltimoBatch As New ADODB.Recordset
Dim RBuscaUltimoBatch2 As New ADODB.Recordset
Dim RBuscaOrden As New ADODB.Recordset
Dim RProduccion As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaFichaTecnicaConMateriaPrima As New ADODB.Recordset
Dim RCuentaFichaTecnicaConMateriaPrima As New ADODB.Recordset
Dim RBuscaAtributo As New ADODB.Recordset

Dim VLineas As Boolean
Dim BActivar As Boolean
Dim BVer As Boolean

Dim VDia As String
Dim VMes As String
Dim VAño As String

Dim VSumaDefectos As Integer
Dim VSumaDefectos2 As Integer

Dim RCuentaTarimas As New ADODB.Recordset
Dim RCuentaTarimas2 As New ADODB.Recordset
Dim BEditar As Boolean
Dim RVerificaTarima As New ADODB.Recordset
Dim BuscaTarimas As New ADODB.Recordset

Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaDefectos As New ADODB.Recordset

Dim MinWidth As Long
Dim pw As Long
Dim fw As Long

Dim VCampos As String
Dim VValores As String
Dim VUpdate As String



Sub botones()
    If Bandera = True Then
         FrameProduccion.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         CmdBotones.Item(6).Enabled = False
         CmdBotones.Item(8).Enabled = False
         CmdBotones.Item(9).Enabled = False
         CmdBotones.Item(10).Enabled = False
         'LblBatch.Visible = True
         TxtTexto.Item(1).SetFocus
         FrameBuscar.Visible = False
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         DBGridProduccion.Visible = False
    Else
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         CmdBotones.Item(6).Enabled = True
         CmdBotones.Item(8).Enabled = True
         CmdBotones.Item(9).Enabled = True
         CmdBotones.Item(10).Enabled = True
         FrameProduccion.Enabled = False
         'LblBatch.Visible = False
         FrameBuscar.Visible = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         DBGridProduccion.Visible = True
    End If
End Sub

Private Sub CboCal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub
Private Sub CboColor_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub


Private Sub CmdActualizar_Click()
'    VPLFecha = Date - 1
 '   DataProduccion.RecordSource = ("Select * from ProduccionLiberada Where Fec_Prd >= #" & Format(VPLFecha, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(Date, "mm/dd/yyyy") & "# Order BY Fec_prd")
    Set RProduccion = New ADODB.Recordset
    Call Abrir_Recordset(RProduccion, "Select * From ProduccionLiberada Order By Fec_Prd, Hor_Prd")
    Set DBGridProduccion.DataSource = RProduccion
    SSTab1.Tab = 1
End Sub

Private Sub CmdBotones_Click(Index As Integer)
On Error Resume Next
    'AGREGAR
    If Index = 0 Then
                    Bandera = True
                    botones
                    'PONE EN BLANCO TODOS LOS CAMPOS
                    Limpia_Campos
                    TxtDatos.Text = ""
                    TxtTexto.Item(0).SetFocus
                    
                    
                    'SI LA HORA ES MENOR QUE LAS 7 DE LA MAÑANA ENTONCES DA LA FECHA ANTERIOR
                    If Format(Time, "hh:mm") < "07:00" Then
                       TxtTexto.Item(2).Text = Format(DateValue(Date) - 1, "dd/mm/yyyy")
                    Else
                        TxtTexto.Item(2).Text = Format(Date, "dd/mm/yyyy")
                    End If
                    If Format(Time, "hh:mm") >= "07:00" And Format(Time, "hh:mm") <= "19:00" Then
                        TxtTur.Text = "1"
                    Else
                        TxtTur.Text = "2"
                    End If
                    
                    MskHor.Text = Time
                    CboCal.Text = "A"
                    BEditar = False
                    BVer = True
                    TxtTexto.Item(0).Enabled = True
                    TxtTexto.Item(1).Enabled = True
                    TxtTexto.Item(2).Enabled = True
                    TxtTexto.Item(6).Enabled = True
                    
                    'PONE EN BLANCO LA BARRA
                    TxtTexto.Item(4).Text = ""
                    
    'EDITAR
    ElseIf Index = 1 Then
                    If RProduccion.RecordCount > 0 Then
                            MousePointer = 0
                            'ASIGNAMOS A LA VARIABLE FECHA DEL SISTEMA MENOS 1
                            VUltimaFecha = DateValue(Date) - 1
                            VFechaActual = DateValue(Date)
                            
                            'SI PUEDE EDITAR NO VALIDA LAS FECHAS
                            If GEditar = True Then
                            Else
                                    If (DateValue(TxtTexto.Item(2).Text) >= VUltimaFecha And DateValue(TxtTexto.Item(2).Text) <= VFechaActual) Then
                                    Else
                                        MsgBox "No Puede EDITAR Produccion De 2 o mas dias de la fecha actual, Llame al Encargado", vbOKOnly + vbInformation, "Informacion"
                                        Exit Sub
                                    End If
                            End If
            
                            Bandera = True
                            botones
                            TxtTexto.Item(1).SetFocus
                            BEditar = True
                            TxtTexto.Item(0).Enabled = False
                            TxtTexto.Item(1).Enabled = False
                            TxtTexto.Item(2).Enabled = False
                            TxtTexto.Item(6).Enabled = False
                    Else
                            MsgBox "No Hay Registros Activos", vbOKOnly + vbInformation, "Informacion"
                    End If
                    
    'GRABAR
    ElseIf Index = 2 Then
                        
                   If GOrigenDeDatos = "AmaproAccess" Then
                   Else
                        TxtTexto.Item(2).Text = Format(TxtTexto.Item(2).Text, "dd/mm/yyyy")
                   End If
                   
                   'VALIDA LA FECHA
                   If Not IsDate(TxtTexto.Item(2).Text) Then
                        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                   End If
                   
                   'REVISA LA CANTIDAD
                   If Not IsNumeric(TxtTexto.Item(8).Text) Then
                        MsgBox "Cantidad De Unidades Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        MousePointer = 0
                        Exit Sub
                   End If
                                      
                   'REVISA EL TIPO DE CALIDAD
                   'If CboCal.Text <> "A" Then
                   '     MsgBox "Calidad Incorrecta, Solo Puede Ser 'A'", vbOKOnly + vbInformation, "Informacion"
                   '     MousePointer = 0
                   '     Exit Sub
                   'End If
                
                   'REVISA EL BATCH
                   If Not IsNumeric(TxtTexto.Item(7).Text) Then
                        MsgBox "Numero de Batch Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        MousePointer = 0
                        Exit Sub
                   End If
                                      
                   'REVISA LA TARIMA
                   If Not IsNumeric(TxtTexto.Item(6).Text) Then
                        MsgBox "Numero De Tarima Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                        MousePointer = 0
                        Exit Sub
                   End If
                   
                   'REVISA EL TAMAÑO DE LOS DIGITOS DEL COLOR
                   If Len(CboColor.Text) > 10 Then
                        MsgBox "El Color No Puede Ser Mayor De 10 Digitos", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                   
                   VPLLinea = TxtTexto.Item(0).Text
                   VPLFicha = TxtTexto.Item(1).Text
                   VPLFecha = TxtTexto.Item(2).Text
                   VPLTarima = TxtTexto.Item(6).Text
                                           
                   'SI NO ESTA EDITANDO SOLO GRABANDO
                   If BEditar = False Then
                            'VERIFICA SI YA EXISTE LA TARIMA EN PRODUCCION LIBERADA
                            Set RVerificaTarima = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RVerificaTarima, "Select * from ProduccionLiberada Where Linea = '" & VPLLinea & "' and Esp_tec = '" & VPLFicha & "' and Fec_prd = #" & Format(VPLFecha, "mm/dd/yyyy") & "# and Tarima = " & VPLTarima)
                                Else 'ORACLE
                                    Call Abrir_Recordset(RVerificaTarima, "Select * from ProduccionLiberada Where UPPER(Linea) = '" & UCase(VPLLinea) & "' and UPPER(Esp_tec) = '" & UCase(VPLFicha) & "' and Fec_prd = To_Date('" & VPLFecha & "', 'dd/mm/yyyy')" & " and Tarima = " & VPLTarima)
                                End If
                            
                            If RVerificaTarima.RecordCount > 0 Then
                                 mensaje = MsgBox("Tarima Ya Esta Liberada: Tarima " & VPLTarima & " De Ficha " & VPLFicha & " Con Fecha " & VPLFecha & " Ya Existe, No Se Puede Grabar ", vbOKOnly + vbInformation, "Verificacion")
                                 Exit Sub
                            End If
                            
                            'VERIFICA SI YA EXISTE LA TARIMA EN PRODUCCION NORMAL
                            Set RVerificaTarima = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RVerificaTarima, "Select * from Produccion Where Linea = '" & VPLLinea & "' and Esp_tec = '" & VPLFicha & "' and Fec_prd = #" & Format(VPLFecha, "mm/dd/yyyy") & "# and Tarima = " & VPLTarima)
                                Else 'ORACLE
                                    Call Abrir_Recordset(RVerificaTarima, "Select * from Produccion Where UPPER(Linea) = '" & UCase(VPLLinea) & "' and UPPER(Esp_tec) = '" & UCase(VPLFicha) & "' and Fec_prd = To_Date('" & VPLFecha & "', 'dd/mm/yyyy')" & " and Tarima = " & VPLTarima)
                                End If
                            If RVerificaTarima.RecordCount > 0 Then
                                 mensaje = MsgBox("La Tarima: " & VPLTarima & " De Ficha " & VPLFicha & " Con Fecha " & VPLFecha & " Ya Existe En La Captura De PRODUCCION, No Se Puede Grabar ", vbOKOnly + vbInformation, "Verificacion")
                                 Exit Sub
                            End If
                   End If
                   
                   If TxtTexto.Item(3).Text <> "" Then
                   
                        Set RBuscaOrden = New ADODB.Recordset
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                         Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where Documento = '" & TxtTexto.Item(3).Text & "' And FichaTecnica = '" & TxtTexto.Item(1).Text & "'")
                                     Else 'ORACLE
                                         Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtTexto.Item(3).Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtTexto.Item(1).Text) & "'")
                                     End If
                                     
                                         If RBuscaOrden.RecordCount > 0 Then
                                             
                                         Else
                                             MousePointer = 0
                                             MsgBox "Orden y Ficha Tecnica No Corresponden a la Ficha Tecnica que Tiene Esa Orden", vbOKOnly + vbInformation, "Informacion"
                                             Exit Sub
                                         End If
                   End If
                   
                   'GENERA CODIGO DE BARRAS
                    'TxtTexto.Item(4).Text = Format(TxtTexto.Item(2).Text, "dd-mm-yyyy") & "-" & TxtTexto.Item(0).Text & "-" & TxtTexto.Item(1).Text & "-" & TxtTexto.Item(6).Text
                    TxtTexto.Item(4).Text = Format(TxtTexto.Item(2).Text, "ddmmyy") & TxtTexto.Item(0).Text & TxtTexto.Item(1).Text & TxtTexto.Item(6).Text
                   
                                        
                   'GRABA DATOS
                   If BEditar = False Then
                        VCampos = "Linea, Esp_Tec, Fec_Prd, Tarima, Hor_Prd, Orden, Batch, Envases, Turno, NoMP9301, ColorMP9301, Cod_Emp, Calidad, Observaciones, Barra"
                        
                        VValores = "'" & TxtTexto.Item(0).Text & "'," 'LINEA
                        VValores = VValores & "'" & TxtTexto.Item(1).Text & "'," 'FICHATECNICA
                        If GOrigenDeDatos = "AmaproAccess" Then
                             VValores = VValores & "#" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "#," 'FECHA
                        Else 'ORACLE
                             VValores = VValores & "To_Date('" & Format(TxtTexto.Item(2).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & "," 'FECHA
                        End If
                        VValores = VValores & TxtTexto.Item(6).Text & "," 'TARIMA
                        VValores = VValores & "'" & MskHor.Text & "'," 'HORA
                        VValores = VValores & "'" & TxtTexto.Item(3).Text & "'," 'ORDEN
                        VValores = VValores & TxtTexto.Item(7).Text & "," 'BATCH
                        VValores = VValores & TxtTexto.Item(8).Text & "," 'ENVASES
                        VValores = VValores & "'" & TxtTur.Text & "'," 'TURNO
                        VValores = VValores & "'" & TxtTexto.Item(25).Text & "'," 'NO MP9301
                        VValores = VValores & "'" & CboColor.Text & "'," 'COLOR MP9301
                        VValores = VValores & "'" & TxtTexto.Item(16).Text & "'," 'COD EMPLEADO
                        VValores = VValores & "'" & CboCal.Text & "'," 'CALIDAD
                        VValores = VValores & "'" & TxtObservaciones.Text & "'," 'OBSERVACIONES
                        VValores = VValores & "'" & TxtTexto.Item(4).Text & "'" 'BARRA
                   
                        'INICIA UNA TRANSACCION
                       'SI ESTA GRABANDO UN REGISTRO NUEVO
                        Conexion.BeginTrans
                            Conexion.Execute "Insert Into ProduccionLiberada (" & VCampos & ") Values(" & VValores & ")"
                   'SI ESTA EDITANTO UN REGISTRO Y LUEGO LO GRABA
                   Else
                            VUpdate = "Linea = '" & TxtTexto.Item(0).Text & "',"
                            VUpdate = VUpdate & "Esp_Tec = '" & TxtTexto.Item(1).Text & "',"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VUpdate = VUpdate & "Fec_Prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "#," 'FECHA
                            Else 'ORACLE
                                VUpdate = VUpdate & "Fec_Prd = To_Date('" & Format(TxtTexto.Item(2).Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & "," 'FECHA
                            End If
                            VUpdate = VUpdate & "Tarima = " & TxtTexto.Item(6).Text & ","
                            VUpdate = VUpdate & "Hor_Prd = '" & MskHor.Text & "',"
                            VUpdate = VUpdate & "Orden = '" & TxtTexto.Item(3).Text & "',"
                            VUpdate = VUpdate & "Batch = " & TxtTexto.Item(7).Text & ","
                            VUpdate = VUpdate & "Envases = " & TxtTexto.Item(8).Text & ","
                            VUpdate = VUpdate & "Turno = '" & TxtTur.Text & "',"
                            VUpdate = VUpdate & "NoMp9301 = '" & TxtTexto.Item(25).Text & "',"
                            VUpdate = VUpdate & "ColorMP9301 = '" & CboColor.Text & "',"
                            VUpdate = VUpdate & "Cod_Emp = '" & TxtTexto.Item(16).Text & "',"
                            VUpdate = VUpdate & "Calidad = '" & CboCal.Text & "',"
                            VUpdate = VUpdate & "Observaciones = '" & TxtObservaciones.Text & "',"
                            VUpdate = VUpdate & "Barra = '" & TxtTexto.Item(4).Text & "'"
                            'VALIDA LA LLAVE DE LA BASE DE DATOS
                            If GOrigenDeDatos = "AmaproAccess" Then
                                    VUpdate = VUpdate & " Where Linea = '" & VPLLinea & "' and Esp_tec = '" & VPLFicha & "' and Fec_prd = #" & Format(VPLFecha, "mm/dd/yyyy") & "# And Tarima = " & VPLTarima
                            Else 'ORACLE
                                    VUpdate = VUpdate & " Where Linea = '" & VPLLinea & "' and Esp_tec = '" & VPLFicha & "' and Fec_prd = TO_DATE('" & VPLFecha & "', 'dd/mm/yyyy') and Tarima = " & VPLTarima
                            End If
                            'EJECUTA EL UPDATE
                            Conexion.Execute "Update ProduccionLiberada Set " & VUpdate
                   End If
                   
                   'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err = -2147467259 Then
                            Conexion.RollbackTrans
                            MsgBox "Fecha, Linea, Ficha Tecnica, Tarima Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147467259 And Err <> 0 Then
                            Conexion.RollbackTrans
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            Conexion.RollbackTrans
                            MsgBox "Fecha, Linea, Ficha Tecnica, Tarima Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            Conexion.RollbackTrans
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            TxtTexto.Item(0).SetFocus
                            Exit Sub
                        End If
                    End If
                    
                        'SI ESTA AGREGANDO DATOS
                        If BEditar = False Then
                                    'ACTUALIZA EL CONTADOR DE TARIMAS EN LAS LINEAS
                                    Conexion.Execute "Update Lineas set Tarima = Tarima + 1 Where Linea = '" & TxtTexto.Item(0) & "'"
                                    
                                        If Err <> 0 Then
                                                Conexion.RollbackTrans
                                                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                                                Err.Clear
                                                TxtTexto.Item(0).SetFocus
                                                MousePointer = 0
                                                
                                        End If
                                    
                        End If 'termina editar
                        
                      'TERMINA LA TRANSACCION SI ESTA AGREGANDO
                      'PORQUE CUANDO EDITA NO HAY NESESIDAD DE USAR EL BEGIN TRANS
                      If BEditar = False Then
                        Conexion.CommitTrans
                      End If
                                            
                      Bandera = False
                      botones
                      BVer = False
                  
                  RProduccion.Requery
                  RProduccion.MoveLast
                  
                  'VUELVE A PONER ENABLED LOS TEXTOS DEL CAMPO LLAVEN PARA QUE SE MIREN BIEN
                  TxtTexto.Item(0).Enabled = True
                  TxtTexto.Item(1).Enabled = True
                  TxtTexto.Item(2).Enabled = True
                  TxtTexto.Item(6).Enabled = True
                  
                  CmdBotones.Item(0).SetFocus
    
    'CANCELAR
    ElseIf Index = 3 Then
                    Bandera = False
                    botones
                    'VUELVE A LLENAS LOS CAMPOS CON EL RECORDSET ACTUAL
                    Llena_Campos
                    'VUELVE A PONER ENABLED LOS TEXTOS DEL CAMPO LLAVEN PARA QUE SE MIREN BIEN
                    TxtTexto.Item(0).Enabled = True
                    TxtTexto.Item(1).Enabled = True
                    TxtTexto.Item(2).Enabled = True
                    TxtTexto.Item(6).Enabled = True

    'BORRAR
    ElseIf Index = 4 Then
                MousePointer = 0
                 If GBorrar = True Then
                 Else
                        'ASIGNAMOS A LA VARIABLE FECHA DEL SISTEMA MENOS 1
                        VUltimaFecha = DateValue(Date) - 1
                        VFechaActual = DateValue(Date)
                        If (DateValue(TxtTexto.Item(2).Text) >= VUltimaFecha And DateValue(TxtTexto.Item(2).Text) <= VFechaActual) Then
                        Else
                            MsgBox "No Puede BORRAR Produccion De 2 o mas dias de la fecha actual, Llame al Encargado", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                 End If
    
                  mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
                  
                  If mensaje = vbOK Then
                       'BORRA EL REGISTRO
                       RProduccion.Delete
                       
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
                        RProduccion.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RProduccion.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        Llena_Campos
                  End If
    
    'IMPRIMIR
    ElseIf Index = 5 Then
                
                'BORRA LA IDENTIFICACION INGRESADA A LA BASE DE DATOS
                'Db.Execute "Delete * From ReporteIdentificacionInterno"
        
                'VDia = Day(TxtTexto.Item(2).Text)
                'VMes = Month(TxtTexto.Item(2).Text)
                'VAño = Year(TxtTexto.Item(2).Text)
                
                'Set RReporteIdentificacionInterno = Db.OpenRecordset("Select * From ReporteIdentificacionInterno")
                '            RReporteIdentificacionInterno.AddNew
                '                    RReporteIdentificacionInterno!Linea = TxtTexto.Item(0).Text
                '                    RReporteIdentificacionInterno!Esp_Tec = TxtTexto.Item(1).Text
                '                    RReporteIdentificacionInterno!fec_prd = TxtTexto.Item(2).Text
                '                    RReporteIdentificacionInterno!Tarima = TxtTexto.Item(6).Text
                '                    RReporteIdentificacionInterno!Envases = TxtTexto.Item(8).Text
                '                    RReporteIdentificacionInterno!Hor_prd = MskHor.Text
                '                    RReporteIdentificacionInterno!Batch = TxtTexto.Item(7).Text
                '                    RReporteIdentificacionInterno!Cod_Emp = TxtTexto.Item(16).Text
                '                    RReporteIdentificacionInterno!Orden = TxtTexto.Item(3).Text
                '            RReporteIdentificacionInterno.Update
                
                    
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "Identificacion2.rpt"
                Else
                    GNombreReporte = "Identificacion2O.rpt"
                End If
                GCriteriaReporte = "{produccionLiberada.fec_prd} = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# and {produccionLiberada.linea} = '" & TxtTexto.Item(0).Text & "' and {produccionLiberada.Esp_Tec} = '" & TxtTexto.Item(1).Text & "' and {produccionLiberada.tarima} = " & TxtTexto.Item(6).Text
                FrmReporte.Show
                
                'GCriteriaReporte = "{ReporteIdentificacionInterno.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño & "," & VMes & "," & VDia & ") and {ReporteIdentificacionInterno.Linea} = '" & TxtTexto.Item(0).Text & "' and {ReporteIdentificacionInterno.Tarima} = " & TxtTexto.Item(6).Text & " and {ReporteIdentificacionInterno.Esp_Tec} = '" & TxtTexto.Item(1).Text & "'"
                'GNombreReporte =  "\Identificacion2.rpt"
                'FrmReporte.Show
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
               
    'TARIMAS REVISADAS
    ElseIf Index = 6 Then
                 'ASIGNA VALOR A LA VARIABLE PARA QUE CUANDO SALGA DE LA FORMA DE CAPTURAR
                 'DEFECTOS Y REGRESE A ESTA FORMA SE ACTIVA EL EVENTO ACTIVATE PERO DABA ERROR
 '                BActivar = False
                
                'Asignacion de variables para agregar los defectos
                 VPLLinea = TxtTexto.Item(0).Text
                 VPLFicha = TxtTexto.Item(1).Text
                 VPLFecha = TxtTexto.Item(2).Text
                 VPLTarima = TxtTexto.Item(6).Text
                 
                 If VPLLinea = "" Then
                    MsgBox "La Linea No Puede Estar Vacia ", vbOKOnly + vbInformation, "Informacion"
                    TxtTexto.Item(0).SetFocus
                    Exit Sub
                 End If
                    
                 If VPLFicha = "" Then
                    MsgBox "La Ficha Tecnica No Puede Estar Vacia ", vbOKOnly + vbInformation, "Informacion"
                    TxtTexto.Item(1).SetFocus
                    Exit Sub
                 End If
                      
                 If Not IsDate(VPLFecha) Then
                    MsgBox "Fecha Incorrecta ", vbOKOnly + vbInformation, "Informacion"
                    TxtTexto.Item(2).SetFocus
                    Exit Sub
                 End If
                 
                 If Not IsNumeric(VPLTarima) Then
                    MsgBox "Numero De Tarima Incorrecta ", vbOKOnly + vbInformation, "Informacion"
                    TxtTexto.Item(6).SetFocus
                    Exit Sub
                 End If
                     
                    CapturaProduccionLiberadaTarimas.Show 1
    
    'IDENTIFICACION NO CONFORME
    ElseIf Index = 7 Then
    'BORRA LA IDENTIFICACION INGRESADA A LA BASE DE DATOS
                Conexion.Execute "Delete From ReporteIdentificacionInterno"
        
                VDia = Day(TxtTexto.Item(2).Text)
                VMes = Month(TxtTexto.Item(2).Text)
                VAño = Year(TxtTexto.Item(2).Text)
                
                'CAMPOS DE LA BASE DE DATOS
                VCampos = "Linea, Esp_Tec, Fec_Prd, Tarima, Envases, Hor_Prd, Batch, Cod_Emp, Orden"
                'VALORES A AGREGAR A LOS CAMPOS
                VValores = "'" & TxtTexto.Item(0).Text & "'," 'LINEA
                VValores = VValores & "'" & TxtTexto.Item(1).Text & "'," 'FICHATECNICA
                If GOrigenDeDatos = "AmaproAccess" Then
                    VValores = VValores & "#" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "#," 'FECHA
                Else 'ORACLE
                    VValores = VValores & "To_Date('" & TxtTexto.Item(2).Text & "', 'dd/mm/yyyy')" & "," 'FECHA
                End If
                VValores = VValores & TxtTexto.Item(6).Text & "," 'TARIMA
                VValores = VValores & TxtTexto.Item(8).Text & "," 'ENVASES
                VValores = VValores & "'" & MskHor.Text & "'," 'HORA
                VValores = VValores & TxtTexto.Item(7).Text & "," 'BATCH
                VValores = VValores & "'" & TxtTexto.Item(16).Text & "'," 'CODIGO EMPLEADO
                VValores = VValores & "'" & TxtTexto.Item(3).Text & "'" 'ORDEN
                'REALIZA EL INSERT
                Conexion.Execute "Insert Into ReporteIdentificacionInterno (" & VCampos & ") Values(" & VValores & ")"
                
                If Err <> 0 Then
                    MousePointer = 0
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
                End If

                
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "IdentificacionNoConforme.rpt"
                Else
                    GNombreReporte = "IdentificacionNoConformeO.rpt"
                End If
                
                GCriteriaReporte = "{ReporteidentificacionInterno.fec_prd} = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# and {ReporteidentificacionInterno.linea} = '" & TxtTexto.Item(0).Text & "' and {ReporteidentificacionInterno.Esp_Tec} = '" & TxtTexto.Item(1).Text & "' and {ReporteidentificacionInterno.tarima} = " & TxtTexto.Item(6).Text
                FrmReporte.Show
            
        If Err <> 0 Then
            MousePointer = 0
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
                
                
                
    
        'ETIQUETAS
    ElseIf Index = 9 Then
    
                                'BORRA LA IDENTIFICACION INGRESADA A LA BASE DE DATOS
                Conexion.Execute "Delete From ReporteIdentificacionInterno"
        
                VDia = Day(TxtTexto.Item(2).Text)
                VMes = Month(TxtTexto.Item(2).Text)
                VAño = Year(TxtTexto.Item(2).Text)
               
                'BUSCA LA CANTIDAD DE UNIDADES POR CAJA
                Set RBuscaUnidadesxCaja = New ADODB.Recordset
                Call Abrir_Recordset(RBuscaUnidadesxCaja, "Select Unidadesxcaja From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(1).Text & "'")
                    If RBuscaUnidadesxCaja.RecordCount > 0 Then
                        If IsNull(RBuscaUnidadesxCaja(0)) Then
                            VUnidadesxCaja = 0
                        Else
                            VUnidadesxCaja = RBuscaUnidadesxCaja(0)
                        End If
                    Else
                        VUnidadesxCaja = 0
                    End If
                    
                'CAMPOS DE LA BASE DE DATOS
                VCampos = "Linea, Esp_Tec, Fec_Prd, Tarima, Envases, Hor_Prd, Batch, Cod_Emp, Orden"
                'VALORES A AGREGAR A LOS CAMPOS
                VValores = "'" & TxtTexto.Item(0).Text & "'," 'LINEA
                VValores = VValores & "'" & TxtTexto.Item(1).Text & "'," 'FICHATECNICA
                If GOrigenDeDatos = "AmaproAccess" Then
                    VValores = VValores & "#" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "#," 'FECHA
                Else 'ORACLE
                    VValores = VValores & "To_Date('" & TxtTexto.Item(2).Text & "', 'dd/mm/yyyy')" & "," 'FECHA
                End If
                VValores = VValores & TxtTexto.Item(6).Text & "," 'TARIMA
                VValores = VValores & VUnidadesxCaja & "," 'UNIDADES X CAJA
                VValores = VValores & "'" & MskHor.Text & "'," 'HORA
                VValores = VValores & TxtTexto.Item(7).Text & "," 'BATCH
                VValores = VValores & "'" & TxtTexto.Item(16).Text & "'," 'CODIGO EMPLEADO
                VValores = VValores & "'" & TxtTexto.Item(3).Text & "'" 'ORDEN
                'REALIZA EL INSERT
                Cont = 0
                Do Until Cont > 5
                    Conexion.Execute "Insert Into ReporteIdentificacionInterno (" & VCampos & ") Values(" & VValores & ")"
                    If Err <> 0 Then
                        MousePointer = 0
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If
                    Cont = Cont + 1
                Loop
                
                
                
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "IdentificacionEtiquetas.rpt"
                Else
                    GNombreReporte = "IdentificacionEtiquetasO.rpt"
                End If
                GCriteriaReporte = "{ReporteIdentificacionInterno.fec_prd} = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# and {ReporteIdentificacionInterno.linea} = '" & TxtTexto.Item(0).Text & "' and {ReporteIdentificacionInterno.Esp_Tec} = '" & TxtTexto.Item(1).Text & "' and {ReporteIdentificacionInterno.tarima} = " & TxtTexto.Item(6).Text
                FrmReporte.Show
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
    
    'SALIDA
    ElseIf Index = 8 Then
        Unload Me
    End If
    
    
End Sub

Private Sub CmdBotones2_Click(Index As Integer)

            MousePointer = 11
                If Index = 1 Then
                    RProduccion.MoveFirst
                'REGISTRO ANTERIOR
                ElseIf Index = 2 Then
                    RProduccion.MovePrevious
                'SIGUIENTE REGISTRO
                ElseIf Index = 3 Then
                    RProduccion.MoveNext
                'ULTIMO REGISTRO
                ElseIf Index = 4 Then
                    RProduccion.MoveLast
                End If
                
                'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
                If RProduccion.BOF Then
                    RProduccion.MoveFirst
                ElseIf RProduccion.EOF Then
                    RProduccion.MoveLast
                End If
                
                'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
                Llena_Campos

            'TARIMAS Y DEFECTOS
            If TxtTexto.Item(2).Text <> "" Then
                TxtDatos.Text = ""
                'BUSCA LAS TARIMAS QUE INTEGRAN LA TARIMA LIBERADA
                Set BuscaTarimas = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(BuscaTarimas, "Select Plt.* from produccionLiberadaConTarimas Plt Where Linea = '" & TxtTexto.Item(0) & "' and Esp_tec = '" & TxtTexto.Item(1).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(6).Text)
                    Else 'ORACLE
                        Call Abrir_Recordset(BuscaTarimas, "Select Plt.* from produccionLiberadaConTarimas Plt Where Linea = '" & TxtTexto.Item(0) & "' and Esp_tec = '" & TxtTexto.Item(1).Text & "' and Fec_prd = To_Date('" & TxtTexto.Item(2).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(6).Text)
                    End If
                'SI HAY TARIMAS
                If BuscaTarimas.RecordCount > 0 Then
                        Do Until BuscaTarimas.EOF
                            TxtDatos.Text = TxtDatos.Text & Left(BuscaTarimas!fec_prdL & Space(10), 10) & Space(1) & Left(BuscaTarimas!LineaL & Space(2), 2) & Space(2) & Left(BuscaTarimas!Esp_TecL & Space(15), 15) & Space(2) & Left(BuscaTarimas!TarimaL & Space(4), 4) & Space(5) & Left(BuscaTarimas!CalidadL & Space(1), 1) & Space(5) & Left(BuscaTarimas!Revisados & Space(6), 6) & Space(2) & Left(BuscaTarimas!NoConforme & Space(6), 6) & Space(2) & Left(BuscaTarimas!Liberados & Space(6), 6) & Space(2) & Left(BuscaTarimas!EnTarima & Space(6), 6) & Space(5) & Left(BuscaTarimas!Minutos & Space(3), 3) & " " & Left(BuscaTarimas!Empleado & Space(10), 10) & vbCrLf
                                'BUSCA LOS DEFECTOS DE LA TARIMA
                                Set RBuscaDefectos = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaDefectos, "Select PLD.Defecto, D.Descrip, PLD.Cantidad from produccionLiberadaConDefectos PLD, Defectos D Where PLD.LineaL = '" & BuscaTarimas!LineaL & "' and PLD.Esp_tecL = '" & BuscaTarimas!Esp_TecL & "' and PLD.Fec_prdL = #" & Format(BuscaTarimas!fec_prdL, "mm/dd/yyyy") & "# and PLD.TarimaL = " & BuscaTarimas!TarimaL & " And PLD.Defecto = D.Defecto")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBuscaDefectos, "Select PLD.Defecto, D.Descrip, PLD.Cantidad from produccionLiberadaConDefectos PLD, Defectos D Where PLD.LineaL = '" & BuscaTarimas!LineaL & "' and PLD.Esp_tecL = '" & BuscaTarimas!Esp_TecL & "' and PLD.Fec_prdL = TO_Date('" & BuscaTarimas!fec_prdL & "', 'dd/mm/yyyy')" & " and PLD.TarimaL = " & BuscaTarimas!TarimaL & " And PLD.Defecto = D.Defecto")
                                    End If
                                    If RBuscaDefectos.RecordCount > 0 Then
                                        Do Until RBuscaDefectos.EOF
                                            TxtDatos.Text = TxtDatos.Text & Space(27) & Left(RBuscaDefectos!Defecto & Space(4), 4) & " " & Left(RBuscaDefectos!Descrip & Space(35), 35) & Right(Space(6) & RBuscaDefectos!Cantidad, 6) & vbCrLf
                                            RBuscaDefectos.MoveNext
                                        Loop
                                    End If
                            BuscaTarimas.MoveNext
                        Loop
                End If
            End If
            
            'BATCH
            If IsNumeric(TxtTexto.Item(7).Text) Then
                'CUENTA CUANTAS TARIMAS PRODUCCION LIBERADA
                Set RCuentaTarimas = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                            If RCuentaTarimas.RecordCount > 0 Then
                                LblBatch.Caption = RCuentaTarimas(0)
                            Else
                                LblBatch.Caption = 1
                            End If
                    
                'BUSCA EL ULTIMO BATCH DE LA LINEA
                Set RBuscaUltimoBatch = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaUltimoBatch, "Select max(Batch) From ProduccionLiberada Where Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaUltimoBatch, "Select max(Batch) From ProduccionLiberada Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                        If RBuscaUltimoBatch.RecordCount > 0 Then
                            If IsNull(RBuscaUltimoBatch(0)) Then
                                LblUltimoBatch.Caption = 0
                            Else
                                LblUltimoBatch.Caption = RBuscaUltimoBatch(0)
                            End If
                        Else
                            LblUltimoBatch.Caption = 1
                        End If
                
                'CUENTA CUANTAS TARIMAS PRODUCCION NORMAL
                Set RCuentaTarimas2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                        If RCuentaTarimas2.RecordCount > 0 Then
                            LblBatch2.Caption = RCuentaTarimas2(0)
                        Else
                            LblBatch2.Caption = 1
                        End If
                    
                
                'BUSCA EL ULTIMO BATCH DE LA LINEA PRODUCCION NORMAL
                Set RBuscaUltimoBatch2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaUltimoBatch2, "Select max(Batch) From Produccion Where Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaUltimoBatch2, "Select max(Batch) From Produccion Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                        If RBuscaUltimoBatch2.RecordCount > 0 Then
                            If IsNull(RBuscaUltimoBatch2(0)) Then
                                LblUltimoBatch2.Caption = 0
                            Else
                                LblUltimoBatch2.Caption = RBuscaUltimoBatch2(0)
                            End If
                        Else
                            LblUltimoBatch2.Caption = 1
                        End If
                
            End If
        MousePointer = 0
End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
                    
                    Set RProduccion = New ADODB.Recordset

            'FECHAS
            If OptOpcion.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# Order By Fec_Prd, Linea, Tarima")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy')" & " Order By Fec_Prd, Linea, Tarima")
                    End If
            'FECHAS Y LINEA
            ElseIf OptOpcion.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Linea = '" & TxtBuscar.Text & "' Order By Fec_Prd, Linea, Tarima")
                    Else 'ORACLE
                        Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy') And UPPER(Linea) = '" & UCase(TxtBuscar.Text) & "' Order By Fec_Prd, Linea, Tarima")
                    End If
            'FECHAS Y FICHA TECNICA
            ElseIf OptOpcion.Item(2).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where Fec_Prd >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(DTPFecFin.Value, "mm/dd/yyyy") & "# And Esp_Tec = '" & TxtBuscar.Text & "' Order By Fec_Prd, Linea, Tarima")
                    Else 'ORACLE
                            Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where Fec_Prd >= To_Date('" & DtpFecIni.Value & "', 'dd/mm/yyyy')" & " And Fec_Prd <= To_Date('" & DTPFecFin.Value & "', 'dd/mm/yyyy') And UPPER(Esp_Tec) = '" & UCase(TxtBuscar.Text) & "' Order By Fec_Prd, Linea, Tarima")
                    End If
            'BATCH
            ElseIf OptOpcion.Item(4).Value = True Then
                    If Not IsNumeric(TxtBuscar.Text) Then
                        MsgBox "Numero De Batch Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    Else
                        Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where batch = " & TxtBuscar.Text & " And Linea = '" & TxtBuscar2.Text & "' Order By Fec_Prd, Linea, Tarima")
                    End If
            'NO IDENTIFICACION
            ElseIf OptOpcion.Item(5).Value = True Then
                If Not IsNumeric(TxtBuscar.Text) Then
                        MsgBox "Numero De Identificacion Incorrecto", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                Else
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where NoMP9301 = " & TxtBuscar.Text & " And ColorMP9301 Like '" & TxtBuscar2.Text & "%' Order By Fec_Prd, Linea, Tarima")
                        Else 'ORACLE
                            Call Abrir_Recordset(RProduccion, "Select * from ProduccionLiberada where UPPER(NoMP9301) = " & UCase(TxtBuscar.Text) & " And UPPER(ColorMP9301) Like '" & UCase(TxtBuscar2.Text) & "%' Order By Fec_Prd, Linea, Tarima")
                        End If
                End If
            End If
            
                
                Set DBGridProduccion.DataSource = RProduccion

            If Err <> 0 Then
                MsgBox "Error" & Err.Number & Err.Description, vbOKOnly + vbInformation, "Error"
                Exit Sub
            End If
            
            SSTab1.Tab = 1

End Sub



Private Sub Dbgridconsultas_HeadClick(ByVal ColIndex As Integer)
            RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub


Private Sub DBGridProduccion_SelChange(Cancel As Integer)
        Llena_Campos2
End Sub

Private Sub Form_Activate()
            'TARIMAS Y DEFECTOS
            If TxtTexto.Item(2).Text <> "" Then
                TxtDatos.Text = ""
                'BUSCA LAS TARIMAS QUE INTEGRAN LA TARIMA LIBERADA
                Set BuscaTarimas = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(BuscaTarimas, "Select Plt.* from produccionLiberadaConTarimas Plt Where Linea = '" & TxtTexto.Item(0) & "' and Esp_tec = '" & TxtTexto.Item(1).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(6).Text)
                    Else 'ORACLE
                        Call Abrir_Recordset(BuscaTarimas, "Select Plt.* from produccionLiberadaConTarimas Plt Where Linea = '" & TxtTexto.Item(0) & "' and Esp_tec = '" & TxtTexto.Item(1).Text & "' and Fec_prd = To_Date('" & TxtTexto.Item(2).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(6).Text)
                    End If
                'SI HAY TARIMAS
                If BuscaTarimas.RecordCount > 0 Then
                        Do Until BuscaTarimas.EOF
                            TxtDatos.Text = TxtDatos.Text & Left(BuscaTarimas!fec_prdL & Space(10), 10) & Space(1) & Left(BuscaTarimas!LineaL & Space(2), 2) & Space(2) & Left(BuscaTarimas!Esp_TecL & Space(15), 15) & Space(2) & Left(BuscaTarimas!TarimaL & Space(4), 4) & Space(5) & Left(BuscaTarimas!CalidadL & Space(1), 1) & Space(5) & Left(BuscaTarimas!Revisados & Space(6), 6) & Space(2) & Left(BuscaTarimas!NoConforme & Space(6), 6) & Space(2) & Left(BuscaTarimas!Liberados & Space(6), 6) & Space(2) & Left(BuscaTarimas!EnTarima & Space(6), 6) & Space(5) & Left(BuscaTarimas!Minutos & Space(3), 3) & " " & Left(BuscaTarimas!Empleado & Space(10), 10) & vbCrLf
                                'BUSCA LOS DEFECTOS DE LA TARIMA
                                Set RBuscaDefectos = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaDefectos, "Select PLD.Defecto, D.Descrip, PLD.Cantidad from produccionLiberadaConDefectos PLD, Defectos D Where PLD.LineaL = '" & BuscaTarimas!LineaL & "' and PLD.Esp_tecL = '" & BuscaTarimas!Esp_TecL & "' and PLD.Fec_prdL = #" & Format(BuscaTarimas!fec_prdL, "mm/dd/yyyy") & "# and PLD.TarimaL = " & BuscaTarimas!TarimaL & " And PLD.Defecto = D.Defecto")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RBuscaDefectos, "Select PLD.Defecto, D.Descrip, PLD.Cantidad from produccionLiberadaConDefectos PLD, Defectos D Where PLD.LineaL = '" & BuscaTarimas!LineaL & "' and PLD.Esp_tecL = '" & BuscaTarimas!Esp_TecL & "' and PLD.Fec_prdL = TO_Date('" & BuscaTarimas!fec_prdL & "', 'dd/mm/yyyy')" & " and PLD.TarimaL = " & BuscaTarimas!TarimaL & " And PLD.Defecto = D.Defecto")
                                    End If
                                    If RBuscaDefectos.RecordCount > 0 Then
                                        Do Until RBuscaDefectos.EOF
                                            TxtDatos.Text = TxtDatos.Text & Space(27) & Left(RBuscaDefectos!Defecto & Space(4), 4) & " " & Left(RBuscaDefectos!Descrip & Space(35), 35) & Right(Space(6) & RBuscaDefectos!Cantidad, 6) & vbCrLf
                                            RBuscaDefectos.MoveNext
                                        Loop
                                    End If
                            BuscaTarimas.MoveNext
                        Loop
                End If
            End If
            
            
            'BATCH
            If IsNumeric(TxtTexto.Item(7).Text) Then
                'CUENTA CUANTAS TARIMAS PRODUCCION LIBERADA
                Set RCuentaTarimas = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                            If RCuentaTarimas.RecordCount > 0 Then
                                LblBatch.Caption = RCuentaTarimas(0)
                            Else
                                LblBatch.Caption = 1
                            End If
                    
                'BUSCA EL ULTIMO BATCH DE LA LINEA
                Set RBuscaUltimoBatch = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaUltimoBatch, "Select max(Batch) From ProduccionLiberada Where Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaUltimoBatch, "Select max(Batch) From ProduccionLiberada Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                        If RBuscaUltimoBatch.RecordCount > 0 Then
                            If IsNull(RBuscaUltimoBatch(0)) Then
                                LblUltimoBatch.Caption = 0
                            Else
                                LblUltimoBatch.Caption = RBuscaUltimoBatch(0)
                            End If
                        Else
                            LblUltimoBatch.Caption = 1
                        End If
                
                'CUENTA CUANTAS TARIMAS PRODUCCION NORMAL
                Set RCuentaTarimas2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                        If RCuentaTarimas2.RecordCount > 0 Then
                            LblBatch2.Caption = RCuentaTarimas2(0)
                        Else
                            LblBatch2.Caption = 1
                        End If
                    
                
                'BUSCA EL ULTIMO BATCH DE LA LINEA PRODUCCION NORMAL
                Set RBuscaUltimoBatch2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaUltimoBatch2, "Select max(Batch) From Produccion Where Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaUltimoBatch2, "Select max(Batch) From Produccion Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                        If RBuscaUltimoBatch2.RecordCount > 0 Then
                            If IsNull(RBuscaUltimoBatch2(0)) Then
                                LblUltimoBatch2.Caption = 0
                            Else
                                LblUltimoBatch2.Caption = RBuscaUltimoBatch2(0)
                            End If
                        Else
                            LblUltimoBatch2.Caption = 1
                        End If
                
            End If
End Sub

Private Sub MskHor_GotFocus()
            MskHor.SelStart = 0
            MskHor.SelLength = Len(MskHor.Text)
End Sub

Private Sub MskHor_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If
End Sub

Private Sub OptCodigo_Click()
        LblBusqueda.Caption = "Codigo"
End Sub

Private Sub OptDescripcion_Click()
        LblBusqueda.Caption = "Descripcion"
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        'FECHAS
        If OptOpcion.Item(0).Value = True Then
            TxtBuscar.Visible = False
            TxtBuscar2.Visible = False
            DtpFecIni.Visible = True
            DTPFecFin.Visible = True
            Lbletiqueta.Item(0).Caption = ""
            LblBuscar2.Caption = ""
            Lbletiqueta.Item(2).Visible = True
            Lbletiqueta.Item(3).Visible = True
        'FECHAS Y LINEA
        ElseIf OptOpcion.Item(1).Value = True Then
            TxtBuscar.Visible = True
            TxtBuscar2.Visible = False
            TxtBuscar.SetFocus
            DtpFecIni.Visible = True
            DTPFecFin.Visible = True
            Lbletiqueta.Item(0).Caption = "Numero Linea"
            LblBuscar2.Caption = ""
            Lbletiqueta.Item(2).Visible = True
            Lbletiqueta.Item(3).Visible = True
        'FECHAS Y FICHA TECNICA
        ElseIf OptOpcion.Item(2).Value = True Then
            TxtBuscar.Visible = True
            TxtBuscar2.Visible = False
            TxtBuscar.SetFocus
            DtpFecIni.Visible = True
            DTPFecFin.Visible = True
            Lbletiqueta.Item(0).Caption = "Ficha Tecnica"
            LblBuscar2.Caption = ""
            Lbletiqueta.Item(2).Visible = True
            Lbletiqueta.Item(3).Visible = True
        'BATCH
        ElseIf OptOpcion.Item(4).Value = True Then
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            TxtBuscar2.Visible = True
            LblBuscar2.Visible = True
            DtpFecIni.Visible = False
            DTPFecFin.Visible = False
            Lbletiqueta.Item(0).Caption = "Batch"
            Lbletiqueta.Item(2).Visible = False
            LblBuscar2.Caption = "Linea"
            Lbletiqueta.Item(3).Visible = False
        'IDENTIFICACION
        ElseIf OptOpcion.Item(5).Value = True Then
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            TxtBuscar2.Visible = True
            LblBuscar2.Visible = True
            DtpFecIni.Visible = False
            DTPFecFin.Visible = False
            Lbletiqueta.Item(0).Caption = "# Identificacion"
            Lbletiqueta.Item(2).Visible = False
            LblBuscar2.Caption = "Color De Identificacion"
            Lbletiqueta.Item(3).Visible = False
        End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
            If SSTab1.Tab = 0 Then
                CmdBotones.Item(4).Enabled = True
            
                If CmdBotones.Item(2).Enabled = False Then
                    Llena_Campos
                    
                    
                    
                        'TARIMAS Y DEFECTOS
                        If TxtTexto.Item(2).Text <> "" Then
                            TxtDatos.Text = ""
                            'BUSCA LAS TARIMAS QUE INTEGRAN LA TARIMA LIBERADA
                            Set BuscaTarimas = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(BuscaTarimas, "Select Plt.* from produccionLiberadaConTarimas Plt Where Linea = '" & TxtTexto.Item(0) & "' and Esp_tec = '" & TxtTexto.Item(1).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(6).Text)
                                Else 'ORACLE
                                    Call Abrir_Recordset(BuscaTarimas, "Select Plt.* from produccionLiberadaConTarimas Plt Where Linea = '" & TxtTexto.Item(0) & "' and Esp_tec = '" & TxtTexto.Item(1).Text & "' and Fec_prd = To_Date('" & TxtTexto.Item(2).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(6).Text)
                                End If
                            'SI HAY TARIMAS
                            If BuscaTarimas.RecordCount > 0 Then
                                    Do Until BuscaTarimas.EOF
                                        TxtDatos.Text = TxtDatos.Text & Left(BuscaTarimas!fec_prdL & Space(10), 10) & Space(1) & Left(BuscaTarimas!LineaL & Space(2), 2) & Space(2) & Left(BuscaTarimas!Esp_TecL & Space(15), 15) & Space(2) & Left(BuscaTarimas!TarimaL & Space(4), 4) & Space(5) & Left(BuscaTarimas!CalidadL & Space(1), 1) & Space(5) & Left(BuscaTarimas!Revisados & Space(6), 6) & Space(2) & Left(BuscaTarimas!NoConforme & Space(6), 6) & Space(2) & Left(BuscaTarimas!Liberados & Space(6), 6) & Space(2) & Left(BuscaTarimas!EnTarima & Space(6), 6) & Space(5) & Left(BuscaTarimas!Minutos & Space(3), 3) & " " & Left(BuscaTarimas!Empleado & Space(10), 10) & vbCrLf
                                            'BUSCA LOS DEFECTOS DE LA TARIMA
                                            Set RBuscaDefectos = New ADODB.Recordset
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RBuscaDefectos, "Select PLD.Defecto, D.Descrip, PLD.Cantidad from produccionLiberadaConDefectos PLD, Defectos D Where PLD.LineaL = '" & BuscaTarimas!LineaL & "' and PLD.Esp_tecL = '" & BuscaTarimas!Esp_TecL & "' and PLD.Fec_prdL = #" & Format(BuscaTarimas!fec_prdL, "mm/dd/yyyy") & "# and PLD.TarimaL = " & BuscaTarimas!TarimaL & " And PLD.Defecto = D.Defecto")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RBuscaDefectos, "Select PLD.Defecto, D.Descrip, PLD.Cantidad from produccionLiberadaConDefectos PLD, Defectos D Where PLD.LineaL = '" & BuscaTarimas!LineaL & "' and PLD.Esp_tecL = '" & BuscaTarimas!Esp_TecL & "' and PLD.Fec_prdL = TO_Date('" & BuscaTarimas!fec_prdL & "', 'dd/mm/yyyy')" & " and PLD.TarimaL = " & BuscaTarimas!TarimaL & " And PLD.Defecto = D.Defecto")
                                                End If
                                                If RBuscaDefectos.RecordCount > 0 Then
                                                    Do Until RBuscaDefectos.EOF
                                                        TxtDatos.Text = TxtDatos.Text & Space(27) & Left(RBuscaDefectos!Defecto & Space(4), 4) & " " & Left(RBuscaDefectos!Descrip & Space(35), 35) & Right(Space(6) & RBuscaDefectos!Cantidad, 6) & vbCrLf
                                                        RBuscaDefectos.MoveNext
                                                    Loop
                                                End If
                                        BuscaTarimas.MoveNext
                                    Loop
                            End If
                    End If
                End If
            ElseIf SSTab1.Tab = 1 Then
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

Private Sub TxtBuscar2_GotFocus()
        TxtBuscar2.SelStart = 0
        TxtBuscar2.SelLength = Len(TxtBuscar2.Text)
End Sub

Private Sub TxtBuscar2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtBusqueda_Change()
    'SI BUSCA LINEAS
    If VLineas = True Then
        Set RBusqueda = New ADODB.Recordset
        If OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBusqueda, "Select * from Lineas Where Descrip Like '%" & TxtBusqueda.Text & "%' Order by Linea")
            Else 'ORACLE
                Call Abrir_Recordset(RBusqueda, "Select * from Lineas Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%' Order by Linea")
            End If
        Else
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBusqueda, "Select * from Lineas Where Linea Like '%" & TxtBusqueda.Text & "%' Order by Linea")
            Else 'ORACLE
                Call Abrir_Recordset(RBusqueda, "Select * from Lineas Where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%' Order by Linea")
            End If
        End If
    
    End If
            
            Set DBGridConsultas.DataSource = RBusqueda
            DBGridConsultas.Columns(1).Width = "4000"
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

Private Sub Command1_Click()
    FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_DblClick()
    'PARA SELECCIONAR LA LINEA
    If VLineas = True Then
        TxtTexto.Item(0).Text = DBGridConsultas.Columns(0)
        TxtTexto.Item(0).SetFocus
        FrameConsultas.Visible = False
    End If
    
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 43 Then
          'PARA SELECCIONAR LA LINEA
          If VLineas = True Then
              TxtTexto.Item(0).Text = DBGridConsultas.Columns(0)
              TxtTexto.Item(0).SetFocus
              FrameConsultas.Visible = False
          End If
    End If
End Sub

Private Sub DbgridProduccion_HeadClick(ByVal ColIndex As Integer)
        RProduccion.Sort = RProduccion.Fields(ColIndex).Name
    
End Sub

Private Sub Form_Load()
            
            Set RProduccion = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RProduccion, "Select * From ProduccionLiberada Where Fec_Prd >= #" & Format((Date - 1), "mm/dd/yyyy") & "# And Fec_Prd <= #" & Format(Date, "mm/dd/yyyy") & "#")
            Else 'ORACLE
                Call Abrir_Recordset(RProduccion, "Select * From ProduccionLiberada Where Fec_Prd >= To_Date('" & (Date - 1) & "', 'dd/mm/yyyy') And Fec_Prd <= To_Date('" & Date & "', 'dd/mm/yyyy')")
            End If
                                    
            'LLENA EL DATA GRID CON EL RECORDSET
            Set DBGridProduccion.DataSource = RProduccion
            Llena_Campos
                
            'VALIDA SI EL USUARIO PUEDE EDITAR
            If GEditar = True Then
                DBGridProduccion.AllowUpdate = True
            Else
                DBGridProduccion.AllowUpdate = False
            End If
            
                BVer = False
                
                DtpFecIni.Value = Date
                DTPFecFin.Value = Date
            
    
End Sub

Private Sub TxtObservaciones_GotFocus()
        TxtObservaciones.SelStart = 0
        TxtObservaciones.SelLength = Len(TxtObservaciones.Text)
End Sub

Private Sub TxtObservaciones_KeyPress(KeyAscii As Integer)
        
        If KeyAscii = 13 Then
          SendKeys "{tab}"
        End If

End Sub

Private Sub TxtTexto_Change(Index As Integer)
'SI ESTA EN VISTA GENERAL QUE NO BUSQUE LOS CHANGE PORQUE SE TARDA MUCHO
If SSTab1.Tab = 1 Then
    Exit Sub
End If
    
    
    'LINEA
    If Index = 0 Then
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(0).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
            End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
    End If
    
    'FICHA TECNICA
    If Index = 1 Then
        Set RBuscaFichaTecnica = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select * From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(1).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaFichaTecnica, "Select * From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(1).Text) & "'")
            End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                        LblFichaTecnica.Caption = RBuscaFichaTecnica!Descrip
                        LblEmpaque.Caption = RBuscaFichaTecnica!MaterialEmpaque
                Else
                        LblFichaTecnica.Caption = ""
                        LblEmpaque.Caption = ""
                End If
        
                    'TARIMAS Y DEFECTOS
                    If TxtTexto.Item(2).Text <> "" Then
                        TxtDatos.Text = ""
                        If TxtTexto.Item(6).Text <> "" Then
                                Set BuscaTarimas = New ADODB.Recordset
                                'BUSCA LAS TARIMAS QUE INTEGRAN LA TARIMA LIBERADA
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(BuscaTarimas, "Select Plt.* from produccionLiberadaConTarimas Plt Where Linea = '" & TxtTexto.Item(0) & "' and Esp_tec = '" & TxtTexto.Item(1).Text & "' and Fec_prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# and Tarima = " & TxtTexto.Item(6).Text)
                                Else 'ORACLE
                                    Call Abrir_Recordset(BuscaTarimas, "Select Plt.* from produccionLiberadaConTarimas Plt Where Linea = '" & TxtTexto.Item(0) & "' and Esp_tec = '" & TxtTexto.Item(1).Text & "' and Fec_prd = To_Date('" & TxtTexto.Item(2).Text & "', 'dd/mm/yyyy')" & " and Tarima = " & TxtTexto.Item(6).Text)
                                End If
                                'SI HAY TARIMAS
                                If BuscaTarimas.RecordCount > 0 Then
                                        Do Until BuscaTarimas.EOF
                                            TxtDatos.Text = TxtDatos.Text & Left(BuscaTarimas!fec_prdL & Space(10), 10) & Space(1) & Left(BuscaTarimas!LineaL & Space(2), 2) & Space(2) & Left(BuscaTarimas!Esp_TecL & Space(15), 15) & Space(2) & Left(BuscaTarimas!TarimaL & Space(4), 4) & Space(5) & Left(BuscaTarimas!CalidadL & Space(1), 1) & Space(5) & Left(BuscaTarimas!Revisados & Space(6), 6) & Space(2) & Left(BuscaTarimas!NoConforme & Space(6), 6) & Space(2) & Left(BuscaTarimas!Liberados & Space(6), 6) & Space(2) & Left(BuscaTarimas!EnTarima & Space(6), 6) & Space(5) & Left(BuscaTarimas!Minutos & Space(3), 3) & " " & Left(BuscaTarimas!Empleado & Space(10), 10) & vbCrLf
                                                'BUSCA LOS DEFECTOS DE LA TARIMA
                                                Set RBuscaDefectos = New ADODB.Recordset
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Call Abrir_Recordset(RBuscaDefectos, "Select PLD.Defecto, D.Descrip, PLD.Cantidad from produccionLiberadaConDefectos PLD, Defectos D Where PLD.LineaL = '" & BuscaTarimas!LineaL & "' and PLD.Esp_tecL = '" & BuscaTarimas!Esp_TecL & "' and PLD.Fec_prdL = #" & Format(BuscaTarimas!fec_prdL, "mm/dd/yyyy") & "# and PLD.TarimaL = " & BuscaTarimas!TarimaL & " And PLD.Defecto = D.Defecto")
                                                    Else 'ORACLE
                                                        Call Abrir_Recordset(RBuscaDefectos, "Select PLD.Defecto, D.Descrip, PLD.Cantidad from produccionLiberadaConDefectos PLD, Defectos D Where UPPER(PLD.LineaL) = '" & UCase(BuscaTarimas!LineaL) & "' and UPPER(PLD.Esp_tecL) = '" & UCase(BuscaTarimas!Esp_TecL) & "' and PLD.Fec_prdL = To_Date('" & BuscaTarimas!fec_prdL & "', 'dd/mm/yyyy')" & " and UPPER(PLD.TarimaL) = " & UCase(BuscaTarimas!TarimaL) & " And UPPER(PLD.Defecto) = UPPER(D.Defecto)")
                                                    End If
                                                    If RBuscaDefectos.RecordCount > 0 Then
                                                        Do Until RBuscaDefectos.EOF
                                                            TxtDatos.Text = TxtDatos.Text & Space(27) & Left(RBuscaDefectos!Defecto & Space(4), 4) & " " & Left(RBuscaDefectos!Descrip & Space(35), 35) & Right(Space(6) & RBuscaDefectos!Cantidad, 6) & vbCrLf
                                                            RBuscaDefectos.MoveNext
                                                        Loop
                                                    End If
                                            BuscaTarimas.MoveNext
                                        Loop
                                End If
                        End If
                    End If
    End If
    
    'BATCH
    If Index = 7 Then
        If IsNumeric(TxtTexto.Item(7).Text) Then
            'CUENTA TARIMAS EN PRODUCCIO LIBERADA
            Set RCuentaTarimas = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                End If
                        If RCuentaTarimas.RecordCount > 0 Then
                            LblBatch.Caption = RCuentaTarimas(0)
                        Else
                            LblBatch.Caption = 1
                        End If
            'BUSCA EL ULTIMO BATCH DE LA LINEA
            Set RBuscaUltimoBatch = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaUltimoBatch, "Select max(Batch) From ProduccionLiberada Where Linea = '" & TxtTexto.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaUltimoBatch, "Select max(Batch) From ProduccionLiberada Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                End If
                If RBuscaUltimoBatch.RecordCount > 0 Then
                    If IsNull(RBuscaUltimoBatch(0)) Then
                        LblUltimoBatch.Caption = 0
                    Else
                        LblUltimoBatch.Caption = RBuscaUltimoBatch(0)
                    End If
                Else
                    LblUltimoBatch.Caption = 1
                End If
            
            'CUENTA CUANTAS TARIMAS PRODUCCION NORMAL
                Set RCuentaTarimas2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                    If RCuentaTarimas2.RecordCount > 0 Then
                        LblBatch2.Caption = RCuentaTarimas2(0)
                    Else
                        LblBatch2.Caption = 1
                    End If
                    
                
                'BUSCA EL ULTIMO BATCH DE LA LINEA PRODUCCION NORMAL
                Set RBuscaUltimoBatch2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaUltimoBatch2, "Select max(Batch) From Produccion Where Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaUltimoBatch2, "Select max(Batch) From Produccion Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                        If RBuscaUltimoBatch2.RecordCount > 0 Then
                            If IsNull(RBuscaUltimoBatch2(0)) Then
                                LblUltimoBatch2.Caption = 0
                            Else
                                LblUltimoBatch2.Caption = RBuscaUltimoBatch2(0)
                            End If
                        Else
                            LblUltimoBatch2.Caption = 1
                        End If
                
        End If
    End If
      
    
    
    
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
    'LINEA
    If Index = 0 Then
            VLineas = True
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select * from Lineas")
            TxtTexto.Item(0).Text = ""
            
            Set DBGridConsultas.DataSource = RBusqueda
            FrameConsultas.Visible = True
            DBGridConsultas.Columns(1).Width = "4000"
            TxtBusqueda.SetFocus
    End If

End Sub

Private Sub TxtTexto_GotFocus(Index As Integer)
            TxtTexto.Item(Index).SelStart = 0
            TxtTexto.Item(Index).SelLength = Len(TxtTexto.Item(Index))
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
    'SI PRECIONAN A ENTER EN CUALQUIER TEXT
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
            'LINEA
                If Index = 0 Then
                        VLineas = True
                        Set RBusqueda = New ADODB.Recordset
                        Call Abrir_Recordset(RBusqueda, "Select * from Lineas")
                        TxtTexto.Item(0).Text = ""
                        
                        Set DBGridConsultas.DataSource = RBusqueda
                        FrameConsultas.Visible = True
                        DBGridConsultas.Columns(1).Width = "4000"
                        TxtBusqueda.SetFocus
                End If
    End If
    
    
End Sub

Private Sub Txttexto_LostFocus(Index As Integer)
    'LINEA
    If Index = 0 Then
        'SI NO ESTA EDITANDO BUSCA LOS ULTIMOS DATOS
        If BEditar = False Then
                If TxtTexto.Item(0).Text = "+" Then
                
                ElseIf TxtTexto.Item(0).Text = "" Then
                
                Else
                                'VERIFICA SI LA FICHA TECNICA ESTA ACTIVA
                                Set RLineas = New ADODB.Recordset
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RLineas, "Select Esp_Tec, Tarima, Orden from Lineas Where Linea = '" & TxtTexto.Item(0).Text & "' and Activa = -1")
                                    Else 'ORACLE
                                        Call Abrir_Recordset(RLineas, "Select Esp_Tec, Tarima, Orden from Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "' and Activa = -1")
                                    End If
                                'SI LA LINEA ESTA ACTIVA
                                If RLineas.RecordCount > 0 Then
                                                        'FICHA TECNICA
                                                        TxtTexto.Item(1).Text = RLineas!Esp_Tec
                                                        'TARIMA
                                                        TxtTexto.Item(6).Text = Val(RLineas!Tarima) + 1
                                                        'ORDEN DE PRODUCCION
                                                        TxtTexto.Item(3).Text = RLineas!Orden
                                                        
                                                        'BUSCA LA FICHA TECNICA Y JALA LA CANTIDAD DE ENVASES
                                                        Set RBuscaFichaTecnica = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                Call Abrir_Recordset(RBuscaFichaTecnica, "Select * From FichaTecnica Where Esp_Tec = '" & RLineas!Esp_Tec & "'")
                                                            Else 'ORACLE
                                                                Call Abrir_Recordset(RBuscaFichaTecnica, "Select * From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(RLineas!Esp_Tec) & "'")
                                                            End If
                                                        
                                                        'SI ENCUENTRA LA FICHA TECNICA
                                                        If RBuscaFichaTecnica.RecordCount > 0 Then
                                                            'ENVASES
                                                            TxtTexto.Item(8).Text = RBuscaFichaTecnica!Envases
                                                        Else
                                                            TxtTexto.Item(8).Text = 0
                                                        End If
                                                        
                                                        'BUSCA EL ULTIMO BATCH DE LA LINEA
                                                                Set RBuscaUltimoBatch = New ADODB.Recordset
                                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                                        Call Abrir_Recordset(RBuscaUltimoBatch, "Select max(Batch) From ProduccionLiberada Where Linea = '" & TxtTexto.Item(0).Text & "'")
                                                                    Else 'ORACLE
                                                                        Call Abrir_Recordset(RBuscaUltimoBatch, "Select max(Batch) From ProduccionLiberada Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                                                                    End If
                                                                    If RBuscaUltimoBatch.RecordCount > 0 Then
                                                                        If IsNull(RBuscaUltimoBatch(0)) Then
                                                                             LblUltimoBatch.Caption = 0
                                                                        Else
                                                                             LblUltimoBatch.Caption = RBuscaUltimoBatch(0)
                                                                        End If
                                                                    Else
                                                                        LblUltimoBatch.Caption = 0
                                                                    End If
                                                        
                                                        
                                                'BUSCA EL ULTIMO REGISTRO INGRESADO Y EXTRAE LOS DATOS
                                                        Set RBuscaProduccion = New ADODB.Recordset
                                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                                    Call Abrir_Recordset(RBuscaProduccion, "Select * From ProduccionLiberada Where Linea = '" & TxtTexto.Item(0).Text & "' and Esp_Tec = '" & RLineas!Esp_Tec & "' and Fec_Prd = #" & Format(TxtTexto.Item(2).Text, "mm/dd/yyyy") & "# Order By Tarima")
                                                            Else 'ORACLE
                                                                    Call Abrir_Recordset(RBuscaProduccion, "Select * From ProduccionLiberada Where UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "' and UPPER(Esp_Tec) = '" & UCase(RLineas!Esp_Tec) & "' and Fec_Prd = To_Date('" & TxtTexto.Item(2).Text & "', 'dd/mm/yyyy')" & " Order By Tarima")
                                                            End If
                                                        
                                                        If RBuscaProduccion.RecordCount > 0 Then
                                                                'SE MUEVE AL ULTIMO REGISTRO
                                                                RBuscaProduccion.MoveLast
                                                                                                        
                                                                'TURNO
                                                                If Not IsNull(RBuscaProduccion!Turno) Then
                                                                    TxtTur.Text = RBuscaProduccion!Turno
                                                                End If
                                                                
                                                                'COLOR DE HOJA
                                                                If Not IsNull(RBuscaProduccion!ColorMP9301) Then
                                                                    CboColor.Text = RBuscaProduccion!ColorMP9301
                                                                End If
                                                                
                                                                'BATCH
                                                                If Not IsNull(RBuscaProduccion!Batch) Then
                                                                    TxtTexto.Item(7).Text = RBuscaProduccion!Batch
                                                                End If
                                                                
                                                                'CUENTA CUANTAS TARIMAS LLEVA EL BATCH
                                                                Set RCuentaTarimas = New ADODB.Recordset
                                                                Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text)
                                                                If RCuentaTarimas.RecordCount > 0 Then
                                                                            LblBatch.Caption = RCuentaTarimas(0)
                                                                Else
                                                                            LblBatch.Caption = 1
                                                                End If
                                                                
                                                                'CUENTA CUANTAS TARIMAS PRODUCCION NORMAL
                                                                Set RCuentaTarimas2 = New ADODB.Recordset
                                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                                                                    Else 'ORACLE
                                                                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                                                                    End If
                                                                    If RCuentaTarimas2.RecordCount > 0 Then
                                                                        LblBatch2.Caption = RCuentaTarimas2(0)
                                                                    Else
                                                                        LblBatch2.Caption = 1
                                                                    End If
                                                                    
                
                
                                                                
                                                                'USUARIO
                                                                If Not IsNull(RBuscaProduccion!Cod_Emp) Then
                                                                    TxtTexto.Item(16).Text = RBuscaProduccion!Cod_Emp
                                                                End If
                                                                
                                                        
                                                        End If
                                
                            Else
                                    MsgBox "Esta Linea No Esta Activa", vbOKOnly + vbExclamation, "Informacion"
                            End If
                End If
        End If
    'FECHA
    ElseIf Index = 2 Then
                                'VALIDA LA FECHA
                                If IsDate(TxtTexto.Item(2).Text) Then
                                            TxtTexto.Item(2).Text = Format(TxtTexto.Item(2).Text, "dd/mm/yyyy")
                                End If
    
    'BATCH
    ElseIf Index = 7 Then
        If IsNumeric(TxtTexto.Item(7).Text) Then
            Set RCuentaTarimas = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From ProduccionLiberada Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                End If
                
                If RCuentaTarimas.RecordCount > 0 Then
                    LblBatch.Caption = RCuentaTarimas(0)
                Else
                    LblBatch.Caption = 1
                End If
                
                'CUENTA CUANTAS TARIMAS PRODUCCION NORMAL
                Set RCuentaTarimas2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And Linea = '" & TxtTexto.Item(0).Text & "'")
                    Else
                        Call Abrir_Recordset(RCuentaTarimas2, "Select Count(*) From Produccion Where Batch = " & TxtTexto.Item(7).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                    If RCuentaTarimas2.RecordCount > 0 Then
                        LblBatch2.Caption = RCuentaTarimas2(0)
                    Else
                        LblBatch2.Caption = 1
                    End If
                    
                
                
        Else
            TxtTexto.Item(7).Text = "0"
        End If
    End If
End Sub

Private Sub TxtTur_GotFocus()
        TxtTur.SelStart = 0
        TxtTur.SelLength = Len(TxtTur.Text)
End Sub

Private Sub TxtTur_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Public Sub Llena_Campos()
On Error Resume Next
            If RProduccion.RecordCount > 0 Then
                TxtTexto.Item(0).Text = RProduccion!Linea
                TxtTexto.Item(1).Text = RProduccion!Esp_Tec
                TxtTexto.Item(2).Text = RProduccion!fec_prd
                TxtTexto.Item(6).Text = RProduccion!Tarima
                If Not IsNull(RProduccion!Hor_prd) Then
                    MskHor.Text = RProduccion!Hor_prd
                Else
                    MskHor.Text = ""
                End If
                If Not IsNull(RProduccion!Orden) Then
                    TxtTexto.Item(3).Text = RProduccion!Orden
                Else
                    TxtTexto.Item(3).Text = ""
                End If
                TxtTexto.Item(7).Text = RProduccion!Batch
                TxtTexto.Item(8).Text = RProduccion!Envases
                'xtTexto.Item(9).Text = RProduccion!Muestra
                If Not IsNull(RProduccion!Turno) Then
                    TxtTur.Text = RProduccion!Turno
                Else
                    TxtTur.Text = ""
                End If
                If Not IsNull(RProduccion!NoMP9301) Then
                    TxtTexto.Item(25).Text = RProduccion!NoMP9301
                Else
                    TxtTexto.Item(25).Text = ""
                End If
                If Not IsNull(RProduccion!ColorMP9301) Then
                    CboColor.Text = RProduccion!ColorMP9301
                Else
                    CboColor.Text = ""
                End If
                If Not IsNull(RProduccion!Cod_Emp) Then
                    TxtTexto.Item(16).Text = RProduccion!Cod_Emp
                Else
                    TxtTexto.Item(16).Text = ""
                End If
                If Not IsNull(RProduccion!Calidad) Then
                    CboCal.Text = RProduccion!Calidad
                Else
                    CboCal.Text = ""
                End If
                If Not IsNull(RProduccion!Observaciones) Then
                    TxtObservaciones.Text = RProduccion!Observaciones
                Else
                    TxtObservaciones.Text = ""
                End If
                If Not IsNull(RProduccion!Barra) Then
                    TxtTexto.Item(4).Text = RProduccion!Barra
                Else
                    TxtTexto.Item(4).Text = ""
                End If
                
                If Err <> 0 Then
                    'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
                End If
            Else
                TxtTexto.Item(0).Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(2).Text = ""
                TxtTexto.Item(6).Text = ""
                MskHor.Text = ""
                TxtTexto.Item(3).Text = ""
                TxtTexto.Item(7).Text = ""
                TxtTexto.Item(8).Text = ""
                'xtTexto.Item(9).Text = ""
                TxtTur.Text = ""
                TxtTexto.Item(25).Text = ""
                CboColor.Text = ""
                TxtTexto.Item(16).Text = ""
                CboCal.Text = ""
                TxtObservaciones.Text = ""
                TxtTexto.Item(4).Text = ""
            End If

End Sub

Public Sub Limpia_Campos()
                TxtTexto.Item(0).Text = ""
                TxtTexto.Item(1).Text = ""
                TxtTexto.Item(2).Text = ""
                TxtTexto.Item(6).Text = ""
                MskHor.Text = ""
                TxtTexto.Item(3).Text = ""
                TxtTexto.Item(7).Text = ""
                TxtTexto.Item(8).Text = ""
                'xtTexto.Item(9).Text = ""
                TxtTur.Text = ""
                TxtTexto.Item(25).Text = ""
                CboColor.Text = ""
                TxtTexto.Item(16).Text = ""
                CboCal.Text = ""
                TxtObservaciones.Text = ""
                TxtTexto.Item(4).Text = ""
End Sub


Public Sub Llena_Campos2()
            If RProduccion.RecordCount > 0 Then
                TxtTexto.Item(0).Text = RProduccion!Linea
                TxtTexto.Item(1).Text = RProduccion!Esp_Tec
                TxtTexto.Item(2).Text = RProduccion!fec_prd
                TxtTexto.Item(6).Text = RProduccion!Tarima
                If Not IsNull(RProduccion!Hor_prd) Then
                    MskHor.Text = RProduccion!Hor_prd
                Else
                    MskHor.Text = ""
                End If
                If Not IsNull(RProduccion!Orden) Then
                    TxtTexto.Item(3).Text = RProduccion!Orden
                Else
                    TxtTexto.Item(3).Text = ""
                End If
                TxtTexto.Item(7).Text = RProduccion!Batch
                TxtTexto.Item(8).Text = RProduccion!Envases
                If Not IsNull(RProduccion!Cod_Emp) Then
                    TxtTexto.Item(16).Text = RProduccion!Cod_Emp
                Else
                    TxtTexto.Item(16).Text = ""
                End If
                If Not IsNull(RProduccion!Barra) Then
                    TxtTexto.Item(4).Text = RProduccion!Barra
                Else
                    TxtTexto.Item(4).Text = ""
                End If
            End If
End Sub
