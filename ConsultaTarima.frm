VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ConsultaTarima 
   BackColor       =   &H80000003&
   Caption         =   "Consulta De Bulto/Tarima"
   ClientHeight    =   8130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11745
   Icon            =   "ConsultaTarima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11745
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   15
      Top             =   1680
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   11245
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   -2147483645
      TabCaption(0)   =   "Entrada y Traslados"
      TabPicture(0)   =   "ConsultaTarima.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(10)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(8)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(7)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(6)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(15)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(13)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(11)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(14)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(16)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(17)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(18)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "DbgridTraslados"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "MskFecEnt"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "TxtTexto(9)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "TxtTexto(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "TxtTexto(6)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtTexto(5)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "TxtTexto(4)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "TxtTexto(3)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "TxtTexto(15)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TxtTexto(13)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "TxtTexto(10)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "TxtTexto(11)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "TxtTexto(12)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtTexto(14)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtTexto(16)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "TxtTexto(17)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      TabCaption(1)   =   "Salidas y Cierre Bulto y Consumos"
      TabPicture(1)   =   "ConsultaTarima.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DataGridConsumos"
      Tab(1).Control(1)=   "DbGridDespachos"
      Tab(1).Control(2)=   "DbGridCierreTarima"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Ajustes"
      TabPicture(2)   =   "ConsultaTarima.frx":0F06
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DbGridAjustes"
      Tab(2).ControlCount=   1
      Begin VB.TextBox TxtTexto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   17
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TxtTexto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   16
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
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
         Index           =   14
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   1920
         Width           =   3855
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   12
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
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
         Index           =   11
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2640
         Width           =   3855
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   10
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1095
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
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   13
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1095
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   15
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1095
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
         ForeColor       =   &H00008000&
         Height          =   285
         Index           =   6
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   840
         Width           =   1095
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
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   7
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtTexto 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
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
         Index           =   9
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   1200
         Width           =   3855
      End
      Begin MSMask.MaskEdBox MskFecEnt 
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   840
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSDataGridLib.DataGrid DbgridTraslados 
         Height          =   3135
         Left            =   120
         TabIndex        =   33
         Top             =   3120
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   5530
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
         Caption         =   "Traslados"
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
      Begin MSDataGridLib.DataGrid DbGridCierreTarima 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   34
         Top             =   2280
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8438015
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
         Caption         =   "Cierre Bulto/Tarima"
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
      Begin MSDataGridLib.DataGrid DbGridDespachos 
         Height          =   1455
         Left            =   -74880
         TabIndex        =   35
         Top             =   720
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   2566
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12640511
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
         Caption         =   "Salidas"
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
      Begin MSDataGridLib.DataGrid DbGridAjustes 
         Height          =   2535
         Left            =   -74880
         TabIndex        =   46
         Top             =   720
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   4471
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   12648447
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
         Caption         =   "Ajustes"
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
      Begin MSDataGridLib.DataGrid DataGridConsumos 
         Height          =   1935
         Left            =   -74880
         TabIndex        =   47
         Top             =   4320
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   3413
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   8438015
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
         Caption         =   "Consumos"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         Height          =   195
         Index           =   18
         Left            =   3000
         TabIndex        =   49
         Top             =   2280
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         Height          =   195
         Index           =   17
         Left            =   3000
         TabIndex        =   45
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Kilos"
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
         Left            =   10680
         TabIndex        =   43
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   16
         Left            =   360
         TabIndex        =   41
         Top             =   1920
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   38
         Top             =   2640
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Kilos"
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
         Left            =   10680
         TabIndex        =   36
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Cantidad Entrada"
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
         Index           =   13
         Left            =   7920
         TabIndex        =   32
         Top             =   2280
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Cantidad Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   15
         Left            =   7920
         TabIndex        =   31
         Top             =   2640
         Width           =   1365
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Fecha Entrada"
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
         Left            =   360
         TabIndex        =   30
         Top             =   840
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "# Transaccion"
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
         Left            =   360
         TabIndex        =   29
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         Index           =   6
         Left            =   360
         TabIndex        =   28
         Top             =   1560
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
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
         Index           =   7
         Left            =   360
         TabIndex        =   27
         Top             =   2280
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Peso Entrada"
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
         Left            =   8280
         TabIndex        =   26
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H000080FF&
         Caption         =   "Peso Actual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   10
         Left            =   8280
         TabIndex        =   25
         Top             =   1200
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000080FF&
         BackStyle       =   1  'Opaque
         Height          =   2295
         Index           =   0
         Left            =   120
         Top             =   720
         Width           =   11295
      End
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
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
      ForeColor       =   &H8000000D&
      Height          =   285
      Index           =   8
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin MSMask.MaskEdBox MskFecPro 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      Format          =   "dd/mm/yyyy"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   1
      Text            =   "77"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   11040
      Picture         =   "ConsultaTarima.frx":0F22
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salida"
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton CmdConsultar 
      Height          =   495
      Left            =   11040
      Picture         =   "ConsultaTarima.frx":2F94
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "consultar"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox TxtTexto 
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label LblBod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   14
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      Caption         =   "Ubicacion"
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
      Left            =   3120
      TabIndex        =   13
      Top             =   1200
      Width           =   870
   End
   Begin VB.Label LblLin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
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
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   3120
      TabIndex        =   11
      Top             =   480
      Width           =   7200
   End
   Begin VB.Label LblFicTec 
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
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
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   10
      Top             =   840
      Width           =   7200
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
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
      Index           =   3
      Left            =   240
      TabIndex        =   9
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
      Caption         =   "Fecha Produccion"
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
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
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
      Index           =   1
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000003&
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
      Left            =   240
      TabIndex        =   6
      Top             =   840
      Width           =   1230
   End
End
Attribute VB_Name = "ConsultaTarima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaLinea2 As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RBuscaEncabezadoEntrada As New ADODB.Recordset
Dim RBuscaProveedor As New ADODB.Recordset
Dim RBuscaTransportista As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset


Dim RTraslados As New ADODB.Recordset
Dim RSalidas As New ADODB.Recordset
Dim RCierreTarima As New ADODB.Recordset
Dim RAjustes As New ADODB.Recordset
Dim RConsumos As New ADODB.Recordset

Dim Columnas As String
Dim Tablas As String
Dim Criteria As String
Dim VPesoUnidad As Currency


Private Sub CmdConsultar_Click()
On Error Resume Next
        If GOrigenDeDatos = "AmaproOracle" Then
            MskFecPro.Text = Format(MskFecPro.Text, "dd/mm/yyyy")
        End If
        'REVISA FECHA DE TARIMA
        If Not IsDate(MskFecPro.Text) Then
                MsgBox "Fecha De Tarima Incorrecta", vbInformation + vbCritical, "Error"
                MskFecPro.SetFocus
                Exit Sub
        End If
        'REVISA SI ES NUMERICO LA TARIMA
        If Not IsNumeric(TxtTexto.Item(1).Text) Then
                MsgBox "Numero De Tarima Incorrecto", vbInformation + vbCritical, "Error"
                TxtTexto.Item(1).SetFocus
                Exit Sub
        End If

        'BUSCA LA TARIMA
        Set RBuscaTarima = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaTarima, "Select D.Documento, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, E.FechaEntrada, D.PesoEntrada, DD.Descripcion, E.Proveedor, E.Linea, E.NumeroDocumento, D.OrdenProduccion From DetalleEntradasInventario D, EncabezadoEntradasInventario E, Documentos DD Where D.FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And D.Tarima = " & TxtTexto.Item(1).Text & " And D.Linea = '" & TxtTexto.Item(2).Text & "' And D.FichaTecnica = '" & TxtTexto.Item(0).Text & "' And D.Documento = E.Documento And E.TipoDeDocumento = DD.CodigoDocumento")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaTarima, "Select D.Documento, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, E.FechaEntrada, D.PesoEntrada, DD.Descripcion, E.Proveedor, E.Linea, E.NumeroDocumento, D.OrdenProduccion From DetalleEntradasInventario D, EncabezadoEntradasInventario E, Documentos DD Where D.FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And D.Tarima = " & TxtTexto.Item(1).Text & " And UPPER(D.Linea) = '" & UCase(TxtTexto.Item(2).Text) & "' And UPPER(D.FichaTecnica) = '" & UCase(TxtTexto.Item(0).Text) & "' And D.Documento = E.Documento And UPPER(E.TipoDeDocumento) = UPPER(DD.CodigoDocumento)")
            End If
            
            
            If RBuscaTarima.RecordCount > 0 Then
                TxtTexto.Item(3).Text = RBuscaTarima(0)
                TxtTexto.Item(4).Text = RBuscaTarima(1)
                TxtTexto.Item(5).Text = RBuscaTarima(2)
                TxtTexto.Item(8).Text = RBuscaTarima(3)
                TxtTexto.Item(13).Text = Format(RBuscaTarima(4), "#,###,##0.00") 'CANTIDAD ENTRADA
                TxtTexto.Item(15).Text = Format(RBuscaTarima(5), "#,###,##0.00") ' SALDO
                MskFecEnt.Text = RBuscaTarima(6)
                TxtTexto.Item(6).Text = Format(RBuscaTarima(7), "#,###,##0.00") 'PESO ENTRADA
                If RBuscaTarima(4) > 0 Then
                    VPesoUnidad = RBuscaTarima(7) / RBuscaTarima(4)
                Else
                    VPesoUnidad = 0
                End If
                TxtTexto.Item(7).Text = Format(RBuscaTarima(5) * VPesoUnidad, "#,###,##0.00") 'PESO ACTUAL
                TxtTexto.Item(9).Text = RBuscaTarima(8)
                TxtTexto.Item(10).Text = RBuscaTarima(9)
                TxtTexto.Item(12).Text = RBuscaTarima(10)
                If IsNull(RBuscaTarima(11)) Then
                        TxtTexto.Item(16).Text = ""
                Else
                        TxtTexto.Item(16).Text = RBuscaTarima(11)
                End If
                If IsNull(RBuscaTarima(12)) Then
                        TxtTexto.Item(17).Text = ""
                Else
                        TxtTexto.Item(17).Text = RBuscaTarima(12)
                End If
            Else
                MsgBox "Tarima/Bulto No Existe", vbOKOnly + vbInformation, "Informacion"
                TxtTexto.Item(1).SetFocus
                Exit Sub
            End If
            
'BUSCA LOS TRASLADOS DEL PRODUCTO TERMINADO ________________________________________________________
        Set RTraslados = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RTraslados, "Select ED.Fecha, ED.Documento, ED.NumeroDocumento, D.Descripcion, B.Descripcion, B2.Descripcion, DD.CantidadSalida, DD.DiferenciaReqCorMas, DD.DiferenciaReqCor, DD.CantidadDesperdicio, DD.CantidadDesperdicioProveedor, DD.CantidadReal From DetalleTrasladosInventario DD, EncabezadoTrasladosInventario ED, Documentos D, BodegasInventario B, BodegasInventario B2 Where DD.FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And DD.Tarima = " & TxtTexto.Item(1).Text & " And DD.LineaProduccion = '" & TxtTexto.Item(2).Text & "' And DD.FichaTecnica = '" & TxtTexto.Item(0).Text & "' AND DD.Documento = ED.Documento And ED.TipoDeDocumento = D.CodigoDocumento And ED.BodegaSalida = B.CodigoBodega And DD.BodegaEntrada = B2.CodigoBodega Order By ED.Fecha")
                ' Call Abrir_Recordset(RTraslados, "Select ED.Fecha, ED.NumeroDocumento, D.Descripcion, B.Descripcion, D2.Descripcion, DD.CantidadSalida, DD.DiferenciaReqCorMas, DD.DiferenciaReqCor, DD.CantidadDesperdicio, DD.CantidadDesperdicioProveedor, DD.CantidadReal From DetalleTrasladosInventario DD, EncabezadoTrasladosInventario ED, Documentos D, BodegasInventario B Where DD.FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And DD.Tarima = " & TxtTexto.Item(1).Text & " And DD.LineaProduccion = '" & TxtTexto.Item(2).Text & "' And DD.FichaTecnica = '" & TxtTexto.Item(0).Text & "' AND DD.Documento = ED.Documento And ED.TipoDeDocumento = D.CodigoDocumento And ED.BodegaSalida = B.CodigoBodega And DD.BodegaEntrada = B2.CodigoBodega Order By ED.Fecha")
            Else 'ORACLE
                Call Abrir_Recordset(RTraslados, "Select ED.Fecha, ED.NumeroDocumento, D.Descripcion, B.Descripcion, B2.Descripcion, DD.CantidadSalida, DD.DiferenciaReqCorMas, DD.DiferenciaReqCor, DD.CantidadDesperdicio, DD.CantidadDesperdicioProveedor, DD.CantidadReal From DetalleTrasladosInventario DD, EncabezadoTrasladosInventario ED, Documentos D, BodegasInventario B, BodegasInventario B2 Where DD.FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And DD.Tarima = " & TxtTexto.Item(1).Text & " And UPPER(DD.LineaProduccion) = '" & UCase(TxtTexto.Item(2).Text) & "' And UPPER(DD.FichaTecnica) = '" & UCase(TxtTexto.Item(0).Text) & "' AND DD.Documento = ED.Documento And UPPER(ED.TipoDeDocumento) = UPPER(D.CodigoDocumento) And UPPER(ED.BodegaSalida) = UPPER(B.CodigoBodega) And UPPER(DD.BodegaEntrada) = UPPER(B2.CodigoBodega) Order By ED.Fecha")
            End If
            
            
            Set DbgridTraslados.DataSource = RTraslados
            
            If Err <> 0 Then
                MsgBox Err.Number & Err.Description
            End If
        
            DbgridTraslados.Columns(0).Width = "1000"
            DbgridTraslados.Columns(1).Width = "800"
            DbgridTraslados.Columns(2).Width = "800"
            DbgridTraslados.Columns(3).Width = "1000"
            DbgridTraslados.Columns(4).Width = "1400"
            DbgridTraslados.Columns(5).Width = "1400"
            DbgridTraslados.Columns(6).Width = "1000"
            DbgridTraslados.Columns(7).Width = "800"
            DbgridTraslados.Columns(8).Width = "800"
            DbgridTraslados.Columns(9).Width = "800"
            DbgridTraslados.Columns(10).Width = "800"
            DbgridTraslados.Columns(11).Width = "1000"
            
            DbgridTraslados.Columns(0).Caption = "Fecha"
            DbgridTraslados.Columns(1).Caption = "Transaccion"
            DbgridTraslados.Columns(2).Caption = "# Documento"
            DbgridTraslados.Columns(3).Caption = "Documento"
            DbgridTraslados.Columns(4).Caption = "Bodega Salida"
            DbgridTraslados.Columns(5).Caption = "Bodega Entrada"
            DbgridTraslados.Columns(6).Caption = "Salida"
            DbgridTraslados.Columns(7).Caption = "De +"
            DbgridTraslados.Columns(8).Caption = "De -"
            DbgridTraslados.Columns(9).Caption = "Des.Pro"
            DbgridTraslados.Columns(10).Caption = "Des.Prov"
            DbgridTraslados.Columns(11).Caption = "Trasladado"
            
            DbgridTraslados.Columns(6).NumberFormat = "#,###,##0.00"
            DbgridTraslados.Columns(6).Alignment = dbgRight
            DbgridTraslados.Columns(7).NumberFormat = "#,###,##0.00"
            DbgridTraslados.Columns(7).Alignment = dbgRight
            DbgridTraslados.Columns(8).NumberFormat = "#,###,##0.00"
            DbgridTraslados.Columns(8).Alignment = dbgRight
            DbgridTraslados.Columns(9).NumberFormat = "#,###,##0.00"
            DbgridTraslados.Columns(9).Alignment = dbgRight
            DbgridTraslados.Columns(10).NumberFormat = "#,###,##0.00"
            DbgridTraslados.Columns(10).Alignment = dbgRight
            DbgridTraslados.Columns(11).NumberFormat = "#,###,##0.00"
            DbgridTraslados.Columns(11).Alignment = dbgRight
            
'SALIDAS ____________________________________________________________________________________________________________
            
        'BUSCA LA LOS DESPACHOS DE LA TARIMA
        Set RSalidas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RSalidas, "Select ED.Documento, ED.Fecha, ED.NumeroDocumento, D.Descripcion, C.Descripcion, DD.Cantidad From DetalleSalidasInventario DD, EncabezadoSalidasInventario ED, Documentos D, Clientes C Where DD.FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And DD.Tarima = " & TxtTexto.Item(1).Text & " And DD.Linea = '" & TxtTexto.Item(2).Text & "' And DD.FichaTecnica = '" & TxtTexto.Item(0).Text & "' AND DD.Documento = ED.Documento And ED.TipoDeDocumento = D.CodigoDocumento And ED.Cliente = C.CodigoCliente")
            Else 'ORACLE
                Call Abrir_Recordset(RSalidas, "Select ED.Documento, ED.Fecha, ED.NumeroDocumento, D.Descripcion, C.Descripcion, DD.Cantidad From DetalleSalidasInventario DD, EncabezadoSalidasInventario ED, Documentos D, Clientes C Where DD.FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And DD.Tarima = " & TxtTexto.Item(1).Text & " And UPPER(DD.Linea) = '" & UCase(TxtTexto.Item(2).Text) & "' And UPPER(DD.FichaTecnica) = '" & UCase(TxtTexto.Item(0).Text) & "' AND DD.Documento = ED.Documento And UPPER(ED.TipoDeDocumento) = UPPER(D.CodigoDocumento) And UPPER(ED.Cliente) = UPPER(C.CodigoCliente)")
            End If
        
            
                Set DbGridDespachos.DataSource = RSalidas
                
                DbGridDespachos.Columns(0).Width = "1000"
                DbGridDespachos.Columns(1).Width = "1000"
                DbGridDespachos.Columns(2).Width = "1000"
                DbGridDespachos.Columns(3).Width = "1500"
                DbGridDespachos.Columns(4).Width = "3000"
                DbGridDespachos.Columns(5).Width = "1000"
                
                DbGridDespachos.Columns(0).Caption = "Transaccion"
                DbGridDespachos.Columns(1).Caption = "Fecha"
                DbGridDespachos.Columns(2).Caption = "# Documento"
                DbGridDespachos.Columns(3).Caption = "Documento"
                DbGridDespachos.Columns(4).Caption = "Cliente"
                DbGridDespachos.Columns(5).Caption = "Cantidad"
                
                DbGridDespachos.Columns(5).NumberFormat = "#,###,##0.00"
                DbGridDespachos.Columns(5).Alignment = dbgRight
            
        
'CIERRE TARIMA_________________________________________________________________________________________________________
        Set RCierreTarima = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RCierreTarima, "Select Fecha, Existencia, CantidadMas, CantidadMenos, ExistenciaNueva, CantidadProcesada, DesperdicioProceso, DesperdicioProveedor, (CantidadProcesada + DesperdicioProveedor), Usuario From CierreBulto Where FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTexto.Item(1).Text & " And LineaProduccion = '" & TxtTexto.Item(2).Text & "' And FichaTecnica = '" & TxtTexto.Item(0).Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RCierreTarima, "Select Fecha, Existencia, CantidadMas, CantidadMenos, ExistenciaNueva, CantidadProcesada, DesperdicioProceso, DesperdicioProveedor, (CantidadProcesada + DesperdicioProveedor), Usuario From CierreBulto Where FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTexto.Item(1).Text & " And UPPER(LineaProduccion) = '" & UCase(TxtTexto.Item(2).Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtTexto.Item(0).Text) & "'")
            End If
        
                If RCierreTarima.RecordCount > 0 Then
                Set DbGridCierreTarima.DataSource = RCierreTarima
                
                    DbGridCierreTarima.Columns(0).Caption = "Fecha"
                    DbGridCierreTarima.Columns(1).Caption = "Existencia"
                    DbGridCierreTarima.Columns(2).Caption = "De +"
                    DbGridCierreTarima.Columns(3).Caption = "De -"
                    DbGridCierreTarima.Columns(4).Caption = "Exi.Nueva"
                    DbGridCierreTarima.Columns(5).Caption = "Procesado"
                    DbGridCierreTarima.Columns(6).Caption = "Des.Proc."
                    DbGridCierreTarima.Columns(7).Caption = "Des.Prov."
                    DbGridCierreTarima.Columns(8).Caption = "Descargado"
                    DbGridCierreTarima.Columns(9).Caption = "Usuario"
                    
                    
                    DbGridCierreTarima.Columns(0).Width = "1000"
                    DbGridCierreTarima.Columns(1).Width = "1000"
                    DbGridCierreTarima.Columns(2).Width = "1000"
                    DbGridCierreTarima.Columns(3).Width = "1000"
                    DbGridCierreTarima.Columns(4).Width = "1000"
                    DbGridCierreTarima.Columns(5).Width = "1000"
                    DbGridCierreTarima.Columns(6).Width = "1000"
                    DbGridCierreTarima.Columns(7).Width = "1000"
                    DbGridCierreTarima.Columns(8).Width = "1200"
                    DbGridCierreTarima.Columns(9).Width = "1000"
                    
                
                    DbGridCierreTarima.Columns(1).NumberFormat = "#,###,##0.00"
                    DbGridCierreTarima.Columns(2).NumberFormat = "#,###,##0.00"
                    DbGridCierreTarima.Columns(3).NumberFormat = "#,###,##0.00"
                    DbGridCierreTarima.Columns(4).NumberFormat = "#,###,##0.00"
                    DbGridCierreTarima.Columns(5).NumberFormat = "#,###,##0.00"
                    DbGridCierreTarima.Columns(6).NumberFormat = "#,###,##0.00"
                    DbGridCierreTarima.Columns(7).NumberFormat = "#,###,##0.00"
                    DbGridCierreTarima.Columns(8).NumberFormat = "#,###,##0.00"
            End If
                    
                    
'CONSUMOS_________________________________________________________________________________________________________
        Set RConsumos = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RConsumos, "Select D.Documento, E.Fecha, D.Orden, D.Desperdicio, D.Cantidad From DetalleConsumoMateriaPrima D, EncabezadoCapturaParos E Where D.Fecha = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And D.Tarima = " & TxtTexto.Item(1).Text & " And D.Linea = '" & TxtTexto.Item(2).Text & "' And D.FichaTecnica = '" & TxtTexto.Item(0).Text & "' And D.Documento = E.Documento")
            Else 'ORACLE
                Call Abrir_Recordset(RConsumos, "Select D.Documento, E.Fecha, D.Orden, D.Desperdicio, D.Cantidad From DetalleConsumoMateriaPrima D, EncabezadoCapturaParos E Where D.Fecha = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " D.And Tarima = " & TxtTexto.Item(1).Text & " And UPPER(D.Linea) = '" & UCase(TxtTexto.Item(2).Text) & "' And UPPER(D.FichaTecnica) = '" & UCase(TxtTexto.Item(0).Text) & "' And D.Documento = E.Documento")
            End If
        
                
                Set DataGridConsumos.DataSource = RConsumos
                
                    DataGridConsumos.Columns(0).Caption = "Documento"
                    DataGridConsumos.Columns(1).Caption = "Fecha"
                    DataGridConsumos.Columns(2).Caption = "Orden"
                    DataGridConsumos.Columns(3).Caption = "Desperdicio"
                    DataGridConsumos.Columns(4).Caption = "Cantidad"
                    
                    
                    DbGridCierreTarima.Columns(0).Width = "1000"
                    DbGridCierreTarima.Columns(1).Width = "1000"
                    DbGridCierreTarima.Columns(2).Width = "1000"
                    DbGridCierreTarima.Columns(3).Width = "1000"
                    DbGridCierreTarima.Columns(4).Width = "1000"
                    
                    DbGridCierreTarima.Columns(3).NumberFormat = "#,###,##0.00"
                    DbGridCierreTarima.Columns(4).NumberFormat = "#,###,##0.00"
                    
                    

'BUSCA LOS AJUSTES _________________________________________________________________________
        Set RAjustes = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RAjustes, "Select Fecha, Efecto, Cantidad, Observaciones, Usuario From AjustesProductoTerminado Where FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Tarima = " & TxtTexto.Item(1).Text & " And Linea = '" & TxtTexto.Item(2).Text & "' And FichaTecnica = '" & TxtTexto.Item(0).Text & "'")
            Else 'ORACLE
                    Call Abrir_Recordset(RAjustes, "Select Fecha, Efecto, Cantidad, Observaciones, Usuario From AjustesProductoTerminado Where FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And Tarima = " & TxtTexto.Item(1).Text & " And UPPER(Linea) = '" & UCase(TxtTexto.Item(2).Text) & "' And UPPER(FichaTecnica) = '" & UCase(TxtTexto.Item(0).Text) & "'")
            End If
                    Set DbGridAjustes.DataSource = RAjustes
                    
                    DbGridAjustes.Columns(0).Width = "1000"
                    DbGridAjustes.Columns(1).Width = "800"
                    DbGridAjustes.Columns(2).Width = "1200"
                    DbGridAjustes.Columns(3).Width = "4000"
                    DbGridAjustes.Columns(4).Width = "1000"
                    
                    DbGridAjustes.Columns(0).Caption = "Fecha"
                    DbGridAjustes.Columns(1).Caption = "Efecto"
                    DbGridAjustes.Columns(2).Caption = "Cantidad"
                    DbGridAjustes.Columns(2).NumberFormat = "#,###,##0.00"
                    DbGridAjustes.Columns(3).Caption = "Observaciones"
                    DbGridAjustes.Columns(4).Caption = "Usuario"

                

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub DbGridAjustes_HeadClick(ByVal ColIndex As Integer)
            RAjustes.Sort = RAjustes.Fields(ColIndex).Name
End Sub

Private Sub DbGridCierreTarima_HeadClick(ByVal ColIndex As Integer)
            RCierreTarima.Sort = RCierreTarima.Fields(ColIndex).Name
End Sub

Private Sub DbGridDespachos_HeadClick(ByVal ColIndex As Integer)
            RSalidas.Sort = RSalidas.Fields(ColIndex).Name
End Sub

Private Sub DbgridTraslados_HeadClick(ByVal ColIndex As Integer)
            RTraslados.Sort = RTraslados.Fields(ColIndex).Name
End Sub

Private Sub MskFecPro_Change()
        'LIMPIA TODOS LOS TEXT CUANDO VAYA A DIGITAR ALGUN DATO
                TxtTexto.Item(3).Text = ""
                TxtTexto.Item(4).Text = ""
                TxtTexto.Item(5).Text = ""
                TxtTexto.Item(8).Text = ""
                TxtTexto.Item(13).Text = ""
                TxtTexto.Item(15).Text = ""

        
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

Private Sub TxtTexto_Change(Index As Integer)
        'BUSCA FICHA TECNICA
        If Index = 0 Then
            Set RBuscaFicha = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTexto.Item(0).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                End If
                If RBuscaFicha.RecordCount > 0 Then
                        lblFicTec.Caption = RBuscaFicha!Descrip
                Else
                        lblFicTec.Caption = ""
                End If
        'BUSCA LINEA
        ElseIf Index = 2 Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(2).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(2).Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                        LblLin.Caption = RBuscaLinea!Descrip
                Else
                        LblLin.Caption = ""
                End If
        End If
        
        
        'LIMPIA TODOS LOS TEXT CUANDO VAYA A DIGITAR ALGUN DATO
        If (Index = 0 Or Index = 1 Or Index = 2) Then
                TxtTexto.Item(3).Text = ""
                TxtTexto.Item(4).Text = ""
                TxtTexto.Item(5).Text = ""
                TxtTexto.Item(8).Text = ""
                TxtTexto.Item(13).Text = ""
                TxtTexto.Item(15).Text = ""
        End If
                
        'BODEGA
        If Index = 8 Then
                Set RBuscaBodega = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtTexto.Item(8).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtTexto.Item(8).Text) & "'")
                    End If
                    If RBuscaBodega.RecordCount > 0 Then
                        LblBod.Caption = RBuscaBodega!Descripcion
                    Else
                        LblBod.Caption = ""
                    End If
        End If
        
        If Index = 10 Then
                Set RBuscaProveedor = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtTexto.Item(10).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaProveedor, "Select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtTexto.Item(10).Text) & "'")
                    End If
                    If RBuscaProveedor.RecordCount > 0 Then
                        TxtTexto.Item(11).Text = RBuscaProveedor!Descripcion
                    Else
                        TxtTexto.Item(11).Text = ""
                    End If
        End If
        
        If Index = 12 Then
                Set RBuscaLinea2 = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaLinea2, "Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(12).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaLinea2, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTexto.Item(12).Text) & "'")
                    End If
                    If RBuscaLinea2.RecordCount > 0 Then
                        TxtTexto.Item(14).Text = RBuscaLinea2!Descrip
                    Else
                        TxtTexto.Item(14).Text = ""
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

Private Sub Txttexto_LostFocus(Index As Integer)

    If IsNumeric(TxtTexto.Item(1).Text) Then
        If Index = 1 Then
            If MskFecPro.Text = "" Then
                    Set RBuscaTarima = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaTarima, "Select FechaProduccion From DetalleEntradasInventario Where Tarima = " & TxtTexto.Item(1).Text & " And FichaTecnica = '" & TxtTexto.Item(0).Text & "' And Linea = '" & TxtTexto.Item(2).Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaTarima, "Select FechaProduccion From DetalleEntradasInventario Where Tarima = " & TxtTexto.Item(1).Text & " And UPPER(FichaTecnica) = '" & UCase(TxtTexto.Item(0).Text) & "' And Linea = '" & TxtTexto.Item(2).Text & "'")
                            End If
                        
                        If RBuscaTarima.RecordCount > 0 Then
                                MskFecPro.Text = RBuscaTarima!FechaProduccion
                        Else
                            MsgBox "Ficha Tecnica Con Este Bulto y Linea No Existe", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
            End If
        End If
    End If
        
End Sub
