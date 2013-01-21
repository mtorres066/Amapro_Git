VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Lineas 
   BackColor       =   &H000000FF&
   Caption         =   "LINEAS De Produccion"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11910
   Icon            =   "Lineas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11910
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
      Height          =   8295
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   7095
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   11295
         _ExtentX        =   19923
         _ExtentY        =   12515
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
         Picture         =   "Lineas.frx":0E42
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   9
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   3975
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   7
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
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
   End
   Begin MSDataGridLib.DataGrid DbGridBultos 
      Height          =   2535
      Left            =   3600
      TabIndex        =   14
      Top             =   1440
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      TabAcrossSplits =   -1  'True
      TabAction       =   2
      WrapCellPointer =   -1  'True
      FormatLocked    =   -1  'True
      AllowDelete     =   -1  'True
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
      ColumnCount     =   7
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
      BeginProperty Column02 
         DataField       =   "FechaProduccion"
         Caption         =   "Fecha Produccion"
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
      BeginProperty Column03 
         DataField       =   "LineaProduccion"
         Caption         =   "LineaProduccion"
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
         DataField       =   "CodigoMateriaPrima"
         Caption         =   "Materia Prima"
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
      BeginProperty Column06 
         DataField       =   "Bulto"
         Caption         =   "Bulto"
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
            Locked          =   -1  'True
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1140.095
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   450.142
         EndProperty
         BeginProperty Column04 
            Locked          =   -1  'True
            ColumnWidth     =   1110.047
         EndProperty
         BeginProperty Column05 
            Locked          =   -1  'True
            ColumnWidth     =   2264.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   540.284
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameLineas 
      Caption         =   "Datos de la Linea"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   11715
      Begin VB.ComboBox CboGra 
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
         Height          =   315
         ItemData        =   "Lineas.frx":2EB4
         Left            =   6720
         List            =   "Lineas.frx":2EBE
         TabIndex        =   19
         Text            =   "NO"
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox TxtPla 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   18
         Top             =   4080
         Width           =   1572
      End
      Begin VB.TextBox TxtCodLin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
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
         Left            =   1800
         MaxLength       =   2
         TabIndex        =   28
         ToolTipText     =   " "
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtDesLin 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   27
         ToolTipText     =   " "
         Top             =   720
         Width           =   7875
      End
      Begin VB.TextBox TxtEspTec 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   26
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtTar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CheckBox ChkAct 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5160
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtVel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   1800
         TabIndex        =   23
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox TxtOrden 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   22
         Top             =   2280
         Width           =   1575
      End
      Begin VB.TextBox TxtGru 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Height          =   285
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   21
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton CmdAgregar2 
         Caption         =   "Agregar Materias Primas"
         Height          =   375
         Left            =   1440
         TabIndex        =   20
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox TxtUni 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   17
         Top             =   3720
         Width           =   1572
      End
      Begin VB.TextBox TxtTipPro 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   1800
         MaxLength       =   15
         TabIndex        =   16
         Top             =   3360
         Width           =   1572
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Incluye en Grafica Reporte Ejecutivo"
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
         Left            =   3480
         TabIndex        =   41
         Top             =   4080
         Width           =   3165
      End
      Begin VB.Label Label1 
         Caption         =   "Planta"
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
         Index           =   9
         Left            =   240
         TabIndex        =   40
         Top             =   4080
         Width           =   1455
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
         Left            =   240
         TabIndex        =   39
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Descipcion"
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
         Left            =   240
         TabIndex        =   38
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   37
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   "Estado Linea"
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
         Left            =   3960
         TabIndex        =   36
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   4
         Left            =   240
         TabIndex        =   35
         Top             =   1920
         Width           =   732
      End
      Begin VB.Label LblFicha 
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
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3480
         TabIndex        =   34
         Top             =   1080
         Width           =   6135
      End
      Begin VB.Label Label1 
         Caption         =   "Velocidad"
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
         Index           =   6
         Left            =   240
         TabIndex        =   33
         Top             =   2640
         Width           =   1332
      End
      Begin VB.Label Label1 
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
         Height          =   252
         Index           =   5
         Left            =   240
         TabIndex        =   32
         Top             =   2280
         Width           =   972
      End
      Begin VB.Image ImageVerde 
         Height          =   240
         Left            =   3480
         Picture         =   "Lineas.frx":2ECA
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image ImageRoja 
         Height          =   240
         Left            =   3480
         Picture         =   "Lineas.frx":4F3C
         Top             =   360
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Grupo"
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
         Left            =   240
         TabIndex        =   31
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Unidad Medida"
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
         Index           =   7
         Left            =   240
         TabIndex        =   30
         Top             =   3720
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo De Producto"
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
         Left            =   240
         TabIndex        =   29
         Top             =   3360
         Width           =   1470
      End
   End
   Begin MSDataGridLib.DataGrid DbGrid1 
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      ToolTipText     =   "click en el lado izquierdo de la fila para seleccionar"
      Top             =   5040
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   5953
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
      ColumnCount     =   12
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
      BeginProperty Column03 
         DataField       =   "Activa"
         Caption         =   "Activa"
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
         DataField       =   "Velocidad"
         Caption         =   "Velocidad"
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
      BeginProperty Column07 
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
      BeginProperty Column08 
         DataField       =   "UnidadMedida"
         Caption         =   "Unidad Medida"
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
         DataField       =   "TipoProducto"
         Caption         =   "Tipo Producto"
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
      BeginProperty Column11 
         DataField       =   "Planta"
         Caption         =   "Planta"
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
            ColumnWidth     =   494.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2700.284
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1244.976
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   315.213
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   404.787
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   464.882
         EndProperty
         BeginProperty Column06 
            ColumnWidth     =   1230.236
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   824.882
         EndProperty
         BeginProperty Column08 
            ColumnWidth     =   1154.835
         EndProperty
         BeginProperty Column09 
            ColumnWidth     =   1170.142
         EndProperty
         BeginProperty Column10 
            ColumnWidth     =   900.284
         EndProperty
         BeginProperty Column11 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   465
      Left            =   10320
      MouseIcon       =   "Lineas.frx":6FAE
      Picture         =   "Lineas.frx":73F0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4560
      Width           =   1515
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   465
      Left            =   8280
      MouseIcon       =   "Lineas.frx":9462
      Picture         =   "Lineas.frx":98A4
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4560
      Width           =   2000
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   465
      Left            =   6240
      MouseIcon       =   "Lineas.frx":9DD6
      Picture         =   "Lineas.frx":A218
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   2000
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   468
      Left            =   4200
      MouseIcon       =   "Lineas.frx":A74A
      Picture         =   "Lineas.frx":AB8C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   2000
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   468
      Left            =   2160
      MouseIcon       =   "Lineas.frx":B0BE
      Picture         =   "Lineas.frx":B500
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4560
      Width           =   2000
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   468
      Left            =   120
      MouseIcon       =   "Lineas.frx":BA32
      Picture         =   "Lineas.frx":BE74
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4560
      Width           =   2000
   End
End
Attribute VB_Name = "Lineas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim BEditar As Boolean

Dim RLineas As New ADODB.Recordset
Dim RBuscaLineas As New ADODB.Recordset
Dim RBuscaMateriaPrima As New ADODB.Recordset
Dim RLineasBultos As New ADODB.Recordset
Dim RBuscaBultos As New ADODB.Recordset
Dim RBuscaBultos2 As New ADODB.Recordset
Dim RBuscaOrden As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VTexto As String


Sub botones()
    If Bandera = True Then
         FrameLineas.Enabled = True
         
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
         TxtCodLin.SetFocus
         
         DbGrid1.Visible = False
    Else
         FrameLineas.Enabled = False
         
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         
         DbGrid1.Visible = True
    End If
End Sub

Private Sub ChkAct_Click()
    If ChkAct.Value = 1 Then
        ChkAct.Caption = "Activa"
        ImageVerde.Visible = True
        ImageRoja.Visible = False
    Else
        ChkAct.Caption = "Inactiva"
        ImageVerde.Visible = False
        ImageRoja.Visible = True
        
    End If
End Sub


Private Sub ChkAct_Validate(Cancel As Boolean)
    If ChkAct.Value = 1 Then
        ChkAct.Caption = "Activa"
        ImageVerde.Visible = True
        ImageRoja.Visible = False
    Else
        ChkAct.Caption = "Inactiva"
        ImageVerde.Visible = False
        ImageRoja.Visible = True
        
    End If

End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
        
        If GEditar = True Then
        Else
            MsgBox "Usted No Esta Autorizado", vbOKOnly + vbInformation
            Exit Sub
        End If
        
        Bandera = True
        botones
        Limpia_Campos
        TxtCodLin.Enabled = True
        TxtCodLin.SetFocus
        BEditar = False
        CboGra.Text = "NO"
End Sub


Private Sub CmdAgregar2_Click()
    On Error Resume Next
        
    'BUSCA QUE MATERIAS PRIMAS TIENE ASIGNADA LA FICHA TECNICA
    Set RBuscaBultos = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaBultos, "Select FM.CodigoMateriaPrima, FT.Descrip from FichaTecnicaConMateriaPrima FM, FichaTecnica FT where FM.CodigoMateriaPrima = FT.Esp_Tec And FM.Esp_Tec = '" & TxtEspTec.Text & "'")
        Else
            Call Abrir_Recordset(RBuscaBultos, "Select FM.CodigoMateriaPrima, FT.Descrip from FichaTecnicaConMateriaPrima FM, FichaTecnica FT where UPPER(FM.CodigoMateriaPrima) = UPPER(FT.Esp_Tec) And UPPER(FM.Esp_Tec) = '" & UCase(TxtEspTec.Text) & "'")
        End If
    If RBuscaBultos.RecordCount > 0 Then
        Do Until RBuscaBultos.EOF
                
                VTexto = "'" & TxtCodLin.Text & "', '" 'LINEA
                VTexto = VTexto & TxtEspTec.Text & "', '" 'FICHA TECNICA
                VTexto = VTexto & RBuscaBultos(0) & "', '" 'CODIGO MATERIA PRIMA
                VTexto = VTexto & RBuscaBultos(1) & "',  " 'DESCRIPCION
                VTexto = VTexto & "0" & ", " 'BULTO O TARIMA
                If GOrigenDeDatos = "AmaproAccess" Then
                    VTexto = VTexto & "#" & Format(Date, "mm/dd/yyyy") & "#, '" 'FECHA PRODUCCION
                Else
                    VTexto = VTexto & "To_Date('" & Format(Date, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", '" 'FECHA PRODUCCION
                End If
                VTexto = VTexto & "77" & "'" 'LINEA PRODUCCION
                
                Conexion.Execute "Insert Into LineasBultos Values(" & VTexto & ")"
                
                
                'SI YA EXISTE LA MATERIA PRIMA
                If GOrigenDeDatos = "AmaproAcces" Then
                        If Err.Number = -2147467259 Then
                            'SI SE DUPLICA ESTA BIEN, NO HAY PROBLEMA
                        ElseIf (Err.Number <> -2147467259 And Err.Number <> 0) Then
                            MsgBox "Error " & Err.Number & " " & Err.Description
                        End If
                Else
                        If Err.Number = -2147217873 Then
                            'SI SE DUPLICA ESTA BIEN, NO HAY PROBLEMA
                        ElseIf (Err.Number <> -2147217873 And Err.Number <> 0) Then
                            MsgBox "Error " & Err.Number & " " & Err.Description
                        End If
                End If
            RBuscaBultos.MoveNext
        Loop
    End If
    
    'BUSCA LAS MATERIAS PRIMAS QUE USA LA FICHA TECNICA Y LOS BULTOS QUE TIENE ASIGNADOS
    Set RBuscaBultos2 = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaBultos2, "Select * from LineasBultos where Linea = '" & TxtCodLin.Text & "' And Esp_tec = '" & TxtEspTec.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaBultos2, "Select * from LineasBultos where UPPER(Linea) = '" & UCase(TxtCodLin.Text) & "' And UPPER(Esp_tec) = '" & UCase(TxtEspTec.Text) & "'")
        End If
    
    Set DbGridBultos.DataSource = RBuscaBultos2

End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
        If GBorrar = False Then
               MsgBox "Usted No Tiene Acceso a Esta Funcion", vbOKOnly + vbInformation, "Informacion"
        Else
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        RLineas.Delete
                        
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
                        RLineas.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RLineas.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If

        End If
            
End Sub


Private Sub CmdCancelar_Click()
        Bandera = False
        botones
        Llena_Campos
        TxtCodLin.Enabled = True
        TxtDesLin.Enabled = True
        
End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
        
        Bandera = True
        botones
        BEditar = True
        
        
        
        If GEditar = True Then
            TxtCodLin.Enabled = False
            TxtDesLin.Enabled = False
            TxtVel.Enabled = True
            TxtGru.Enabled = True
            TxtTipPro.Enabled = True
            TxtUni.Enabled = True
        Else
            TxtCodLin.Enabled = False
            TxtDesLin.Enabled = False
            TxtVel.Enabled = False
            TxtGru.Enabled = False
            TxtTipPro.Enabled = False
            TxtUni.Enabled = False
            TxtEspTec.SetFocus
        End If
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   'VALIDA EL CODIGO DE LINEA
   If TxtCodLin.Text = "" Then
        MsgBox "Codigo Linea No Puede Estar En Blanco", vbOKOnly + vbInformation, "Informacion"
        TxtCodLin.SetFocus
        Exit Sub
    End If
   
    'VALIDA LA DESCRIPCION DE LA LINEA
    If TxtDesLin.Text = "" Then
        MsgBox "Descripcion Linea No Puede Estar En Blanco", vbOKOnly + vbInformation, "Informacion"
        TxtDesLin.SetFocus
        Exit Sub
    End If
   
   'VALIDA LA VELOCIDAD
   If Not IsNumeric(TxtVel.Text) Then
        MsgBox "La Velocidad Debe Ser Numerica", vbOKOnly + vbInformation, "Informacion"
        TxtCodLin.SetFocus
        Exit Sub
   End If
      
   'VALIDA LA TARIMA
   If Not IsNumeric(TxtTar.Text) Then
        MsgBox "La Tarima Debe Ser Numerica", vbOKOnly + vbInformation, "Informacion"
        TxtTar.SetFocus
        Exit Sub
   End If
   
   'VALIDA LA orden
   If TxtOrden <> "" Then
        Set RBuscaOrden = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where Documento = '" & TxtOrden.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaOrden, "Select * From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtOrden.Text) & "'")
            End If
            If RBuscaOrden.RecordCount > 0 Then
            Else
               MsgBox "Numero De Orden No Existe", vbOKOnly + vbInformation, "Informacion"
               Exit Sub
            End If
   Else
        TxtOrden.Text = "0"
   End If
           
                   
   
   'AGREGAR
                    If BEditar = False Then
                            VTexto = "'" & TxtCodLin.Text & "', '" 'CODIGO
                            VTexto = VTexto & TxtDesLin.Text & "', '" 'DESCRIPCION
                            VTexto = VTexto & TxtEspTec.Text & "', " 'FICHA TECNICA
                            If ChkAct.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'IMPRIME CALIDAD
                            Else
                                VTexto = VTexto & "0" & ", " 'IMPRIME CALIDAD
                            End If
                            VTexto = VTexto & TxtTar.Text & ", " 'TARIMA
                            VTexto = VTexto & TxtVel.Text & ", '" 'VELOCIDAD
                            VTexto = VTexto & TxtOrden.Text & "', '" 'ORDEN
                            VTexto = VTexto & TxtGru.Text & "', '" 'GRUPO
                            VTexto = VTexto & GUsuario & "', '" 'USUARIO
                            VTexto = VTexto & TxtUni.Text & "', '" 'UNIDAD DE MEDIDA
                            VTexto = VTexto & TxtTipPro.Text & "', '" 'TIPO DE PRODUCTO
                            VTexto = VTexto & TxtPla.Text & "', '" 'PLANTA
                            VTexto = VTexto & CboGra.Text & "'" 'INCLUYE EN GRAFICA
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into Lineas Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            'VTexto = "Linea = '" & TxtCodLin.Text & "', " 'CODIGO
                            'VTexto = "Descrip = '" & TxtDesLin.Text & "', " 'DESCRIPCION
                            VTexto = "Esp_Tec = '" & TxtEspTec.Text & "', " 'FICHA TECNICA
                            If ChkAct.Value = "1" Then
                                VTexto = VTexto & "Activa = -1, " 'ACTIVA
                            Else
                                VTexto = VTexto & "Activa = 0, " 'ACTIVA
                            End If
                            VTexto = VTexto & "Tarima = " & TxtTar.Text & ", " 'TARIMA
                            VTexto = VTexto & "Velocidad = " & TxtVel.Text & ", " 'VELOCIDAD
                            VTexto = VTexto & "Orden = '" & TxtOrden.Text & "', " 'ORDEN
                            VTexto = VTexto & "Grupo = '" & TxtGru.Text & "', " 'GRUPO
                            VTexto = VTexto & "Usuario = '" & GUsuario & "', " 'USUARIO
                            VTexto = VTexto & "UnidadMedida = '" & TxtUni.Text & "', " 'UNIDAD DE MEDIDA
                            VTexto = VTexto & "TipoProducto = '" & TxtTipPro.Text & "', " 'TIPO DE PRODUCTO
                            VTexto = VTexto & "Planta = '" & TxtPla.Text & "', " '
                            VTexto = VTexto & "IncluyeEnGraficaReporteEjecutivo = '" & CboGra.Text & "'" '
                            
                            
                            VTexto = VTexto & " Where Linea = '" & TxtCodLin.Text & "'" 'LINEA
                            
                            Conexion.Execute "UPDATE Lineas SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                      'SI ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        
                      'SI ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        botones
                        CmdGrabar.SetFocus
                        
                        'HABILITA LA LLAVE
                        TxtCodLin.Enabled = True
                                                
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RLineas.Requery
                        RLineas.MoveLast
                        Llena_Campos
        
        
                         'HABILITA EL GRID PARA MODIFICACIONES
                         DbGridBultos.AllowUpdate = False
                         DbGridBultos.AllowDelete = True
                         
                         TxtCodLin.Enabled = True
                         TxtDesLin.Enabled = True
        
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub DbGrid1_BeforeUpdate(Cancel As Integer)
On Error Resume Next
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        End If

End Sub


Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
                RLineas.Sort = RLineas.Fields(ColIndex).Name
End Sub




Private Sub DbGrid1_SelChange(Cancel As Integer)
                Llena_Campos
End Sub

Private Sub DbGridBultos_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
        'TARIMA
        If (ColIndex = 6 Or ColIndex = 3 Or ColIndex = 2) Then
            
            If Not IsNumeric(DbGridBultos.Columns(6).Text) Then
                MsgBox "Bulto Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
            End If
            If Not IsDate(DbGridBultos.Columns(2).Text) Then
                MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
            End If
            
            RBuscaBultos2.MoveNext
            If Err <> 0 Then
            End If
        End If

End Sub

Private Sub DbGridBultos_BeforeUpdate(Cancel As Integer)
On Error Resume Next
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        End If
End Sub


Private Sub DBGridBusqueda_DblClick()
        TxtEspTec.Text = DbGridBusqueda.Columns(0)
        TxtEspTec.SetFocus
        FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            TxtEspTec.Text = DbGridBusqueda.Columns(0)
            TxtEspTec.SetFocus
            FrameBusqueda.Visible = False
        End If
End Sub


Private Sub Form_Activate()
    'BUSCA LAS MATERIAS PRIMAS QUE USA LA FICHA TECNICA Y LOS BULTOS QUE TIENE ASIGNADOS
    Set RBuscaBultos2 = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaBultos2, "Select * from LineasBultos where Linea = '" & TxtCodLin.Text & "' And Esp_tec = '" & TxtEspTec.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaBultos2, "Select * from LineasBultos where UPPER(Linea) = '" & UCase(TxtCodLin.Text) & "' And UPPER(Esp_tec) = '" & UCase(TxtEspTec.Text) & "'")
        End If
    
    Set DbGridBultos.DataSource = RBuscaBultos2
    
End Sub

Private Sub Form_Load()
        Set RLineas = New ADODB.Recordset
        Call Abrir_Recordset(RLineas, "Select * From Lineas")
        Set DbGrid1.DataSource = RLineas
        Llena_Campos
    
        If GEditar = True Then
                DbGrid1.AllowUpdate = True
        Else
                DbGrid1.AllowUpdate = False
        End If
    
End Sub

Private Sub TxtBusqueda_Change()
                    Set RBusqueda = New ADODB.Recordset
            
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(1).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Descrip Like '%" & TxtBusqueda.Text & "%' And Activa = -1 And TipoInventario = 'PRODUCTO TERMINADO'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%' And Activa = -1 And TipoInventario = 'PRODUCTO TERMINADO'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Esp_Tec Like '%" & TxtBusqueda.Text & "%' And Activa = -1 And TipoInventario = 'PRODUCTO TERMINADO'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%' And Activa = -1 And TipoInventario = 'PRODUCTO TERMINADO'")
                            End If
                            
                    End If
                            
                            Set DbGridBusqueda.DataSource = RBusqueda
                            DbGridBusqueda.Columns(0).Width = "1200"
                            DbGridBusqueda.Columns(1).Width = "5000"
                            DbGridBusqueda.Columns(2).Width = "1300"
                            DbGridBusqueda.Columns(3).Width = "800"


End Sub

Private Sub TxtCodLin_Change()
    'BUSCA LAS MATERIAS PRIMAS QUE USA LA FICHA TECNICA Y LOS BULTOS QUE TIENE ASIGNADOS
    Set RBuscaBultos2 = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaBultos2, "Select * from LineasBultos where Linea = '" & TxtCodLin.Text & "' And Esp_tec = '" & TxtEspTec.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaBultos2, "Select * from LineasBultos where UPPER(Linea) = '" & UCase(TxtCodLin.Text) & "' And UPPER(Esp_tec) = '" & UCase(TxtEspTec.Text) & "'")
        End If
    
    Set DbGridBultos.DataSource = RBuscaBultos2
End Sub

Private Sub TxtCodLin_GotFocus()
    TxtCodLin.SelStart = 0
    TxtCodLin.SelLength = Len(TxtCodLin.Text)
End Sub

Private Sub TxtDesLin_GotFocus()
    TxtDesLin.SelStart = 0
    TxtDesLin.SelLength = Len(TxtDesLin.Text)
End Sub

Private Sub TxtEspTec_Change()
On Error Resume Next
    'BUSCA LA DESCRIPCION DE LA FICHA TECNICA
    Set RBuscaLineas = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaLineas, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtEspTec.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaLineas, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtEspTec.Text) & "'")
        End If
        If RBuscaLineas.RecordCount > 0 Then
            LblFicha.Caption = RBuscaLineas!Descrip
        Else
            LblFicha.Caption = ""
        End If
        
    'BUSCA LAS MATERIAS PRIMAS QUE USA LA FICHA TECNICA Y LOS BULTOS QUE TIENE ASIGNADOS
    Set RBuscaBultos2 = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaBultos2, "Select * from LineasBultos where Linea = '" & TxtCodLin.Text & "' And Esp_tec = '" & TxtEspTec.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaBultos2, "Select * from LineasBultos where UPPER(Linea) = '" & UCase(TxtCodLin.Text) & "' And UPPER(Esp_tec) = '" & UCase(TxtEspTec.Text) & "'")
        End If
    
    Set DbGridBultos.DataSource = RBuscaBultos2
    

End Sub

Private Sub TxtEspTec_DblClick()
        Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial, Envases from FichaTecnica Where Activa = -1 And TipoInventario = 'PRODUCTO TERMINADO'")
            Set DbGridBusqueda.DataSource = RBusqueda
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
            DbGridBusqueda.Columns(0).Width = "1200"
            DbGridBusqueda.Columns(1).Width = "4000"
End Sub

Private Sub TxtEspTec_GotFocus()
        TxtEspTec.SelStart = 0
        TxtEspTec.SelLength = Len(TxtEspTec.Text)
End Sub

Private Sub Txtesptec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

        If KeyAscii = 43 Then
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial, Envases from FichaTecnica Where Activa = -1 And TipoInventario = 'PRODUCTO TERMINADO'")
            Set DbGridBusqueda.DataSource = RBusqueda
            FrameBusqueda.Visible = True
            TxtBusqueda.SetFocus
            DbGridBusqueda.Columns(0).Width = "1200"
            DbGridBusqueda.Columns(1).Width = "4000"
    End If

End Sub


Private Sub Txtcodlin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub txtDeslin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Private Sub TxtGru_GotFocus()
        TxtGru.SelStart = 0
        TxtGru.SelLength = Len(TxtGru.Text)
End Sub

Private Sub TxtGru_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtOrden_GotFocus()
        TxtOrden.SelStart = 0
        TxtOrden.SelLength = Len(TxtOrden.Text)
End Sub

Private Sub TxtOrden_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub Txttar_GotFocus()
        TxtTar.SelStart = 0
        TxtTar.SelLength = Len(TxtTar.Text)
End Sub
Private Sub chkact_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub Txttar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub


Private Sub TxtTipPro_GotFocus()
        TxtTipPro.SelStart = 0
        TxtTipPro.SelLength = Len(TxtTipPro.Text)
End Sub

Private Sub TxtTipPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Private Sub TxtUni_GotFocus()
        TxtUni.SelStart = 0
        TxtUni.SelLength = Len(TxtUni.Text)
End Sub

Private Sub TxtUni_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Private Sub TxtVel_GotFocus()
        TxtVel.SelStart = 0
        TxtVel.SelLength = Len(TxtVel.Text)
End Sub

Private Sub TxtVel_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Public Sub Llena_Campos()
On Error Resume Next

    If RLineas.RecordCount > 0 Then
        If IsNull(RLineas!Linea) Then
            TxtCodLin.Text = ""
        Else
            TxtCodLin.Text = RLineas!Linea
        End If
        
        If IsNull(RLineas!Descrip) Then
            TxtDesLin.Text = ""
        Else
            TxtDesLin.Text = RLineas!Descrip
        End If
        
        If IsNull(RLineas!Esp_Tec) Then
            TxtEspTec.Text = ""
        Else
            TxtEspTec.Text = RLineas!Esp_Tec
        End If
        
        If GOrigenDeDatos = "AmaproAccess" Then
                If RLineas!Activa = "Verdadero" Then
                    ChkAct.Value = "1"
                Else
                    ChkAct.Value = "0"
                End If
        Else
                If RLineas!Activa = "-1" Then
                    ChkAct.Value = "1"
                Else
                    ChkAct.Value = "0"
                End If
        End If
        
        If IsNull(RLineas!Tarima) Then
            TxtTar.Text = ""
        Else
            TxtTar.Text = RLineas!Tarima
        End If
        
        If IsNull(RLineas!Velocidad) Then
            TxtVel.Text = ""
        Else
            TxtVel.Text = RLineas!Velocidad
        End If
        
        If IsNull(RLineas!Orden) Then
            TxtOrden.Text = ""
        Else
            TxtOrden.Text = RLineas!Orden
        End If
        
        If IsNull(RLineas!Grupo) Then
            TxtGru.Text = ""
        Else
            TxtGru.Text = RLineas!Grupo
        End If
        
        If IsNull(RLineas!unidadMedida) Then
            TxtUni.Text = ""
        Else
            TxtUni.Text = RLineas!unidadMedida
        End If
        
        If IsNull(RLineas!TipoProducto) Then
            TxtTipPro.Text = ""
        Else
            TxtTipPro.Text = RLineas!TipoProducto
        End If
        If IsNull(RLineas!Planta) Then
            TxtPla.Text = ""
        Else
            TxtPla.Text = RLineas!Planta
        End If
        If IsNull(RLineas!IncluyeEnGraficaReporteEjecutivo) Then
            CboGra.Text = ""
        Else
            CboGra.Text = RLineas!IncluyeEnGraficaReporteEjecutivo
        End If
        
            
    Else
            Limpia_Campos
    End If
        
        If Err <> 0 Then
        End If
End Sub

Public Sub Limpia_Campos()
        
            TxtCodLin.Text = ""
            TxtDesLin.Text = ""
            TxtEspTec.Text = ""
            ChkAct.Value = "0"
            TxtTar.Text = "0"
            TxtVel.Text = "0"
            TxtOrden.Text = "0"
            TxtGru.Text = ""
            TxtUni.Text = ""
            TxtTipPro.Text = ""
            TxtPla.Text = ""
           
End Sub


