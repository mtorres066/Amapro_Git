VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form CapturaRutinas 
   BackColor       =   &H000000FF&
   Caption         =   "Captura Rutinas"
   ClientHeight    =   7920
   ClientLeft      =   1110
   ClientTop       =   345
   ClientWidth     =   11895
   Icon            =   "CapturaRutinas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7920
   ScaleMode       =   0  'User
   ScaleWidth      =   12000
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "&Borrar"
      Height          =   720
      Left            =   7080
      Picture         =   "CapturaRutinas.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   63
      Top             =   7080
      Width           =   2000
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   4
      Left            =   11400
      MouseIcon       =   "CapturaRutinas.frx":0A0A
      Picture         =   "CapturaRutinas.frx":0E4C
      Style           =   1  'Graphical
      TabIndex        =   62
      ToolTipText     =   "Ultimo Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   3
      Left            =   11040
      MouseIcon       =   "CapturaRutinas.frx":137E
      Picture         =   "CapturaRutinas.frx":17C0
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Siguiente Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   2
      Left            =   480
      MouseIcon       =   "CapturaRutinas.frx":1CF2
      Picture         =   "CapturaRutinas.frx":2134
      Style           =   1  'Graphical
      TabIndex        =   60
      ToolTipText     =   "Registro Anterior"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdBotones2 
      BackColor       =   &H00C0C0C0&
      Height          =   465
      Index           =   1
      Left            =   120
      MouseIcon       =   "CapturaRutinas.frx":2666
      Picture         =   "CapturaRutinas.frx":2AA8
      Style           =   1  'Graphical
      TabIndex        =   59
      ToolTipText     =   "Primer Registro"
      Top             =   7200
      Width           =   375
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   720
      Left            =   9120
      Picture         =   "CapturaRutinas.frx":2FDA
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7080
      Width           =   1875
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   720
      Left            =   5040
      Picture         =   "CapturaRutinas.frx":34F5
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7080
      Width           =   2000
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   720
      Left            =   3000
      Picture         =   "CapturaRutinas.frx":3A2C
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7080
      Width           =   2000
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   720
      Left            =   960
      Picture         =   "CapturaRutinas.frx":3F88
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   2000
   End
   Begin TabDlg.SSTab TabRutinas 
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   12091
      _Version        =   393216
      TabHeight       =   1058
      BackColor       =   255
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "CapturaRutinas.frx":4305
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameRutinas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FrameBuscar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CapturaRutinas.frx":461F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DbGridFichaTecnica"
      Tab(1).Control(1)=   "LblRutDes"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Seleccion O Busqueda"
      TabPicture(2)   =   "CapturaRutinas.frx":4A71
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "FrameBusqueda"
      Tab(2).ControlCount=   1
      Begin MSDataGridLib.DataGrid DbGridFichaTecnica 
         Height          =   5415
         Left            =   -74880
         TabIndex        =   58
         Top             =   1320
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   9551
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   19
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
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   10
         BeginProperty Column00 
            DataField       =   "Fec_Rut"
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
         BeginProperty Column02 
            DataField       =   "Hor_Rut"
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
            DataField       =   "Cabezal"
            Caption         =   "Cabezal"
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
            DataField       =   "Rutina"
            Caption         =   "Rutina"
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
         BeginProperty Column07 
            DataField       =   "Valor"
            Caption         =   "Valor"
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
            DataField       =   "Catalogo"
            Caption         =   "Catalogo"
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
            DataField       =   "llave"
            Caption         =   "llave"
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
               ColumnWidth     =   1319.811
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   510.236
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   750.047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   1395.213
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   764.787
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   4140.284
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnWidth     =   989.858
            EndProperty
            BeginProperty Column08 
               Object.Visible         =   0   'False
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column09 
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameBusqueda 
         Height          =   6015
         Left            =   -74880
         TabIndex        =   20
         Top             =   720
         Width           =   11415
         Begin VB.TextBox TxtBusRut 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            TabIndex        =   55
            Top             =   3240
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DtpFecFin 
            Height          =   255
            Left            =   9960
            TabIndex        =   27
            Top             =   2520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   87687171
            CurrentDate     =   37358
         End
         Begin MSComCtl2.DTPicker DtpFecIni 
            Height          =   255
            Left            =   8520
            TabIndex        =   26
            Top             =   2520
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   87687171
            CurrentDate     =   37358
         End
         Begin VB.Frame FrameOpcionesBusqueda 
            Caption         =   "Opciones De Busqueda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   480
            TabIndex        =   21
            Top             =   240
            Width           =   10575
            Begin VB.OptionButton OptOpcion 
               Caption         =   "Fechas Y Ficha Tecnica y Rutina"
               Height          =   1155
               Index           =   5
               Left            =   8880
               Picture         =   "CapturaRutinas.frx":4EC3
               Style           =   1  'Graphical
               TabIndex        =   56
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton OptOpcion 
               Caption         =   "Fechas Y Catalogo"
               Height          =   1155
               Index           =   4
               Left            =   7200
               Picture         =   "CapturaRutinas.frx":500D
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton OptOpcion 
               Caption         =   "Fechas Y Hora"
               Height          =   1155
               Index           =   2
               Left            =   3720
               Picture         =   "CapturaRutinas.frx":77AF
               Style           =   1  'Graphical
               TabIndex        =   24
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton OptOpcion 
               Caption         =   "Fechas"
               Height          =   1155
               Index           =   0
               Left            =   120
               Picture         =   "CapturaRutinas.frx":7AB9
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   360
               Value           =   -1  'True
               Width           =   1695
            End
            Begin VB.OptionButton OptOpcion 
               Caption         =   "Fechas Y Linea"
               Height          =   1155
               Index           =   1
               Left            =   1920
               Picture         =   "CapturaRutinas.frx":8383
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   360
               Width           =   1695
            End
            Begin VB.OptionButton OptOpcion 
               Caption         =   "Fechas Y Ficha Tecnica"
               Height          =   1155
               Index           =   3
               Left            =   5520
               Picture         =   "CapturaRutinas.frx":868D
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   360
               Width           =   1575
            End
         End
         Begin VB.TextBox TxtBusqueda 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   8520
            TabIndex        =   28
            Top             =   2880
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.CommandButton CmdBusSelDat 
            Caption         =   "Seleccionar Datos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   8520
            Picture         =   "CapturaRutinas.frx":87D7
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   3840
            Width           =   2775
         End
         Begin VB.CommandButton CmdBusAct 
            Caption         =   "Seleccionar Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   8520
            Picture         =   "CapturaRutinas.frx":8C19
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   4920
            Width           =   2775
         End
         Begin VB.Label LblBusRut 
            Caption         =   "Rutina"
            Height          =   255
            Left            =   7680
            TabIndex        =   57
            Top             =   3240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label lbl 
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
            Height          =   255
            Index           =   2
            Left            =   9960
            TabIndex        =   50
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label lbl 
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
            Height          =   255
            Index           =   0
            Left            =   8520
            TabIndex        =   49
            Top             =   2280
            Width           =   1095
         End
         Begin VB.Label LblBusqueda 
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
            Left            =   6240
            TabIndex        =   48
            Top             =   2880
            Width           =   2175
         End
      End
      Begin VB.Frame FrameBuscar 
         Height          =   2055
         Left            =   120
         TabIndex        =   10
         Top             =   4560
         Width           =   11415
         Begin MSMask.MaskEdBox TxtHorRut 
            Height          =   285
            Left            =   1800
            TabIndex        =   12
            Top             =   1440
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton CmdRefrescar 
            Caption         =   "Seleccionar Todos"
            Height          =   840
            Left            =   8400
            Picture         =   "CapturaRutinas.frx":8F23
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   960
            Width           =   2805
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "Seleccionar Rutinas"
            Height          =   855
            Left            =   5520
            Picture         =   "CapturaRutinas.frx":922D
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   960
            Width           =   2775
         End
         Begin VB.TextBox TxtLinRut 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   13
            Top             =   1440
            Width           =   495
         End
         Begin MSComCtl2.DTPicker PFecBus 
            Height          =   255
            Left            =   360
            TabIndex        =   11
            Top             =   1440
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            CustomFormat    =   "dd/MM/yyyy"
            Format          =   87687171
            CurrentDate     =   36915
         End
         Begin VB.Label Label3 
            BackColor       =   &H000080FF&
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
            Left            =   360
            TabIndex        =   47
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackColor       =   &H000080FF&
            Caption         =   "Hora"
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
            Left            =   1800
            TabIndex        =   46
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   2640
            TabIndex        =   45
            Top             =   1080
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H000080FF&
            BackStyle       =   0  'Transparent
            Caption         =   " Captura De Rutinas General"
            BeginProperty Font 
               Name            =   "Comic Sans MS"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   240
            TabIndex        =   44
            Top             =   360
            Width           =   3975
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000080FF&
            BackStyle       =   1  'Opaque
            Height          =   1695
            Left            =   120
            Top             =   240
            Width           =   11175
         End
      End
      Begin VB.Frame FrameRutinas 
         Caption         =   "Datos Generales De La Captura Rutina"
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
         Height          =   3495
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   11415
         Begin VB.TextBox TxtCat 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   10
            TabIndex        =   9
            Top             =   3000
            Width           =   1215
         End
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   2
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TxtFec 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   3
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TxtHor 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   5
            TabIndex        =   4
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TxtFicTec 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   15
            TabIndex        =   5
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TxtCab 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   6
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TxtRut 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            MaxLength       =   4
            TabIndex        =   7
            Top             =   2280
            Width           =   1215
         End
         Begin VB.TextBox TxtVal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   1680
            TabIndex        =   8
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label LblCatalogo 
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
            TabIndex        =   53
            Top             =   3000
            Width           =   8295
         End
         Begin VB.Label lblLabels 
            Caption         =   "Catalogo"
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
            TabIndex        =   52
            Top             =   3000
            Width           =   1815
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
            Height          =   285
            Left            =   3000
            TabIndex        =   51
            Top             =   480
            Width           =   8295
         End
         Begin VB.Label LblOrigen 
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
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   9600
            TabIndex        =   43
            Top             =   1560
            Width           =   1695
         End
         Begin VB.Label lbl 
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
            Height          =   255
            Index           =   18
            Left            =   240
            TabIndex        =   42
            Top             =   480
            Width           =   1215
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
            Height          =   255
            Index           =   19
            Left            =   240
            TabIndex        =   41
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label lblLabels 
            Caption         =   "Hora"
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
            Index           =   20
            Left            =   240
            TabIndex        =   40
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Ficha Tecnica"
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   39
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Cabezal"
            Height          =   255
            Index           =   22
            Left            =   240
            TabIndex        =   38
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Rutina"
            Height          =   255
            Index           =   23
            Left            =   240
            TabIndex        =   37
            Top             =   2280
            Width           =   1815
         End
         Begin VB.Label lblLabels 
            Caption         =   "Valor"
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
            Index           =   26
            Left            =   240
            TabIndex        =   36
            Top             =   2640
            Width           =   1815
         End
         Begin VB.Label LblFondo 
            Height          =   195
            Left            =   3000
            TabIndex        =   35
            Top             =   3120
            Width           =   2535
         End
         Begin VB.Label LblTapa 
            Height          =   315
            Left            =   3000
            TabIndex        =   34
            Top             =   2640
            Width           =   2535
         End
         Begin VB.Label LblInstrumento 
            Height          =   255
            Left            =   8640
            TabIndex        =   33
            Top             =   960
            Width           =   2775
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
            Height          =   285
            Left            =   3000
            TabIndex        =   32
            Top             =   1560
            Width           =   6495
         End
         Begin VB.Label LblRutina 
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
            Height          =   285
            Left            =   3000
            TabIndex        =   31
            Top             =   2280
            Width           =   8295
         End
      End
      Begin VB.Label LblRutDes 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -74880
         TabIndex        =   64
         Top             =   720
         Width           =   11295
      End
   End
End
Attribute VB_Name = "CapturaRutinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String
Dim VFecha As Date
Dim vtexto As String
Dim VMensaje As String

Dim RFicha As New ADODB.Recordset
Dim RRutina As New ADODB.Recordset
Dim RRutina2 As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaCatalogo As New ADODB.Recordset
Dim RRutinas As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim VEditar As Boolean
Dim RRut As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Sub botones()
    If Bandera = True Then
         FrameRutinas.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         
         CmdRefrescar.Enabled = False
         FrameBuscar.Visible = False
         FrameBusqueda.Visible = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdSalida.Enabled = False
        'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         Dbgridfichatecnica.Visible = False
    Else
         
         FrameRutinas.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdRefrescar.Enabled = True
         FrameBuscar.Visible = True
         FrameBusqueda.Visible = True
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         'BOTONES DE DATA
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         Dbgridfichatecnica.Visible = True
    End If
End Sub




Private Sub CmdAgregar_Click()
On Error Resume Next
        Bandera = True
        botones
        Limpia_Campos
        TxtLin.SetFocus
        TxtFec.Text = Date
        TxtHor.Text = Format(Time, "hh:mm")
        TxtVal.Text = "0"
        txtcab.Text = "0"
        
End Sub

Private Sub CmdBorrar_Click()
        On Error Resume Next
                    VMensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If VMensaje = vbOK Then
                        'BORRA EL REGISTRO
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Delete From CapturaRutinas where Fec_Rut = #" & Format(TxtFec.Text, "mm/dd/yyyy") & "# And Linea = '" & TxtLin.Text & "' And Esp_Tec = '" & TxtFicTec.Text & "' And Cabezal = " & txtcab.Text & " And Rutina = '" & TxtRut.Text & "' And Catalogo = '" & TxtCat.Text & "'"
                            Else 'ORACLE
                                Conexion.Execute "Delete From CapturaRutinas where Fec_Rut = To_Date('" & TxtFec.Text & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Cabezal = " & txtcab.Text & " And UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' And UPPER(Catalogo) = '" & UCase(TxtCat.Text) & "'"
                            End If
                        
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
                        RRutinas.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RRutinas.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_Campos
                    End If
           
            
End Sub


Private Sub CmdBotones2_Click(Index As Integer)
MousePointer = 11
    If Index = 1 Then
        RRutinas.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        RRutinas.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        RRutinas.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        RRutinas.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If RRutinas.BOF Then
        RRutinas.MoveFirst
    ElseIf RRutinas.EOF Then
        RRutinas.MoveLast
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_Campos
    
MousePointer = 0

End Sub

Private Sub CmdBusAct_Click()
'    vfecha = Date - 1
    Set RRutinas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut = #" & Format(Date, "mm/dd/yyyy") & "# And C.Rutina = R.Rutina")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut = To_Date('" & Date & "', 'dd/mm/yyyy') And C.Rutina = R.Rutina")
            End If
    Set Dbgridfichatecnica.DataSource = RRutinas
    TabRutinas.Tab = 1

End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
    VEditar = True
        
    'VERIFICA SI EXISTE LA RUTINA
    Set RRutinas = New ADODB.Recordset
    
            
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo from CapturaRutinas C, Rutinas R where C.Fec_Rut = #" & Format(PFecBus.Value, "mm/dd/yyyy") & "# and C.Hor_Rut = '" & TxtHorRut.Text & "' and C.Linea = '" & TxtLinRut.Text & "' And C.Rutina = R.Rutina Order By C.Rutina, C.Cabezal")
        Else 'ORACLE
            Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo from CapturaRutinas C, Rutinas R where C.Fec_Rut = To_date('" & PFecBus.Value & "', 'dd/mm/yyyy')" & " and UPPER(C.Hor_Rut) = '" & UCase(TxtHorRut.Text) & "' and UPPER(C.Linea) = '" & UCase(TxtLinRut.Text) & "' And C.Rutina = R.Rutina Order By C.Rutina, C.Cabezal")
        End If
        
        If RRutinas.RecordCount > 0 Then
        Else
            MsgBox "Rutinas No Existen", vbOKOnly + vbInformation, "Informacion"
            TxtHorRut.SetFocus
            Exit Sub
        End If
                    
    
    Set Dbgridfichatecnica.DataSource = RRutinas
            
    Bandera = True
    botones
    FrameRutinas.Visible = False
    Dbgridfichatecnica.BackColor = &H80000018
    
    Dbgridfichatecnica.Visible = True
    Dbgridfichatecnica.AllowUpdate = True
    Dbgridfichatecnica.Columns(0).Locked = True
    Dbgridfichatecnica.Columns(1).Locked = True
    Dbgridfichatecnica.Columns(2).Locked = True
    Dbgridfichatecnica.Columns(3).Locked = True
    Dbgridfichatecnica.Columns(4).Locked = True
    Dbgridfichatecnica.Columns(5).Locked = True
    Dbgridfichatecnica.Columns(6).Locked = True
    Dbgridfichatecnica.Columns(7).Locked = False
    Dbgridfichatecnica.Columns(8).Locked = True
    
    CmdCancelar.Enabled = False
    
    If Err <> 0 Then
        MsgBox "Error " & Err.Description, vbOKOnly + vbInformation, "Informacion"
    End If
    
    TabRutinas.Tab = 1
End Sub

Private Sub CmdBusSelDat_Click()
        Set RRutinas = New ADODB.Recordset
        
    If OptOpcion.Item(0).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# and C.Fec_Rut <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where Fec_Rut >= To_Date('" & DTPFecIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy') And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            End If
    'LINEA
    ElseIf OptOpcion.Item(1).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# and C.Fec_Rut <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And C.Linea = '" & TxtBusqueda.Text & "' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where Fec_Rut >= To_Date('" & DTPFecIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(TxtBusqueda.Text) & "' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            End If
    'HORA
    ElseIf OptOpcion.Item(2).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# and C.Fec_Rut <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And C.Hor_Rut = '" & TxtBusqueda.Text & "' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where Fec_Rut >= To_Date('" & DTPFecIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Hor_Rut) = '" & UCase(TxtBusqueda.Text) & "' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            End If
    'FICHA TECNICA
    ElseIf OptOpcion.Item(3).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# and C.Fec_Rut <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And C.Esp_Tec = '" & TxtBusqueda.Text & "' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where Fec_Rut >= To_Date('" & DTPFecIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Esp_Tec) = '" & UCase(TxtBusqueda.Text) & "' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            End If
    'CATALOGO
    ElseIf OptOpcion.Item(4).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# and C.Fec_Rut <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And C.Catalogo = '" & TxtBusqueda.Text & "' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where Fec_Rut >= To_Date('" & DTPFecIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Catalogo) = '" & UCase(TxtBusqueda.Text) & "' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            End If
    'RUTINA
    ElseIf OptOpcion.Item(5).Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= #" & Format(DTPFecIni.Value, "mm/dd/yyyy") & "# and C.Fec_Rut <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And C.Rutina = '" & TxtBusRut.Text & "' And C.Esp_Tec Like '" & TxtBusqueda.Text & "%' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where Fec_Rut >= To_Date('" & DTPFecIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & DtpFecFin.Value & "', 'dd/mm/yyyy')" & " And UPPER(Rutina) = '" & UCase(TxtBusqueda.Text) & "' And UPPER(Esp_Tec) Like '" & UCase(TxtBusqueda.Text) & "%' And C.Rutina = R.Rutina Order By C.Fec_Rut, C.Linea, C.Rutina, C.Cabezal")
            End If
    End If
    
    
    Set Dbgridfichatecnica.DataSource = RRutinas
    
    TabRutinas.Tab = 1
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next

    If VEditar = True Then
            Dbgridfichatecnica.AllowUpdate = False
    Else
            Bandera = False
            botones
            Llena_Campos
    End If
End Sub


Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   If VEditar = True Then
            Dbgridfichatecnica.AllowUpdate = False
            FrameRutinas.Visible = True
            Bandera = False
            botones
            Set RRutinas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= #" & Format((Date - 1), "mm/dd/yyyy") & "# And C.Fec_Rut <= #" & Format(Date, "mm/dd/yyyy") & "# And C.Rutina = R.Rutina")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= To_Date('" & (Date - 1) & "', 'dd/mm/yyyy') And C.Fec_Rut <= To_Date('" & Date & "', 'dd/mm/yyyy') And C.Rutina = R.Rutina")
            End If
        
            Set Dbgridfichatecnica.DataSource = RRutinas
            RRutinas.MoveLast
            
            If Err <> 0 Then
            End If
            
            Llena_Campos
            
   Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                vtexto = "Values(#" & Format(TxtFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                 vtexto = "Values(To_Date('" & Format(TxtFec.Text, "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            
                            vtexto = vtexto & "'" & TxtLin.Text & "', '" 'LINEA
                            vtexto = vtexto & TxtHor.Text & "', '" 'HOR
                            vtexto = vtexto & TxtFicTec.Text & "', " 'FICHA TECNICA
                            vtexto = vtexto & txtcab.Text & ", '" 'CABEZAL
                            vtexto = vtexto & TxtRut.Text & "', " 'RUTINA
                            vtexto = vtexto & TxtVal.Text & ", '" 'VALOR
                            vtexto = vtexto & TxtCat.Text & "')" 'CATALOGO
                            'REALIZA EL INSERT
                            Conexion.Execute "Insert Into CapturaRutinas " & vtexto
                     
                     
                     'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                            TxtFec.SetFocus
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Fecha, Linea, FichaTecnica, Rutina, Cabezal, Catalogo Ya Existe ", vbOKOnly + vbInformation, "Informacion"
                            TxtFec.SetFocus
                            Exit Sub
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                                                
                        Bandera = False
                        botones
                        CmdAgregar.SetFocus
                                                
                        'PARA QUE VUELVA A EJECUTAR EL RECORDSET ORIGINAL Y MUESTRE LOS DATOS GRABADOS
                        RRutinas.Requery
                        RRutinas.MoveLast
                        Llena_Campos
             
    End If
    
    Dbgridfichatecnica.BackColor = vbWhite
        
End Sub

Private Sub CmdRefrescar_Click()
    On Error Resume Next
        Set RRutinas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut = #" & Format(Date, "mm/dd/yyyy") & "#")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut = To_Date('" & Date & "', 'dd/mm/yyyy')")
            End If
        
        Set Dbgridfichatecnica.DataSource = RRutinas
        RRutinas.MoveLast
        If Err <> 0 Then
        End If
        Llena_Campos

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub




Private Sub DbGridFichaTecnica_AfterColEdit(ByVal ColIndex As Integer)
On Error Resume Next
        If ColIndex = 7 Then
            On Error Resume Next
            RRutinas.MoveNext
            If Err <> 0 Then
            End If
        End If
        
        If Err <> 0 Then
        MsgBox "Error " & Err.Description, vbOKOnly + vbInformation, "Informacion"
    End If
End Sub

Private Sub DBGridFichaTecnica_BeforeUpdate(Cancel As Integer)
    On Error Resume Next
        If Err <> 0 Then
            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        End If
End Sub


Private Sub DbGridFichaTecnica_DblClick()
On Error Resume Next
        Set RRutina2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RRutina2, "Select Descrip From Rutinas Where Rutina = '" & Dbgridfichatecnica.Columns(5).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RRutina2, "Select Descrip From Rutinas Where UPPER(Rutina) = '" & UCase(Dbgridfichatecnica.Columns(5).Text) & "'")
                End If
            If RRutina2.RecordCount > 0 Then
                LblRutDes.Caption = RRutina2(0)
            Else
                LblRutDes.Caption = ""
            End If
            
            If Err <> 0 Then
        MsgBox "Error " & Err.Description, vbOKOnly + vbInformation, "Informacion"
    End If
End Sub

Private Sub DbGridFichaTecnica_HeadClick(ByVal ColIndex As Integer)

        RRutinas.Sort = RRutinas.Fields(ColIndex).Name
End Sub


Private Sub DbGridFichaTecnica_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
            Set RRutina2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RRutina2, "Select Descrip From Rutinas Where Rutina = '" & Dbgridfichatecnica.Columns(5).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RRutina2, "Select Descrip From Rutinas Where UPPER(Rutina) = '" & UCase(Dbgridfichatecnica.Columns(5).Text) & "'")
                End If
                If Err <> 0 Then
                Else
                        If RRutina2.RecordCount > 0 Then
                            LblRutDes.Caption = RRutina2(0)
                        Else
                            LblRutDes.Caption = ""
                        End If
                End If
            

End Sub

Private Sub DbGridFichaTecnica_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
        Set RRutina2 = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RRutina2, "Select Descrip From Rutinas Where Rutina = '" & Dbgridfichatecnica.Columns(5).Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RRutina2, "Select Descrip From Rutinas Where UPPER(Rutina) = '" & UCase(Dbgridfichatecnica.Columns(5).Text) & "'")
                End If
            If Err <> 0 Then
            
            Else
                If RRutina2.RecordCount > 0 Then
                    LblRutDes.Caption = RRutina2(0)
                Else
                    LblRutDes.Caption = ""
                End If
            End If
            
End Sub

Private Sub Form_Load()
    On Error Resume Next
        Set RRutinas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= #" & Format((Date - 1), "mm/dd/yyyy") & "# And C.Fec_Rut <= #" & Format(Date, "mm/dd/yyyy") & "# And C.Rutina = R.Rutina")
            Else 'ORACLE
                Call Abrir_Recordset(RRutinas, "Select C.Fec_Rut, C.Linea, C.Hor_Rut, C.Esp_Tec, C.Cabezal, C.Rutina, R.Descrip, C.Valor, C.Catalogo From CapturaRutinas C, Rutinas R Where C.Fec_Rut >= To_Date('" & (Date - 1) & "', 'dd/mm/yyyy') And C.Fec_Rut <= To_Date('" & Date & "', 'dd/mm/yyyy') And C.Rutina = R.Rutina")
            End If
        
        Set Dbgridfichatecnica.DataSource = RRutinas
        RRutinas.MoveLast
        If Err <> 0 Then
            MsgBox "Error Al Ir Al Ultimo Registro" & Err.Description, vbOKOnly + vbInformation, "Informacion"
        End If
        Llena_Campos

    
    DTPFecIni.Value = Date
    DtpFecFin.Value = Date
    
    PFecBus.Value = Date
    
        
    
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        If Index = 0 Then
            TxtBusqueda.Visible = False
            LblBusqueda.Caption = ""
        ElseIf Index = 1 Then
            TxtBusqueda.Visible = True
            LblBusqueda.Caption = "Linea"
            TxtBusqueda.SetFocus
        ElseIf Index = 2 Then
            TxtBusqueda.Visible = True
            LblBusqueda.Caption = "Hora"
            TxtBusqueda.SetFocus
        ElseIf Index = 3 Then
            TxtBusqueda.Visible = True
            LblBusqueda.Caption = "Ficha Tecnica"
            TxtBusqueda.SetFocus
        ElseIf Index = 4 Then
            TxtBusqueda.Visible = True
            LblBusqueda.Caption = "Codigo Catalogo"
            TxtBusqueda.SetFocus
        ElseIf Index = 5 Then
            TxtBusqueda.Visible = True
            LblBusqueda.Caption = "Ficha Tecnica"
            TxtBusqueda.SetFocus
        End If
        
        If Index = 5 Then
            TxtBusRut.Visible = True
            LblBusRut.Visible = True
        Else
            TxtBusRut.Visible = False
            LblBusRut.Visible = False
        End If
            
        
        
End Sub

Private Sub PFecBus_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TabRutinas_Click(PreviousTab As Integer)
        If TabRutinas.Tab = 0 Then
            CmdBorrar.Enabled = True
            If CmdGrabar.Enabled = False Then
                Llena_Campos
            End If
        Else
            CmdBorrar.Enabled = False
        End If
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

Private Sub txtcab_GotFocus()
        txtcab.SelStart = 0
        txtcab.SelLength = Len(txtcab.Text)
End Sub

Private Sub txtcab_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
End Sub

Private Sub TxtCat_Change()
        Set RBuscaCatalogo = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaCatalogo, "Select DescripcionVariable From VariablesDescripcion Where CodigoVariable = '" & TxtCat.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaCatalogo, "Select DescripcionVariable From VariablesDescripcion Where UPPER(CodigoVariable) = '" & UCase(TxtCat.Text) & "'")
            End If
            If RBuscaCatalogo.RecordCount > 0 Then
                LblCatalogo.Caption = RBuscaCatalogo!DescripcionVariable
            Else
                LblCatalogo.Caption = ""
            End If
End Sub

Private Sub TxtCat_GotFocus()
        TxtCat.SelStart = 0
        TxtCat.SelLength = Len(TxtCat.Text)
End Sub

Private Sub TxtCat_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtFec_GotFocus()
        TxtFec.SelStart = 0
        TxtFec.SelLength = Len(TxtFec.Text)
End Sub

Private Sub TxtFec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtFicTec_Change()
    Set RFicha = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RFicha, "Select Origen ,Descrip from FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RFicha, "Select Origen ,Descrip from FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
        End If
    If RFicha.RecordCount > 0 Then
        If IsNull(RFicha!Origen) Then
            LblOrigen.Caption = ""
        Else
            LblOrigen.Caption = RFicha!Origen
        End If
            LblFichaTecnica.Caption = RFicha!Descrip
        
    Else
        LblOrigen.Caption = ""
        LblFichaTecnica.Caption = ""
    End If
End Sub

Private Sub TxtFicTec_GotFocus()
        TxtFicTec.SelStart = 0
        TxtFicTec.SelLength = Len(TxtFicTec.Text)
End Sub

Private Sub TxtFicTec_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtFicTec_LostFocus()
        Set RFicha = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RFicha, "Select Variables From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RFicha, "Select Variables From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            If RFicha.RecordCount > 0 Then
                    If IsNull(RFicha!Variables) Then
                        TxtCat.Text = ""
                    Else
                        TxtCat.Text = RFicha!Variables
                    End If
                        TxtCat.Text = RFicha!Variables
            End If
End Sub

Private Sub TxtRut_Change()
    Set RRutina = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RRutina, "Select Descrip From Rutinas Where Rutina = '" & TxtRut.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RRutina, "Select Descrip From Rutinas Where UPPER(Rutina) = '" & UCase(TxtRut.Text) & "'")
        End If
    If RRutina.RecordCount > 0 Then
        LblRutina.Caption = RRutina(0)
    Else
        LblRutina.Caption = ""
    End If
End Sub

Private Sub TxtRut_GotFocus()
        TxtRut.SelStart = 0
        TxtRut.SelLength = Len(TxtRut.Text)
        
End Sub

Private Sub TxtRut_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Private Sub TxtHor_GotFocus()
    TxtHor.SelStart = 0
    TxtHor.SelLength = Len(TxtHor.Text)
End Sub

Private Sub TxtHor_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub



Private Sub TxtHorRut_GotFocus()
        TxtHorRut.SelStart = 0
        TxtHorRut.SelLength = Len(TxtHorRut.Text)
End Sub

Private Sub TxtHorRut_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtLin_Change()
    
    Set RBuscaLinea = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLin.Text & "'")
        Else 'ORACLE
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

Private Sub TxtLin_LostFocus()
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Esp_Tec From Lineas Where Linea = '" & TxtLin.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Esp_Tec From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
                    TxtFicTec.Text = RBuscaLinea!Esp_Tec
                    Set RBuscaFicha = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaFicha, "Select Variables From FichaTecnica Where Esp_Tec = '" & RBuscaLinea!Esp_Tec & "'")
                        Else 'ORACLE
                            Call Abrir_Recordset(RBuscaFicha, "Select Variables From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(RBuscaLinea!Esp_Tec) & "'")
                        End If
                        If RBuscaFicha.RecordCount > 0 Then
                            If Not IsNull(RBuscaFicha!Variables) Then
                                TxtCat.Text = RBuscaFicha!Variables
                            End If
                        End If
            End If
End Sub

Private Sub TxtLinRut_GotFocus()
        TxtLinRut.SelStart = 0
        TxtLinRut.SelLength = Len(TxtLinRut.Text)
End Sub

Private Sub TxtLinRut_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtVal_GotFocus()
    TxtVal.SelStart = 0
    TxtVal.SelLength = Len(TxtVal.Text)
End Sub

Private Sub TxtVal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If

End Sub

Public Sub Llena_Campos()
On Error Resume Next
    If RRutinas.RecordCount > 0 Then
        'FECHA
            If IsNull(RRutinas!Fec_Rut) Then
                TxtFec.Text = ""
            Else
                TxtFec.Text = RRutinas!Fec_Rut
            End If
        'LINEA
            If IsNull(RRutinas!Linea) Then
                TxtLin.Text = ""
            Else
                TxtLin.Text = RRutinas!Linea
            End If
        'HORA
            If IsNull(RRutinas!Hor_rut) Then
                TxtHor.Text = ""
            Else
                TxtHor.Text = RRutinas!Hor_rut
            End If
        'FICHA
            If IsNull(RRutinas!Esp_Tec) Then
                TxtFicTec.Text = ""
            Else
                TxtFicTec.Text = RRutinas!Esp_Tec
            End If
        'CABEZAL
            If IsNull(RRutinas!cabezal) Then
                txtcab.Text = ""
            Else
                txtcab.Text = RRutinas!cabezal
            End If
        'RUTINA
            If IsNull(RRutinas!Rutina) Then
                TxtRut.Text = ""
            Else
                TxtRut.Text = RRutinas!Rutina
            End If
        'VALOR
            If IsNull(RRutinas!Valor) Then
                TxtVal.Text = ""
            Else
                TxtVal.Text = RRutinas!Valor
            End If
        'CATALOGO
            If IsNull(RRutinas!Catalogo) Then
                TxtCat.Text = ""
            Else
                TxtCat.Text = RRutinas!Catalogo
            End If
    Else
            Limpia_Campos
    End If
        
        If Err <> 0 Then
            
        End If

End Sub

Public Sub Limpia_Campos()
        
        TxtFec.Text = ""
        TxtLin.Text = ""
        TxtHor.Text = ""
        TxtFicTec.Text = ""
        txtcab.Text = 0
        TxtRut.Text = ""
        TxtVal.Text = 0
        TxtCat.Text = ""
        
End Sub





