VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form CuadreBultoMateriaPrima 
   BackColor       =   &H000000FF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cuadre De Bulto Materia Prima"
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9480
   Icon            =   "CuadreBultoMateriaPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   9480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8055
      Left            =   4200
      TabIndex        =   39
      Top             =   7800
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   8640
         Picture         =   "CuadreBultoMateriaPrima.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   240
         Width           =   735
      End
      Begin VB.Data DataConsultas 
         Caption         =   "Defectos"
         Connect         =   "Access"
         DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "CuadreBultoMateriaPrima.frx":074C
         Height          =   7695
         Left            =   120
         OleObjectBlob   =   "CuadreBultoMateriaPrima.frx":0768
         TabIndex        =   40
         Top             =   240
         Width           =   9255
      End
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   555
      Left            =   1680
      Picture         =   "CuadreBultoMateriaPrima.frx":1143
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1500
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   1920
      Top             =   6960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin TabDlg.SSTab TabIngresos 
      Height          =   6975
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "CuadreBultoMateriaPrima.frx":1675
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameNumerosIngresos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "CuadreBultoMateriaPrima.frx":198F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridNumerosIngresos"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "CuadreBultoMateriaPrima.frx":1DE1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(8)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(9)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "LblBusCodMatPri"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "FrameOpciones"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "TxtBuscar"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "DtpFecIni"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "DtpFecFin"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "CmdBuscar"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "CmdActualizar"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "TxtBusCodMatPri"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).ControlCount=   11
      Begin VB.TextBox TxtBusCodMatPri 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67680
         TabIndex        =   23
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Actualizar"
         Height          =   735
         Left            =   -69720
         Picture         =   "CuadreBultoMateriaPrima.frx":2233
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Datos"
         Height          =   735
         Left            =   -67680
         Picture         =   "CuadreBultoMateriaPrima.frx":253D
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   5760
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -67680
         TabIndex        =   21
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   37301
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -67680
         TabIndex        =   20
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24641539
         CurrentDate     =   37301
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -67680
         TabIndex        =   22
         ToolTipText     =   " "
         Top             =   3960
         Width           =   1935
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
         ForeColor       =   &H00FF0000&
         Height          =   1935
         Left            =   -74880
         TabIndex        =   29
         Top             =   720
         Width           =   5205
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha Y Materia Prima"
            Height          =   1305
            Index           =   2
            Left            =   3480
            Picture         =   "CuadreBultoMateriaPrima.frx":297F
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Numero Ingreso y Codigo Materia Prima"
            Height          =   1305
            Index           =   0
            Left            =   120
            Picture         =   "CuadreBultoMateriaPrima.frx":3249
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   " "
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha"
            Height          =   1305
            Index           =   1
            Left            =   1800
            Picture         =   "CuadreBultoMateriaPrima.frx":3553
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame FrameNumerosIngresos 
         Caption         =   "Datos De Cuadre Bulto"
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
         Height          =   4935
         Left            =   120
         TabIndex        =   27
         Top             =   1200
         Width           =   9195
         Begin VB.TextBox TxtTur 
            Appearance      =   0  'Flat
            DataField       =   "Turno"
            DataSource      =   "DataCuadre"
            Height          =   285
            Left            =   2160
            MaxLength       =   1
            TabIndex        =   2
            Top             =   1080
            Width           =   615
         End
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            DataField       =   "Linea"
            DataSource      =   "DataCuadre"
            Height          =   285
            Left            =   2160
            MaxLength       =   2
            TabIndex        =   3
            Top             =   1440
            Width           =   1815
         End
         Begin MSMask.MaskEdBox MskNueExi 
            DataField       =   "ExistenciaNueva"
            DataSource      =   "DataCuadre"
            Height          =   285
            Left            =   2160
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   3600
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskDifMen 
            DataField       =   "DiferenciaMenos"
            DataSource      =   "DataCuadre"
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Top             =   3240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   255
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskDifMas 
            DataField       =   "DiferenciaMas"
            DataSource      =   "DataCuadre"
            Height          =   285
            Left            =   2160
            TabIndex        =   7
            Top             =   2880
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   32768
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskExi 
            DataField       =   "Existencia"
            DataSource      =   "DataCuadre"
            Height          =   285
            Left            =   2160
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   2520
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataCuadre"
            Height          =   285
            Left            =   2160
            TabIndex        =   0
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "UsuarioEditar"
            DataSource      =   "DataCuadre"
            Height          =   285
            Index           =   1
            Left            =   7560
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   4560
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "UsuarioAgregar"
            DataSource      =   "DataCuadre"
            Height          =   285
            Index           =   0
            Left            =   7560
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   36
            TabStop         =   0   'False
            Top             =   4200
            Width           =   1455
         End
         Begin MSMask.MaskEdBox MskHor 
            DataField       =   "Hora"
            DataSource      =   "DataCuadre"
            Height          =   285
            Left            =   2160
            TabIndex        =   1
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Enabled         =   0   'False
            MaxLength       =   5
            Format          =   "HH:mm"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtCod 
            Appearance      =   0  'Flat
            DataField       =   "CodigoMateriaPrima"
            DataSource      =   "DataCuadre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2160
            MaxLength       =   15
            TabIndex        =   4
            Top             =   1800
            Width           =   1815
         End
         Begin VB.TextBox TxtNumIng 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "NumeroIngreso"
            DataSource      =   "DataCuadre"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   5
            ToolTipText     =   " "
            Top             =   2160
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
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
            Index           =   11
            Left            =   120
            TabIndex        =   52
            Top             =   1080
            Width           =   510
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario Editar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   10
            Left            =   6240
            TabIndex        =   51
            Top             =   4560
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   7
            Left            =   6240
            TabIndex        =   50
            Top             =   4200
            Width           =   1140
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
            Left            =   4080
            TabIndex        =   49
            Top             =   1440
            Width           =   4935
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   48
            Top             =   1440
            Width           =   480
         End
         Begin VB.Label LblUnidadMedida 
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
            Left            =   4080
            TabIndex        =   47
            Top             =   2520
            Width           =   2655
         End
         Begin VB.Label LblMateriaPrima 
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
            Left            =   4080
            TabIndex        =   46
            Top             =   1800
            Width           =   4935
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Nueva Existencia"
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
            Index           =   5
            Left            =   120
            TabIndex        =   45
            Top             =   3600
            Width           =   1500
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Diferencia Menos"
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
            Index           =   3
            Left            =   120
            TabIndex        =   44
            Top             =   3240
            Width           =   1500
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Existencia"
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
            TabIndex        =   43
            Top             =   2520
            Width           =   885
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Diferencia Mas"
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
            Height          =   195
            Index           =   16
            Left            =   120
            TabIndex        =   42
            Top             =   2880
            Width           =   1290
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Codigo Materia Prima"
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
            Index           =   2
            Left            =   120
            TabIndex        =   32
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Numero Ingreso"
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
            Index           =   1
            Left            =   120
            TabIndex        =   31
            Top             =   2160
            Width           =   1350
         End
         Begin VB.Label Label2 
            Caption         =   "Hora"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   975
         End
      End
      Begin MSDBGrid.DBGrid DBGridNumerosIngresos 
         Bindings        =   "CuadreBultoMateriaPrima.frx":385D
         Height          =   6105
         Left            =   -74880
         OleObjectBlob   =   "CuadreBultoMateriaPrima.frx":3876
         TabIndex        =   16
         Top             =   720
         Width           =   9225
      End
      Begin VB.Label LblBusCodMatPri 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Materia Prima"
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
         Left            =   -69600
         TabIndex        =   38
         Top             =   4320
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Inicial"
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
         Left            =   -68880
         TabIndex        =   35
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
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
         Index           =   8
         Left            =   -68880
         TabIndex        =   34
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Lbletiqueta 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero Ingreso"
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
         Left            =   -70560
         TabIndex        =   30
         Top             =   3960
         Width           =   2775
      End
   End
   Begin VB.Data DataCuadre 
      BackColor       =   &H80000014&
      Caption         =   "Cuadre De Bulto Materia Prima"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "CuadreBultoMateriaPrima"
      Top             =   7680
      Width           =   9315
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   555
      Left            =   7920
      MouseIcon       =   "CuadreBultoMateriaPrima.frx":51AF
      Picture         =   "CuadreBultoMateriaPrima.frx":55F1
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   555
      Left            =   6360
      MouseIcon       =   "CuadreBultoMateriaPrima.frx":5A33
      Picture         =   "CuadreBultoMateriaPrima.frx":5E75
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   4800
      MouseIcon       =   "CuadreBultoMateriaPrima.frx":63A7
      Picture         =   "CuadreBultoMateriaPrima.frx":67E9
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3240
      MouseIcon       =   "CuadreBultoMateriaPrima.frx":6D1B
      Picture         =   "CuadreBultoMateriaPrima.frx":715D
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   555
      Left            =   120
      MouseIcon       =   "CuadreBultoMateriaPrima.frx":768F
      Picture         =   "CuadreBultoMateriaPrima.frx":7AD1
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7080
      Width           =   1500
   End
End
Attribute VB_Name = "CuadreBultoMateriaPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim Mensaje As String
Dim Buscar As String

Dim RBuscaNumeroIngreso As Recordset
Dim RBuscaNumeroIngresoEntradas As Recordset
Dim RBuscaMateriaPrima As Recordset
Dim RBuscaLinea As Recordset

Dim VCodigoMateriaPrima As String
Dim VNumeroIngreso As Double
Dim VNuevaExistencia As Double

Dim BNumeroIngreso As Boolean
Dim BMateriaPrima As Boolean
Dim BLinea As Boolean


Sub botones()
    If Bandera = True Then
         FrameNumerosIngresos.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdEditar.Enabled = False
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataCuadre.Visible = False
         FrameOpciones.Visible = False
         DBGridNumerosIngresos.Visible = False
    Else
         FrameNumerosIngresos.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataCuadre.Visible = True
         FrameOpciones.Visible = True
         DBGridNumerosIngresos.Visible = True
    End If
End Sub

Private Sub CmdActualizar_Click()
    DataCuadre.RecordSource = "Select * From CuadreBultoMateriaPrima"
    DataCuadre.Refresh
    DBGridNumerosIngresos.Refresh
    TabIngresos.Tab = 1
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
       'AGREGA UN REGISTRO
        DataCuadre.Recordset.AddNew
        'SI HAY ERRORES
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        
        'BOTONES DISPONIBLES
        Bandera = True
        botones
     
        'ASIGNA FECHA Y HORA
        MskFec.Text = Date
        MskHor.Text = Time
        
        If MskHor.Text >= "07:00" And MskHor.Text <= "19:00" Then
            TxtTur.Text = "D"
        Else
            TxtTur.Text = "N"
        End If
        
        'ASIGNA USUARIO AGREGAR
        Text1.Item(0).Text = GUsuario
        TxtLin.SetFocus
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next
            
            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If Mensaje = vbOK Then
   
                         VCodigoMateriaPrima = TxtCod.Text
                         VNumeroIngreso = TxtNumIng.Text
                            
                        'GRABA DATOS
                        DataCuadre.Recordset.Delete
                        
                        If Err <> 0 Then
                             MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Else
                              Bandera = False
                              botones
                              CmdAgregar.SetFocus
                        End If

                
                DataCuadre.Recordset.MoveLast
            End If
  
            If DataCuadre.Recordset.EOF Then
                DataCuadre.Recordset.MoveLast
                If Err = 3021 Then
                    Mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdBuscar_Click()
        'NUMERO INGRESO Y CODIGO MATERIA PRIMA
        If OptBusqueda.Item(0).Value = True Then
            If Not IsNumeric(TxtBuscar.Text) Then
                    MsgBox "El Numero De Ingreso Debe Ser Numerico", vbOKOnly + vbExclamation, "Informacion"
                    TxtBuscar.SetFocus
                    Exit Sub
            End If
            DataCuadre.RecordSource = "Select * From CuadreBultoMateriaPrima Where NumeroIngreso = " & TxtBuscar.Text & " And CodigoMateriaPrima = '" & TxtBusCodMatPri.Text & "'"
            TabIngresos.Tab = 0
        'FECHAS
        ElseIf OptBusqueda.Item(1).Value = True Then
            DataCuadre.RecordSource = "Select * From CuadreBultoMateriaPrima Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#"
            TabIngresos.Tab = 1
        'FECHA Y MATERIA PRIMA
        ElseIf OptBusqueda.Item(2).Value = True Then
            DataCuadre.RecordSource = "Select * From CuadreBultoMateriaPrima Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And CodigoMateriaPrima = '" & TxtBuscar.Text & "'"
            TabIngresos.Tab = 1
        End If
        DataCuadre.Refresh
        DBGridNumerosIngresos.Refresh
        
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
        
        DataCuadre.Recordset.CancelUpdate
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        
        Bandera = False
        botones

End Sub

Private Sub CmdEditar_Click()
On Error Resume Next
       'AGREGA UN REGISTRO
        DataCuadre.Recordset.Edit
        'SI HAY ERRORES
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        
        'BOTONES DISPONIBLES
        Bandera = True
        botones
        
        'ASIGNA USUARIO EDITAR
        Text1.Item(1).Text = GUsuario
        TxtNumIng.SetFocus
End Sub


Private Sub CmdGrabar_Click()
   On Error Resume Next
          
    'REVISA LA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
    End If
          
    'VALIDA LOS TURNOS
    If TxtTur.Text <> "D" And TxtTur.Text <> "N" Then
        MsgBox "Turno Incorrecto", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
          
    'REVISA EL CODIGO DE LA MATERIA PRIMA
    If TxtCod.Text = "" Then
        MsgBox "Codigo De Materia Prima No Puede Estar Vacio", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
        
    'REVISA SI EXISTE EL BULTO
    Set RBuscaNumeroIngreso = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCod.Text & "' And Disponibilidad = 'EXISTE'")
        If RBuscaNumeroIngreso.RecordCount > 0 Then
        Else
            MsgBox "Este Bulto Ya Fue Cerrado O No Existe", vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
    
    'ASIGNA VARIABLES PARA CALCULOS
    VNumeroIngreso = TxtNumIng.Text
    VCodigoMateriaPrima = TxtCod.Text
    VNuevaExistencia = MskNueExi.Text
           
   'GRABA DATOS
   DataCuadre.Recordset.Update
   
   If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
   Else
            'MODIFICA LA EXISTENCIA DISPONIBLE DEL BULTO
             Set RBuscaNumeroIngreso = Db.OpenRecordset("Select * From DetalleEntradasMateriaPrima Where NumeroIngreso = " & VNumeroIngreso & " And Codigo = '" & VCodigoMateriaPrima & "' And Disponibilidad = 'EXISTE'")
                 If RBuscaNumeroIngreso.RecordCount > 0 Then
                         RBuscaNumeroIngreso.Edit
                                 RBuscaNumeroIngreso!SaldoDisponibilidad = VNuevaExistencia
                         RBuscaNumeroIngreso.Update
                 End If
    
         Bandera = False
         botones
         CmdAgregar.SetFocus
   End If

   
End Sub

Private Sub CmdImprimir_Click()
On Error Resume Next

MousePointer = 11
'        Set RBuscaTotal = Db.OpenRecordset("Select ValorVenta From EncabezadoEntradas Where Documento = '" & TxtDocIng.Text & "'")
        'VLetras = numlet(CCur(RBuscaTotal(0)))
        'CrReportes.Formulas(0) = "letras = '" & VLetras & "'"
        
                CrReportes.SelectionFormula = "{CuadreBultoMateriaPrima.NumeroIngreso} = " & TxtNumIng.Text & " And {CuadreBultoMateriaPrima.Codigo} = '" & TxtCod.Text & "'"
                CrReportes.ReportFileName = App.Path & "\FormatoNumerosIngresosMPaProducir.rpt"
                CrReportes.Action = 1
                
                If Err <> 0 Then
                     MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                End If
                
MousePointer = 0

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
        FrameConsultas.Visible = False
End Sub



Private Sub DBGridConsultas_DblClick()
        If BNumeroIngreso = True Then
            TxtNumIng.Text = DBGridConsultas.Columns(1).Text
            TxtNumIng.SetFocus
        ElseIf BLinea = True Then
            TxtLin.Text = DBGridConsultas.Columns(0).Text
            TxtLin.SetFocus
        ElseIf BMateriaPrima = True Then
            TxtCod.Text = DBGridConsultas.Columns(0).Text
            TxtCod.SetFocus
        End If
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        If BNumeroIngreso = True Then
            TxtNumIng.Text = DBGridConsultas.Columns(1).Text
            TxtNumIng.SetFocus
        ElseIf BLinea = True Then
            TxtLin.Text = DBGridConsultas.Columns(0).Text
            TxtLin.SetFocus
        ElseIf BMateriaPrima = True Then
            TxtCod.Text = DBGridConsultas.Columns(0).Text
            TxtCod.SetFocus
        End If
        FrameConsultas.Visible = False
    End If
End Sub

Private Sub dbgridnumerosingresos_HeadClick(ByVal ColIndex As Integer)
    DataCuadre.RecordSource = ("Select * from CuadreBultoMateriaPrima order by " & DBGridNumerosIngresos.Columns(ColIndex).DataField)
    DataCuadre.Refresh
    DBGridNumerosIngresos.Refresh
    
End Sub

Private Sub Form_Load()
    DataCuadre.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
End Sub



Private Sub MskDifMas_GotFocus()
        MskDifMas.SelStart = 0
        MskDifMas.SelLength = Len(MskDifMas.Text)
End Sub

Private Sub MskDifMas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskDifMas_LostFocus()
    MskNueExi.Text = ((Val(MskExi.Text) + Val(MskDifMas.Text)) - Val(MskDifMen.Text))
End Sub

Private Sub MskDifMen_GotFocus()
        MskDifMen.SelStart = 0
        MskDifMen.SelLength = Len(MskDifMen.Text)
End Sub

Private Sub MskDifMen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskDifMen_LostFocus()
    MskNueExi.Text = ((Val(MskExi.Text) + Val(MskDifMas.Text)) - Val(MskDifMen.Text))
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

Private Sub MskHor_Change()

End Sub

Private Sub OptBusqueda_Click(Index As Integer)
        If Index = 0 Then
            Label2.Item(8).Visible = False
            Label2.Item(9).Visible = False
            Lbletiqueta.Caption = "Numero Ingreso"
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            DtpFecIni.Visible = False
            DtpFecFin.Visible = False
            LblBusCodMatPri.Visible = True
            TxtBusCodMatPri.Visible = True
        ElseIf Index = 1 Then
            Label2.Item(8).Visible = True
            Label2.Item(9).Visible = True
            Lbletiqueta.Caption = ""
            TxtBuscar.Visible = False
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            LblBusCodMatPri.Visible = False
            TxtBusCodMatPri.Visible = False
        ElseIf Index = 2 Then
            Label2.Item(8).Visible = True
            Label2.Item(9).Visible = True
            Lbletiqueta.Caption = "Codigo Materia Prima"
            TxtBuscar.Visible = True
            TxtBuscar.SetFocus
            DtpFecIni.Visible = True
            DtpFecFin.Visible = True
            LblBusCodMatPri.Visible = False
            TxtBusCodMatPri.Visible = False
        End If
End Sub

Private Sub TabIngresos_Click(PreviousTab As Integer)
        If TabIngresos.Tab = 2 Then
            DtpFecIni.Value = Date
            DtpFecFin.Value = Date
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

Private Sub TxtBusCodMatPri_GotFocus()
        TxtBusCodMatPri.SelStart = 0
        TxtBusCodMatPri.SelLength = Len(TxtBusCodMatPri.Text)
End Sub

Private Sub TxtBusCodMatPri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If

End Sub



Private Sub TxtCod_Change()
        Set RBuscaMateriaPrima = Db.OpenRecordset("Select Descripcion, UnidadMedida From CorrelativosMateriaPrima Where CodigoMateriaPrima = '" & TxtCod.Text & "'")
            If RBuscaMateriaPrima.RecordCount > 0 Then
                LblMateriaPrima.Caption = RBuscaMateriaPrima!Descripcion
                LblUnidadMedida.Caption = RBuscaMateriaPrima!UnidadMedida
            Else
                LblMateriaPrima.Caption = ""
                LblUnidadMedida.Caption = ""
            End If
End Sub

Private Sub TxtCod_DblClick()
        BNumeroIngreso = False
        BMateriaPrima = True
        BLinea = False
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataConsultas.RecordSource = "Select * From CorrelativosMateriaPrima"
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        FrameConsultas.Visible = True
        DBGridConsultas.SetFocus
End Sub

Private Sub TxtCod_GotFocus()
        TxtCod.SelStart = 0
        TxtCod.SelLength = Len(TxtCod.Text)
End Sub

Private Sub TxtCod_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BNumeroIngreso = False
            BMateriaPrima = True
            BLinea = False
            'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
            DataConsultas.RecordSource = "Select * From CorrelativosMateriaPrima"
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            DBGridConsultas.SetFocus
        End If
End Sub

Private Sub TxtLin_Change()
            Set RBuscaLinea = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtLin.Text & "'")
                If RBuscaLinea.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLinea!Descrip
                Else
                    LblLinea.Caption = ""
                End If
End Sub

Private Sub TxtLin_DblClick()
            BNumeroIngreso = False
            BMateriaPrima = False
            BLinea = True
            'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
            DataConsultas.RecordSource = "Select * From Lineas"
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            DBGridConsultas.SetFocus
End Sub

Private Sub TxtLin_GotFocus()
        TxtLin.SelStart = 0
        TxtLin.SelLength = Len(TxtLin.Text)
End Sub

Private Sub TxtLin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            BNumeroIngreso = False
            BMateriaPrima = False
            BLinea = True
            'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
            DataConsultas.RecordSource = "Select * From Lineas"
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            FrameConsultas.Visible = True
            DBGridConsultas.SetFocus
        End If
        
        
End Sub

Private Sub TxtNumIng_DblClick()
        BNumeroIngreso = True
        BMateriaPrima = False
        BLinea = False
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataConsultas.RecordSource = "Select BodegaDisponibilidad, NumeroIngreso, CantidadTraslado, CantidadSalida, SaldoDisponibilidad, UnidadMedida From DetalleEntradasMateriaPrima Where Codigo = '" & TxtCod.Text & "' And SaldoDisponibilidad > 0 Order By BodegaDisponibilidad, NumeroIngreso"
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        FrameConsultas.Visible = True
        DBGridConsultas.SetFocus
        columnas
End Sub

Private Sub TxtNumIng_GotFocus()
    TxtNumIng.SelStart = 0
    TxtNumIng.SelLength = Len(TxtNumIng.Text)
End Sub

Private Sub TxtNumIng_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
    
    If KeyAscii = 43 Then
        BNumeroIngreso = True
        BMateriaPrima = False
        BLinea = False
        'BUSCA EL INVENTARIO DE ACUERDO A LA BODEGA DE SALIDA
        DataConsultas.RecordSource = "Select BodegaDisponibilidad, NumeroIngreso, CantidadTraslado, CantidadSalida, SaldoDisponibilidad, UnidadMedida From DetalleEntradasMateriaPrima Where Codigo = '" & TxtCod.Text & "' And SaldoDisponibilidad > 0 Order By BodegaDisponibilidad, NumeroIngreso"
        DataConsultas.Refresh
        DBGridConsultas.Refresh
        FrameConsultas.Visible = True
        DBGridConsultas.SetFocus
        columnas
    End If
End Sub


Private Sub TxtNumIng_LostFocus()
            If IsNumeric(TxtNumIng.Text) Then
                'BUSCA EL NUMERO DE INGRESO Y ASIGNA LA BODEGA, CODIGO Y CANTIDAD DE ACUERDO COMO ENTRO A LA BODEGA
                Set RBuscaNumeroIngreso = Db.OpenRecordset("Select SaldoDisponibilidad From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCod.Text & "'")
                'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
                If RBuscaNumeroIngreso.RecordCount > 0 Then
                    MskExi.Text = RBuscaNumeroIngreso!SaldoDisponibilidad
                Else
                    MskExi.Text = "0"
                End If
            End If
End Sub

Sub columnas()
    DBGridConsultas.Columns(0).Caption = "Bodega"
    DBGridConsultas.Columns(0).Width = "500"
    DBGridConsultas.Columns(1).Caption = "# Ingreso"
    DBGridConsultas.Columns(1).Width = "1000"
    DBGridConsultas.Columns(2).Caption = "Inicio"
    DBGridConsultas.Columns(3).Caption = "Salidas"
    DBGridConsultas.Columns(4).Caption = "Existencia"
    DBGridConsultas.Columns(5).Caption = "Unidad Medida"
End Sub


Private Sub TxtTur_GotFocus()
        TxtTurno.SelStart = 0
        TxtTurno.SelLength = Len(TxtTurno.Text)
End Sub

Private Sub TxtTur_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub
