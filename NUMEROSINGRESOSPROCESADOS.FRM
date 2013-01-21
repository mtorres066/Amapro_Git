VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form NumerosIngresosProcesados 
   BackColor       =   &H000000FF&
   Caption         =   "Cerrar Bulto"
   ClientHeight    =   8160
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "NumerosIngresosProcesados.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8160
   ScaleWidth      =   9480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8055
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   9495
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   8640
         Picture         =   "NumerosIngresosProcesados.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   45
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   43
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   5040
         TabIndex        =   42
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   1800
            TabIndex        =   48
            Top             =   360
            Width           =   1575
         End
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
         Bindings        =   "NumerosIngresosProcesados.frx":074C
         Height          =   6855
         Left            =   120
         OleObjectBlob   =   "NumerosIngresosProcesados.frx":0768
         TabIndex        =   46
         Top             =   1080
         Width           =   9255
      End
      Begin VB.Label LblBuscar 
         Alignment       =   1  'Right Justify
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   720
         Width           =   975
      End
   End
   Begin Crystal.CrystalReport CrReportes 
      Left            =   2400
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowState     =   2
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   555
      Left            =   6360
      Picture         =   "NumerosIngresosProcesados.frx":1143
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7080
      Width           =   1500
   End
   Begin TabDlg.SSTab TabIngresos 
      Height          =   6975
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   12303
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "NumerosIngresosProcesados.frx":1585
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameNumerosIngresos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "NumerosIngresosProcesados.frx":189F
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridNumerosIngresos"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "NumerosIngresosProcesados.frx":1CF1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Lbletiqueta"
      Tab(2).Control(1)=   "Label2(8)"
      Tab(2).Control(2)=   "Label2(9)"
      Tab(2).Control(3)=   "LblBusCodMatPri"
      Tab(2).Control(4)=   "FrameOpciones"
      Tab(2).Control(5)=   "TxtBuscar"
      Tab(2).Control(6)=   "DtpFecIni"
      Tab(2).Control(7)=   "DtpFecFin"
      Tab(2).Control(8)=   "CmdBuscar"
      Tab(2).Control(9)=   "CmdActualizar"
      Tab(2).Control(10)=   "TxtBusCodMatPri"
      Tab(2).ControlCount=   11
      Begin VB.TextBox TxtBusCodMatPri 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -67680
         TabIndex        =   24
         Top             =   4320
         Width           =   1935
      End
      Begin VB.CommandButton CmdActualizar 
         Caption         =   "Actualizar"
         Height          =   735
         Left            =   -69720
         Picture         =   "NumerosIngresosProcesados.frx":2143
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   5760
         Width           =   1935
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Datos"
         Height          =   735
         Left            =   -67680
         Picture         =   "NumerosIngresosProcesados.frx":244D
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   5760
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -67680
         TabIndex        =   22
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24576003
         CurrentDate     =   37301
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -67680
         TabIndex        =   21
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   24576003
         CurrentDate     =   37301
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         Height          =   285
         Left            =   -67680
         TabIndex        =   23
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
         TabIndex        =   30
         Top             =   720
         Width           =   5205
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Fecha Y Materia Prima"
            Height          =   1305
            Index           =   2
            Left            =   3480
            Picture         =   "NumerosIngresosProcesados.frx":288F
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
         Begin VB.OptionButton OptBusqueda 
            Caption         =   "Numero Ingreso y Codigo Materia Prima"
            Height          =   1305
            Index           =   0
            Left            =   120
            Picture         =   "NumerosIngresosProcesados.frx":3159
            Style           =   1  'Graphical
            TabIndex        =   18
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
            Picture         =   "NumerosIngresosProcesados.frx":3463
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   " "
            Top             =   360
            Width           =   1575
         End
      End
      Begin VB.Frame FrameNumerosIngresos 
         Caption         =   "Datos De Numero Ingreso"
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
         TabIndex        =   28
         Top             =   720
         Width           =   9195
         Begin MSMask.MaskEdBox MskCanTotPro 
            DataField       =   "CantidadTotalProcesada"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   7080
            TabIndex        =   60
            Top             =   5400
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskCanPro 
            DataField       =   "CantidadProcesada"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   7080
            TabIndex        =   10
            Top             =   5040
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MskConFin 
            DataField       =   "ContadorFinal"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   7080
            TabIndex        =   9
            Top             =   4680
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtBodSal 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "BodegaSalida"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   3240
            Width           =   615
         End
         Begin VB.TextBox TxtTur 
            Appearance      =   0  'Flat
            DataField       =   "Turno"
            DataSource      =   "DataNumerosIngresos"
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
            Left            =   4560
            MaxLength       =   1
            TabIndex        =   2
            Top             =   360
            Width           =   375
         End
         Begin MSMask.MaskEdBox MskConIni 
            DataField       =   "ContadorInicial"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   7080
            TabIndex        =   8
            Top             =   4320
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "#,###,###"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtFicTec 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "FichaTecnica"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            MaxLength       =   12
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1815
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            TabIndex        =   0
            Top             =   360
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Text1 
            DataField       =   "UsuarioEditar"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Index           =   1
            Left            =   600
            MaxLength       =   10
            TabIndex        =   39
            Top             =   720
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox Text1 
            DataField       =   "UsuarioAgregar"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   10
            TabIndex        =   38
            Top             =   720
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.TextBox TxtLin 
            Appearance      =   0  'Flat
            DataField       =   "Linea"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            MaxLength       =   2
            TabIndex        =   3
            Top             =   1080
            Width           =   1815
         End
         Begin MSMask.MaskEdBox MskHor 
            DataField       =   "Hora"
            DataSource      =   "DataNumerosIngresos"
            Height          =   285
            Left            =   2040
            TabIndex        =   1
            Top             =   720
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            Format          =   "HH:mm"
            PromptChar      =   "_"
         End
         Begin VB.TextBox TxtCod 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            DataField       =   "CodigoMateriaPrima"
            DataSource      =   "DataNumerosIngresos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2040
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1800
            Width           =   4335
         End
         Begin VB.TextBox TxtNumIng 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "NumeroIngreso"
            DataSource      =   "DataNumerosIngresos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   2040
            TabIndex        =   6
            ToolTipText     =   " "
            Top             =   2520
            Width           =   4335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "Total Procesado"
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
            Left            =   5160
            TabIndex        =   61
            Top             =   5400
            Width           =   1410
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
            Left            =   3960
            TabIndex        =   59
            Top             =   1080
            Width           =   5055
         End
         Begin VB.Label LblBodega 
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
            Left            =   2760
            TabIndex        =   58
            Top             =   3240
            Width           =   3615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "Cantidad Procesada"
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
            Left            =   5160
            TabIndex        =   57
            Top             =   5040
            Width           =   1725
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "Contador Final"
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
            Left            =   5160
            TabIndex        =   56
            Top             =   4680
            Width           =   1245
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Disponible En Bodega"
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
            Index           =   18
            Left            =   120
            TabIndex        =   55
            Top             =   3240
            Width           =   1890
         End
         Begin VB.Label Label2 
            Caption         =   "Turno"
            Height          =   255
            Index           =   17
            Left            =   3960
            TabIndex        =   54
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            Caption         =   "Contador Inicial"
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
            Left            =   5160
            TabIndex        =   53
            Top             =   4320
            Width           =   1350
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
            Height          =   255
            Left            =   3960
            TabIndex        =   52
            Top             =   1440
            Width           =   5055
         End
         Begin VB.Label Label2 
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
            Index           =   3
            Left            =   120
            TabIndex        =   51
            Top             =   1440
            Width           =   1230
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
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   35
            Top             =   1080
            Width           =   480
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   34
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
            TabIndex        =   33
            Top             =   2040
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
            TabIndex        =   32
            Top             =   2760
            Width           =   1350
         End
         Begin VB.Label Label2 
            Caption         =   "Hora"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H008080FF&
            BackStyle       =   1  'Opaque
            Height          =   2055
            Left            =   4560
            Shape           =   2  'Oval
            Top             =   3960
            Width           =   4455
         End
      End
      Begin MSDBGrid.DBGrid DBGridNumerosIngresos 
         Bindings        =   "NumerosIngresosProcesados.frx":376D
         Height          =   6105
         Left            =   -74880
         OleObjectBlob   =   "NumerosIngresosProcesados.frx":378F
         TabIndex        =   17
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
         TabIndex        =   40
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   31
         Top             =   3960
         Width           =   2775
      End
   End
   Begin VB.Data DataNumerosIngresos 
      BackColor       =   &H80000014&
      Caption         =   "Numeros Ingresos"
      Connect         =   "Access"
      DatabaseName    =   "D:\Amapro\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "NumerosIngresosProcesados"
      Top             =   7680
      Width           =   9315
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   555
      Left            =   7920
      MouseIcon       =   "NumerosIngresosProcesados.frx":5430
      Picture         =   "NumerosIngresosProcesados.frx":5872
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   555
      Left            =   4800
      MouseIcon       =   "NumerosIngresosProcesados.frx":5CB4
      Picture         =   "NumerosIngresosProcesados.frx":60F6
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   3240
      MouseIcon       =   "NumerosIngresosProcesados.frx":6628
      Picture         =   "NumerosIngresosProcesados.frx":6A6A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   1680
      MouseIcon       =   "NumerosIngresosProcesados.frx":6F9C
      Picture         =   "NumerosIngresosProcesados.frx":73DE
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7080
      Width           =   1500
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   555
      Left            =   120
      MouseIcon       =   "NumerosIngresosProcesados.frx":7910
      Picture         =   "NumerosIngresosProcesados.frx":7D52
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7080
      Width           =   1500
   End
End
Attribute VB_Name = "NumerosIngresosProcesados"
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
Dim RBuscaTotalProcesado As Recordset
Dim RBuscaExistencia As Recordset
Dim RBuscaContadorInicial As Recordset
Dim RBuscaLinea As Recordset
Dim RBuscaFichaTecnica As Recordset
Dim RBuscaFicha As Recordset
Dim RBuscaBodega As Recordset

Dim VCodigoMateriaPrima As String
Dim VNumeroIngreso As Double
Dim VBodega As String
Dim VCantidadMateriaPrima As Double

Dim BCamposVisibles As Boolean
Dim BLinea As Boolean
Dim BMateriaPrima As Boolean


Sub botones()
    If Bandera = True Then
         FrameNumerosIngresos.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdImprimir.Enabled = False
         CmdSalida.Enabled = False
         TxtCod.SetFocus
         Lbletiqueta.Visible = False
         TxtBuscar.Visible = False
         DataNumerosIngresos.Visible = False
         FrameOpciones.Visible = False
         DBGridNumerosIngresos.Visible = False
    Else
         FrameNumerosIngresos.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdImprimir.Enabled = True
         CmdSalida.Enabled = True
         Lbletiqueta.Visible = True
         TxtBuscar.Visible = True
         DataNumerosIngresos.Visible = True
         FrameOpciones.Visible = True
         DBGridNumerosIngresos.Visible = True
    End If
End Sub

Private Sub CmdActualizar_Click()
    DataNumerosIngresos.RecordSource = "Select * From NumerosIngresosMPaProducir"
    DataNumerosIngresos.Refresh
    DBGridNumerosIngresos.Refresh
End Sub

Private Sub CmdAgregar_Click()
On Error Resume Next
       'AGREGA UN REGISTRO
        DataNumerosIngresos.Recordset.AddNew
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
        
        'ASIGNA USUARIO
        Text1.Item(0).Text = GUsuario
        TxtLin.SetFocus
        MskConIni.Text = "0"
        MskConFin.Text = "0"
        MskCanPro.Text = "0"
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            

            VCodigoMateriaPrima = TxtCod.Text
            VBodega = TxtBodSal.Text
            VCantidadMateriaPrima = MskCanPro.Text
            VNumeroIngreso = TxtNumIng.Text

            Mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If Mensaje = vbOK Then
                'BORRA EL REGISTRO
                DataNumerosIngresos.Recordset.Delete
                
                If Err <> 0 Then
                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                    Exit Sub
                End If
                
                'BUSCA LA EXISTENCIA Y LAS SALIDAS DE LA MATERIA PRIMA Y SUMA LA EXISTENCIA Y RESTA SALIDAS
                Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Salidas From InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and Bodega = '" & VBodega & "'")
                   If RBuscaMateriaPrima.RecordCount > 0 Then
                           RBuscaMateriaPrima.Edit
                               RBuscaMateriaPrima!Existencia = (RBuscaMateriaPrima!Existencia + VCantidadMateriaPrima)
                               RBuscaMateriaPrima!Salidas = (RBuscaMateriaPrima!Salidas - VCantidadMateriaPrima)
                           RBuscaMateriaPrima.Update
                   End If
                
                'BUSCA EL NUMERO DE INGRESO CON CODIGO DE MATERIA PRIMA EN EL DETALLE DE LAS ENTRADAS
                 Set RBuscaNumeroIngresoEntradas = Db.OpenRecordset("Select Disponibilidad, SaldoDisponibilidad, CantidadSalida From DetalleEntradasMateriaPrima Where Codigo = '" & VCodigoMateriaPrima & "' And NumeroIngreso = " & VNumeroIngreso)
                     If RBuscaNumeroIngresoEntradas.RecordCount > 0 Then
                            'ASIGNA LA DISPONIBILIDAD AL BULTO O NUMERO DE INGRESO Y CONTROLA EL SALDO DEL BULTO
                            RBuscaNumeroIngresoEntradas.Edit
                                    RBuscaNumeroIngresoEntradas!CantidadSalida = RBuscaNumeroIngresoEntradas!CantidadSalida - VCantidadMateriaPrima
                                    RBuscaNumeroIngresoEntradas!SaldoDisponibilidad = RBuscaNumeroIngresoEntradas!SaldoDisponibilidad + VCantidadMateriaPrima
                                    'SI EL SALDO ES MAYOR QUE CERO ENTONCES EL BULTO VUELVE A EXISTIR
                                    If RBuscaNumeroIngresoEntradas!SaldoDisponibilidad > 0 Then
                                        RBuscaNumeroIngresoEntradas!Disponibilidad = "EXISTE"
                                    End If
                            RBuscaNumeroIngresoEntradas.Update
                            If Err <> 0 Then
                            End If
                     End If
                
                DataNumerosIngresos.Recordset.MoveLast
            End If
  
            If DataNumerosIngresos.Recordset.EOF Then
                DataNumerosIngresos.Recordset.MoveLast
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
            DataNumerosIngresos.RecordSource = "Select * From NumerosIngresosMPaProducir Where NumeroIngreso = " & TxtBuscar.Text & " And Codigo = '" & TxtBusCodMatPri.Text & "'"
            TabIngresos.Tab = 0
        'FECHAS
        ElseIf OptBusqueda.Item(1).Value = True Then
            DataNumerosIngresos.RecordSource = "Select * From NumerosIngresosMPaProducir Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "#"
            TabIngresos.Tab = 1
        'FECHA Y MATERIA PRIMA
        ElseIf OptBusqueda.Item(2).Value = True Then
            DataNumerosIngresos.RecordSource = "Select * From NumerosIngresosMPaProducir Where Fecha >= #" & Format(DtpFecIni.Value, "mm/dd/yyyy") & "# And Fecha <= #" & Format(DtpFecFin.Value, "mm/dd/yyyy") & "# And Codigo = '" & TxtBuscar.Text & "'"
            TabIngresos.Tab = 1
        End If
        DataNumerosIngresos.Refresh
        DBGridNumerosIngresos.Refresh
        
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
        
        DataNumerosIngresos.Recordset.CancelUpdate
        
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
            Exit Sub
        End If
        
        Bandera = False
        botones

End Sub


Private Sub CmdGrabar_Click()
   On Error Resume Next
   
    VCodigoMateriaPrima = TxtCod.Text
    VBodega = TxtBodSal.Text
    VCantidadMateriaPrima = MskCanPro.Text
    VNumeroIngreso = TxtNumIng.Text
    
    'VERIFICA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta ", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    'VERIFICA EL TURNO
    If TxtTur.Text <> "D" And TxtTur.Text <> "N" Then
        MsgBox "Turno Incorrecto", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
    'VERIFICA LA LINEA
    If TxtLin.Text = "" Then
        MsgBox "Linea Incorrecta", vbOKOnly + vbInformation, "Informacion"
        Exit Sub
    End If
    
       
   'GRABA DATOS
   DataNumerosIngresos.Recordset.Update
   
   If Err <> 0 Then
        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
   Else
                'BUSCA LA EXISTENCIA Y LAS SALIDAS DE LA MATERIA PRIMA Y RESTA LA EXISTENCIA Y SUMA SALIDAS
                Set RBuscaMateriaPrima = Db.OpenRecordset("Select Existencia, Salidas From InventarioMateriaPrima where CodigoMateriaPrima = '" & VCodigoMateriaPrima & "' and Bodega = '" & VBodega & "'")
                   If RBuscaMateriaPrima.RecordCount > 0 Then
                           RBuscaMateriaPrima.Edit
                               RBuscaMateriaPrima!Existencia = (RBuscaMateriaPrima!Existencia - VCantidadMateriaPrima)
                               RBuscaMateriaPrima!Salidas = (RBuscaMateriaPrima!Salidas + VCantidadMateriaPrima)
                           RBuscaMateriaPrima.Update
                   End If
                   
                
                'BUSCA EL NUMERO DE INGRESO CON CODIGO DE MATERIA PRIMA EN EL DETALLE DE LAS ENTRADAS
                 Set RBuscaNumeroIngresoEntradas = Db.OpenRecordset("Select Disponibilidad, SaldoDisponibilidad, CantidadSalida From DetalleEntradasMateriaPrima Where Codigo = '" & VCodigoMateriaPrima & "' And NumeroIngreso = " & VNumeroIngreso)
                     If RBuscaNumeroIngresoEntradas.RecordCount > 0 Then
                            'ASIGNA LA DISPONIBILIDAD AL BULTO O NUMERO DE INGRESO Y CONTROLA EL SALDO DEL BULTO
                            RBuscaNumeroIngresoEntradas.Edit
                                    RBuscaNumeroIngresoEntradas!CantidadSalida = RBuscaNumeroIngresoEntradas!CantidadSalida + VCantidadMateriaPrima
                                    RBuscaNumeroIngresoEntradas!SaldoDisponibilidad = RBuscaNumeroIngresoEntradas!SaldoDisponibilidad - VCantidadMateriaPrima
                                    'SI EL SALDO QUE QUEDA ES CERO ENTONCES EL BULTO YA NO EXISTE
                                    If RBuscaNumeroIngresoEntradas!SaldoDisponibilidad <= 0 Then
                                        RBuscaNumeroIngresoEntradas!Disponibilidad = "NO EXISTE"
                                    End If
                            RBuscaNumeroIngresoEntradas.Update
                            If Err <> 0 Then
                            End If
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
        
                CrReportes.SelectionFormula = "{NumerosIngresosMPaProducir.NumeroIngreso} = " & TxtNumIng.Text & " And {NumerosIngresosMPaProducir.Codigo} = '" & TxtCod.Text & "'"
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
    'LINEAS
    If BLinea = True Then
        TxtLin.Text = DBGridConsultas.Columns(0).Text
        TxtLin.SetFocus
    'MATERIAS PRIMAS
    ElseIf BMateriaPrima = True Then
        TxtCod.Text = DBGridConsultas.Columns(0).Text
        TxtCod.SetFocus
    End If
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        'LINEAS
        If BLinea = True Then
            TxtLin.Text = DBGridConsultas.Columns(0).Text
            TxtLin.SetFocus
        'MATERIAS PRIMAS
        ElseIf BMateriaPrima = True Then
            TxtCod.Text = DBGridConsultas.Columns(0).Text
            TxtCod.SetFocus
        End If
            FrameConsultas.Visible = False
    End If
End Sub

Private Sub dbgridnumerosingresos_HeadClick(ByVal ColIndex As Integer)
    DataNumerosIngresos.RecordSource = ("Select * from NumerosIngresosMPaProducir order by " & DBGridNumerosIngresos.Columns(ColIndex).DataField)
    DataNumerosIngresos.Refresh
    DBGridNumerosIngresos.Refresh
    
End Sub

Private Sub Form_Load()
    DataNumerosIngresos.DatabaseName = BasedeDatos
    DataConsultas.DatabaseName = BasedeDatos
End Sub

Private Sub MskCanPro_GotFocus()
        MskCanPro.SelStart = 0
        MskCanPro.SelLength = Len(MskCanPro.Text)
End Sub

Private Sub MskCanPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskCanPro_LostFocus()
        MskCanTotPro.Text = Val(MskCanPro.Text) + Val(MskCanTotPro.Text)
End Sub

Private Sub MskConFin_GotFocus()
        MskConFin.SelStart = 0
        MskConFin.SelLength = Len(MskConFin.Text)
End Sub

Private Sub MskConFin_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub MskConFin_LostFocus()
        MskCanPro.Text = Val(MskConFin.Text) - Val(MskConIni.Text)
End Sub

Private Sub MskConIni_GotFocus()
        MskConIni.SelStart = 0
        MskConIni.SelLength = Len(MskConIni.Text)
End Sub

Private Sub MskConIni_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
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

Private Sub MskHor_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

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

Private Sub TxtBodSal_Change()
    Set RBuscaBodega = Db.OpenRecordset("Select * From BodegasMateriaPrima Where CodigoBodega = '" & TxtBodSal.Text & "'")
        If RBuscaBodega.RecordCount > 0 Then
            LblBodega.Caption = RBuscaBodega!Descripcion
        Else
            LblBodega.Caption = ""
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


Private Sub TxtConsultas_Change()
    'LINEA
    If BLinea = True Then
        'DESCRIPCION
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec And L.Descrip Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec And L.Descrip Like '*" & TxtConsultas.Text & "*'"
            End If
        'CODIGO
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec And L.Linea Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec And L.Linea Like '*" & TxtConsultas.Text & "*'"
            End If
        End If
    'MATERIA PRIMA
    ElseIf BMateriaPrima = True Then
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion, Bodega From InventarioMateriaPrima Where Descripcion Like '" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima, Bodega"
            Else
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion, Bodega From InventarioMateriaPrima Where Descripcion Like '*" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima, Bodega"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion, Bodega From InventarioMateriaPrima Where CodigoMateriaPrima Like '" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima, Bodega"
            Else
                DataConsultas.RecordSource = "Select CodigoMateriaPrima, Descripcion, Bodega From InventarioMateriaPrima Where CodigoMateriaPrima Like '*" & TxtConsultas.Text & "*' Order By CodigoMateriaPrima, Bodega"
            End If
        End If
    End If
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    AnchoColumnas
    
End Sub

Private Sub TxtConsultas_GotFocus()
    TxtConsultas.SelStart = 0
    TxtConsultas.SelLength = Len(TxtConsultas.Text)
End Sub

Private Sub TxtConsultas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
End Sub


Private Sub TxtFicTec_Change()
    'BUSCAMOS LOS CODIGOS DE MATERIA PRIMA DE ACUERDO A LA FICHA TECNICA
        Set RBuscaFicha = Db.OpenRecordset("Select * From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
        If RBuscaFicha.RecordCount > 0 Then
            LblFichaTecnica.Caption = RBuscaFicha!Descrip
                    TxtCod.Text = RBuscaFicha!Platina
        Else
                    TxtCod.Text = ""
                    LblFichaTecnica.Caption = ""
                    
                    
        End If
End Sub

Private Sub TxtFicTec_LostFocus()
        'BUSCAMOS LOS CODIGOS DE MATERIA PRIMA DE ACUERDO A LA FICHA TECNICA
        Set RBuscaFichaTecnica = Db.OpenRecordset("Select * From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
            'SI ENCUENTRA LA FICHA TECNICA
            If RBuscaFichaTecnica.RecordCount > 0 Then
                    'ASIGNA EL CODIGO DE FICHA TECNICA
                    TxtFicTec.Text = RBuscaFichaTecnica!Esp_tec
                    TxtCod.Text = RBuscaFichaTecnica!Platina
            Else
                    TxtFicTec.Text = ""
                    TxtCod.Text = ""
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
    BLinea = True
    BMateriaPrima = False
    DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec"
    DataConsultas.Refresh
    DBGridConsultas.Refresh
    AnchoColumnas
    FrameConsultas.Visible = True
    TxtConsultas.SetFocus
    
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
            BLinea = True
            BMateriaPrima = False
            DataConsultas.RecordSource = "Select L.Linea, L.Descrip, L.Esp_Tec, FT.Descrip From Lineas as L, FichaTecnica as FT Where L.Esp_Tec = FT.Esp_Tec"
            DataConsultas.Refresh
            DBGridConsultas.Refresh
            AnchoColumnas
            FrameConsultas.Visible = True
            TxtConsultas.SetFocus
    End If
End Sub

Private Sub TxtLin_LostFocus()
    'BUSCAMOS LA LINEA PARA JALAR LA FICHA TECNICA
    Set RBuscaLinea = Db.OpenRecordset("Select Esp_Tec From Lineas Where Linea = '" & TxtLin.Text & "'")
    If RBuscaLinea.RecordCount > 0 Then
        'BUSCAMOS LOS CODIGOS DE MATERIA PRIMA DE ACUERDO A LA FICHA TECNICA
        Set RBuscaFichaTecnica = Db.OpenRecordset("Select * From FichaTecnica Where Esp_Tec = '" & RBuscaLinea!Esp_tec & "'")
            'SI ENCUENTRA LA FICHA TECNICA
            If RBuscaFichaTecnica.RecordCount > 0 Then
                    'ASIGNA EL CODIGO DE FICHA TECNICA
                    TxtFicTec.Text = RBuscaFichaTecnica!Esp_tec
                    TxtCod.Text = RBuscaFichaTecnica!Platina
            Else
                    TxtFicTec.Text = ""
                    TxtCod.Text = ""
            End If
    Else
        'MsgBox "Linea No Existe", vbOKOnly + vbInformation, "Informacion"
    End If
    
End Sub

Private Sub TxtNumIng_GotFocus()
    TxtNumIng.SelStart = 0
    TxtNumIng.SelLength = Len(TxtNumIng.Text)
End Sub

Private Sub TxtNumIng_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub


Sub AnchoColumnas()
    If BLinea = True Then
        DBGridConsultas.Columns(0).Width = "500"
        DBGridConsultas.Columns(1).Width = "3000"
        DBGridConsultas.Columns(2).Width = "1200"
        DBGridConsultas.Columns(3).Width = "4000"
    Else
        DBGridConsultas.Columns(0).Width = "1200"
        DBGridConsultas.Columns(1).Width = "4000"
        DBGridConsultas.Columns(2).Width = "1200"
    End If
    
        
End Sub

Private Sub TxtNumIng_LostFocus()
            If IsNumeric(TxtNumIng.Text) Then
                'BUSCA EL NUMERO DE INGRESO Y ASIGNA LA BODEGA, CODIGO Y CANTIDAD DE ACUERDO COMO ENTRO A LA BODEGA
                Set RBuscaNumeroIngreso = Db.OpenRecordset("Select SaldoDisponibilidad, BodegaDisponibilidad From DetalleEntradasMateriaPrima Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCod.Text & "'")
                'SI ENCUENTRA EL INGRESO ASIGNA A LOS TEXT LA CANTIDAD, BODEGA, CODIGO
                If RBuscaNumeroIngreso.RecordCount > 0 Then
                    If IsNull(RBuscaNumeroIngreso(0)) Then
                        TxtBodSal.Text = ""
                    Else
                        TxtBodSal.Text = RBuscaNumeroIngreso!BodegaDisponibilidad
                    End If
                Else
                    TxtBodSal.Text = ""
                End If
                
                'BUSCA LA SUMA DE EL CONTADOR INICIAL DEL BULTO INGRESADO ANTERIORMENTE
                Set RBuscaContadorInicial = Db.OpenRecordset("Select SUM(ContadorInicial) From NumerosIngresosMPaProducir Where NumeroIngreso = " & TxtNumIng.Text & " And Codigo = '" & TxtCod.Text & "'")
                    If RBuscaContadorInicial.RecordCount > 0 Then
                        If IsNull(RBuscaContadorInicial(0)) Then
                            MskConIni.Text = "0"
                        Else
                            MskConIni.Text = RBuscaContadorInicial(0)
                        End If
                    Else
                        MskConIni.Text = "0"
                    End If
                    
                'BUSCA EL TOTAL PROCESADO EN LOS BULTOS
                Set RBuscaTotalProcesado = Db.OpenRecordset("Select Sum(CantidadTotalProcesada) From NumerosIngresosProcesados Where NumeroIngreso = " & TxtNumIng.Text & " And CodigoMateriaPrima = '" & TxtCod.Text & "'")
                    If RBuscaTotalProcesado.RecordCount > 0 Then
                        If IsNull(RBuscaTotalProcesado(0)) Then
                            MskCanTotPro.Text = "0"
                        Else
                            MskCanTotPro.Text = RBuscaTotalProcesado(0)
                        End If
                    Else
                        MskCanTotPro.Text = "0"
                    End If
            End If

End Sub

Private Sub TxtTur_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If

End Sub
