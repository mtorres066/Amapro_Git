VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form EmpleadosCapturaHoras 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura De Horas Extras"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "EmpleadosCapturaHoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameConsultas 
      Caption         =   "Consulta de Datos "
      Height          =   8175
      Left            =   0
      TabIndex        =   41
      Top             =   0
      Visible         =   0   'False
      Width           =   8415
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
      Begin VB.Frame FrameTipoDeBusqueda 
         Caption         =   "Tipo De Busqueda"
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   3960
         TabIndex        =   46
         Top             =   120
         Width           =   3495
         Begin VB.OptionButton OptCuaPal 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Left            =   1800
            TabIndex        =   48
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptPalIni 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptCod 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   1920
         TabIndex        =   43
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptDes 
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   360
         TabIndex        =   42
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtConsultas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   44
         Top             =   720
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Height          =   735
         Left            =   7560
         Picture         =   "EmpleadosCapturaHoras.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   240
         Width           =   735
      End
      Begin MSDBGrid.DBGrid DBGridConsultas 
         Bindings        =   "EmpleadosCapturaHoras.frx":293C
         Height          =   6975
         Left            =   120
         OleObjectBlob   =   "EmpleadosCapturaHoras.frx":2958
         TabIndex        =   45
         Top             =   1080
         Width           =   8175
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
   Begin TabDlg.SSTab TabEmpleados 
      Height          =   6735
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   11880
      _Version        =   393216
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual "
      TabPicture(0)   =   "EmpleadosCapturaHoras.frx":3333
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameEmpleados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "EmpleadosCapturaHoras.frx":364D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGridEmpleados"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Busqueda De Datos"
      TabPicture(2)   =   "EmpleadosCapturaHoras.frx":3A9F
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "LblBusqueda"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "CmdBuscar(0)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "CmdBuscar(1)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "DtpFecIni"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "DtpFecFin"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "OptOpcion(0)"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "OptOpcion(3)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "OptOpcion(2)"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "OptOpcion(11)"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "TxtBusqueda"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   67
         Top             =   2880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Fechas Y Turno"
         Height          =   195
         Index           =   11
         Left            =   -74400
         TabIndex        =   66
         Top             =   1320
         Width           =   1455
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Fechas Y Linea"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   65
         Top             =   1680
         Width           =   1455
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Fechas y Empleado"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   64
         Top             =   2040
         Width           =   1935
      End
      Begin VB.OptionButton OptOpcion 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   63
         Top             =   960
         Value           =   -1  'True
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DtpFecFin 
         Height          =   255
         Left            =   -69000
         TabIndex        =   26
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   61669377
         CurrentDate     =   37588
      End
      Begin MSComCtl2.DTPicker DtpFecIni 
         Height          =   255
         Left            =   -71160
         TabIndex        =   25
         Top             =   2280
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61669379
         CurrentDate     =   37588
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccionar Todos"
         Height          =   855
         Index           =   1
         Left            =   -66720
         Picture         =   "EmpleadosCapturaHoras.frx":3EF1
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4800
         Width           =   2055
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Seleccion o Busqueda"
         Height          =   855
         Index           =   0
         Left            =   -66720
         Picture         =   "EmpleadosCapturaHoras.frx":41FB
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3840
         Width           =   2055
      End
      Begin VB.Frame FrameEmpleados 
         Caption         =   "Datos Del Empleado"
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
         Height          =   5775
         Left            =   1680
         TabIndex        =   30
         Top             =   720
         Width           =   8175
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasDoblesProyectadas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   16
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   13
            Top             =   5040
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasNocturnasProyectadas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   15
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   11
            Top             =   3960
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasDiurnasProyectadas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   14
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   8
            Top             =   2880
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasLaboradasNocturnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   7
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   2160
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasLaboradasNocturnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   6
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   6
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasExtrasDobles"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   13
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   5400
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasDobles"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   12
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   14
            Top             =   5400
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasExtrasNocturnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   11
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   4320
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasNocturnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   10
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   12
            Top             =   4320
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasExtrasDiurnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   9
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   3240
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasExtrasDiurnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   8
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   9
            Top             =   3240
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "MontoHorasLaboradasDiurnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   5
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            DataField       =   "Usuario"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   2
            Left            =   6240
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "HorasLaboradasDiurnas"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   4
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Empleado"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   3
            Left            =   960
            MaxLength       =   4
            TabIndex        =   3
            Top             =   1440
            Width           =   1455
         End
         Begin MSMask.MaskEdBox MskFec 
            DataField       =   "Fecha"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   960
            TabIndex        =   0
            Top             =   360
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Linea"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   1
            Left            =   960
            MaxLength       =   2
            TabIndex        =   2
            Top             =   1080
            Width           =   1455
         End
         Begin VB.TextBox Txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            DataField       =   "Turno"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Index           =   0
            Left            =   960
            MaxLength       =   1
            TabIndex        =   1
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Dobles Proyectadas"
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
            Index           =   17
            Left            =   120
            TabIndex        =   62
            Top             =   5040
            Width           =   2265
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hor. Ext. Diurnas Proyect."
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
            Index           =   16
            Left            =   120
            TabIndex        =   61
            Top             =   2880
            Width           =   2235
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hor. Ext. Nocturnas Proyect."
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
            Index           =   15
            Left            =   120
            TabIndex        =   60
            Top             =   3960
            Width           =   2460
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Laboradas Nocturnas"
            Height          =   195
            Index           =   14
            Left            =   4200
            TabIndex        =   59
            Top             =   2160
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Laboradas Nocturnas"
            Height          =   195
            Index           =   13
            Left            =   120
            TabIndex        =   58
            Top             =   2160
            Width           =   1995
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Laboradas Diurnas"
            Height          =   195
            Index           =   12
            Left            =   4200
            TabIndex        =   57
            Top             =   1800
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Extras Diurnas"
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
            Index           =   11
            Left            =   120
            TabIndex        =   56
            Top             =   3240
            Width           =   1800
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Extras Diurnas"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   10
            Left            =   4200
            TabIndex        =   55
            Top             =   3240
            Visible         =   0   'False
            Width           =   1515
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Extras Nocturnas"
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
            Index           =   9
            Left            =   120
            TabIndex        =   54
            Top             =   4320
            Width           =   2025
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Extras Nocturnas"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   8
            Left            =   4200
            TabIndex        =   53
            Top             =   4320
            Visible         =   0   'False
            Width           =   1710
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Extras Dobles"
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
            Index           =   7
            Left            =   120
            TabIndex        =   52
            Top             =   5400
            Width           =   1740
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Monto Extras Dobles"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   6
            Left            =   4200
            TabIndex        =   51
            Top             =   5400
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Usuario"
            Height          =   195
            Index           =   5
            Left            =   5640
            TabIndex        =   40
            Top             =   240
            Width           =   540
         End
         Begin VB.Label LblEmpleado 
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
            Left            =   2520
            TabIndex        =   37
            Top             =   1440
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
            Left            =   2520
            TabIndex        =   36
            Top             =   1080
            Width           =   5535
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Horas Laboradas Diurnas"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   35
            Top             =   1800
            Width           =   1800
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Empleado"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   34
            Top             =   1440
            Width           =   705
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   33
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Linea"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   390
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Turno"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   31
            Top             =   720
            Width           =   420
         End
      End
      Begin MSDBGrid.DBGrid DBGridEmpleados 
         Bindings        =   "EmpleadosCapturaHoras.frx":463D
         Height          =   5865
         Left            =   -74880
         OleObjectBlob   =   "EmpleadosCapturaHoras.frx":4659
         TabIndex        =   24
         Top             =   720
         Width           =   11505
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
         Left            =   -72360
         TabIndex        =   68
         Top             =   2880
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Left            =   -69600
         TabIndex        =   39
         Top             =   2280
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Left            =   -71760
         TabIndex        =   38
         Top             =   2280
         Width           =   465
      End
   End
   Begin VB.Data DataEmpleados 
      BackColor       =   &H80000014&
      Caption         =   "Captura Horas Extras"
      Connect         =   "Access"
      DatabaseName    =   "D:\Visual Basic\Amapro Metalenvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EmpleadosCapturaHoras"
      Top             =   7800
      Width           =   11595
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Salida"
      Height          =   800
      Index           =   5
      Left            =   9840
      MouseIcon       =   "EmpleadosCapturaHoras.frx":637C
      Picture         =   "EmpleadosCapturaHoras.frx":67BE
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "B&orrar"
      Height          =   800
      Index           =   4
      Left            =   7920
      MouseIcon       =   "EmpleadosCapturaHoras.frx":8830
      Picture         =   "EmpleadosCapturaHoras.frx":8C72
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   3
      Left            =   6000
      MouseIcon       =   "EmpleadosCapturaHoras.frx":91A4
      Picture         =   "EmpleadosCapturaHoras.frx":95E6
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Index           =   2
      Left            =   4080
      MouseIcon       =   "EmpleadosCapturaHoras.frx":9B18
      Picture         =   "EmpleadosCapturaHoras.frx":9F5A
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Editar"
      Height          =   800
      Index           =   1
      Left            =   2160
      MouseIcon       =   "EmpleadosCapturaHoras.frx":A48C
      Picture         =   "EmpleadosCapturaHoras.frx":A8CE
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6840
      Width           =   1800
   End
   Begin VB.CommandButton CmdBotones 
      Caption         =   "&Agregar"
      Height          =   800
      Index           =   0
      Left            =   240
      MouseIcon       =   "EmpleadosCapturaHoras.frx":AE00
      Picture         =   "EmpleadosCapturaHoras.frx":B242
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6840
      Width           =   1800
   End
End
Attribute VB_Name = "EmpleadosCapturaHoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim BTurno As Boolean
Dim BLinea As Boolean
Dim BEmpleado As Boolean

Dim RBuscaLineas As Recordset
Dim RBuscaEmpleado As Recordset

Dim RBuscaSueldoBase As Recordset
Dim RBuscaFactores As Recordset
Dim VDiasMes As Integer

Dim VValorHoraLaboradaDiurna As Currency
Dim VValorHoraLaboradaNocturna As Currency
Dim VValorHoraExtraDiurna As Currency
Dim VValorHoraExtraNocturna As Currency
Dim VSueldoBase As Currency
Dim VFHorasDiurnas As Single
Dim VFHorasNocturnas As Single
Dim VFPorcentajeDiurnas As Single
Dim VFPorcentajeNocturnas As Single




Sub botones()
    If Bandera = True Then
         FrameEmpleados.Enabled = True
         CmdBotones.Item(0).Enabled = False
         CmdBotones.Item(1).Enabled = False
         CmdBotones.Item(2).Enabled = True
         CmdBotones.Item(3).Enabled = True
         CmdBotones.Item(4).Enabled = False
         CmdBotones.Item(5).Enabled = False
         TxtTexto.Item(0).SetFocus
         DataEmpleados.Visible = False
         DBGridEmpleados.Visible = False
    Else
         FrameEmpleados.Enabled = False
         CmdBotones.Item(0).Enabled = True
         CmdBotones.Item(1).Enabled = True
         CmdBotones.Item(2).Enabled = False
         CmdBotones.Item(3).Enabled = False
         CmdBotones.Item(4).Enabled = True
         CmdBotones.Item(5).Enabled = True
         DataEmpleados.Visible = True
         DBGridEmpleados.Visible = True
    End If
End Sub
Private Sub CmdBotones_Click(Index As Integer)
    On Error Resume Next
        With DataEmpleados.Recordset
            If Index = 0 Then
                    'TxtTexto.Item(4).Text = 0
                    'TxtTexto.Item(5).Text = 0
                    'TxtTexto.Item(6).Text = 0
                    'TxtTexto.Item(7).Text = 0
                    'TxtTexto.Item(8).Text = 0
                    'TxtTexto.Item(9).Text = 0
                    'TxtTexto.Item(10).Text = 0
                    'TxtTexto.Item(11).Text = 0
                    'TxtTexto.Item(12).Text = 0
                    'TxtTexto.Item(13).Text = 0
                    'TxtTexto.Item(14).Text = 0
                    'TxtTexto.Item(15).Text = 0
                    'TxtTexto.Item(16).Text = 0
                    
            
            
                    'AGREGA UN REGISTRO
                    .AddNew
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = True
                    botones
                    MskFec.Text = Date
                    MskFec.SetFocus
                    TxtTexto.Item(2).Text = GUsuario
                                
                    'BUSCA LOS FACTORES PARA CALCULOS DE HORAS
                    Set RBuscaFactores = Db.OpenRecordset("Select * From EmpleadosFactores")
                        If RBuscaFactores.RecordCount > 0 Then
                                VFHorasDiurnas = RBuscaFactores(0)
                                VFHorasNocturnas = RBuscaFactores(1)
                                VFPorcentajeDiurnas = RBuscaFactores(2)
                                VFPorcentajeNocturnas = RBuscaFactores(3)
                        End If
                    
            'EDITAR
            ElseIf Index = 1 Then
                    'EDITA EL REGISTRO
                    .Edit
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = True
                    botones
                    TxtTexto.Item(0).SetFocus
                    TxtTexto.Item(2).Text = GUsuario
            'GRABAR
            ElseIf Index = 2 Then
                            'EMPLEADO
                                Set RBuscaSueldoBase = Db.OpenRecordset("Select SueldoBase From Empleados Where Codigo = '" & TxtTexto.Item(3).Text & "'")
                                    If RBuscaSueldoBase.RecordCount > 0 Then
                                        VSueldoBase = RBuscaSueldoBase!SueldoBase
                                        VValorHoraLaboradaDiurna = Format(((VSueldoBase / VDiasMes) / VFHorasDiurnas), "#,###,##0.00")
                                        VValorHoraLaboradaNocturna = Format(((VSueldoBase / VDiasMes) / VFHorasNocturnas), "#,###,##0.00")
                                        VValorHoraExtraDiurna = Format((VValorHoraLaboradaDiurna * VFPorcentajeDiurnas), "#,###,##0.00")
                                        VValorHoraExtraNocturna = Format((VValorHoraLaboradaNocturna * VFPorcentajeNocturnas), "#,###,##0.00")
                                    Else
                                        VSueldoBase = "0"
                                        VValorHoraLaboradaDiurna = "0"
                                        VValorHoraLaboradaNocturna = "0"
                                        VValorHoraExtraDiurna = "0"
                                        VValorHoraExtraNocturna = "0"
                                    End If
                                    'MONTO LABORADAS DIURNAS
                                    TxtTexto.Item(5).Text = Format(TxtTexto.Item(4).Text * VValorHoraLaboradaDiurna, "#,###,##0.00")
                                    'MONTO LABORADAS NOCTURNAS
                                    TxtTexto.Item(7).Text = Format(TxtTexto.Item(6).Text * VValorHoraLaboradaNocturna, "#,###,##0.00")
                                    'MONTO HORAS EXTRAS DIURNAS
                                    TxtTexto.Item(9).Text = Format(TxtTexto.Item(8).Text * VValorHoraExtraDiurna, "#,###,##0.00")
                                    'MONTO HORAS EXTRAS NOCTURNAS
                                    TxtTexto.Item(11).Text = Format(TxtTexto.Item(10).Text * VValorHoraExtraNocturna, "#,###,##0.00")
                                    'MONTO HORAS EXTRAS DOBLES
                                    TxtTexto.Item(13).Text = Format((TxtTexto.Item(12).Text * (VValorHoraExtraDiurna * 2)), "#,###,##0.00")
                            

            
                   
                     'GRABA EL REGISTRO
                     .Update
                    'SI SE DUPLICA LA LLAVE
                     If Err = 3022 Then
                        MsgBox "En Esta Fecha y Turno Y Linea y Empleado Ya Existe", vbOKOnly + vbInformation, "Informacion"
                        TxtTexto.Item(0).SetFocus
                        Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                     ElseIf Err <> 3022 And Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                     End If
                        Bandera = False
                        botones
                        CmdBotones.Item(0).SetFocus
            'CANCELAR
            ElseIf Index = 3 Then
                    'CANCELA LOS CAMBIOS Y DEJA LOS DATOS COMO ESTABAN
                    .CancelUpdate
                    'SI HAY ERRORES
                    If Err <> 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                        Exit Sub
                    End If
                    Bandera = False
                    botones
            'BORRAR
            ElseIf Index = 4 Then
                    mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")
        
                    If mensaje = vbOK Then
                        'BORRA EL REGISTRO
                        DataEmpleados.Recordset.Delete
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                        'SE MUEVE AL ULTIMO REGISTRO
                        DataEmpleados.Recordset.MoveNext
                    End If
                    'SI ESTA EN EL FIN DE ARCHIVO
                    If DataEmpleados.Recordset.EOF Then
                        DataEmpleados.Recordset.MoveLast
                        If Err = 3021 Then
                            mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                        End If
                    End If
            'SALIDA
            ElseIf Index = 5 Then
                    Unload Me
            End If
        End With
End Sub

Private Sub CmdBuscar_Click(Index As Integer)
    With DataEmpleados
        'SELECCIONAR DATOS
        If Index = 0 Then
                'FECHAS
                If OptOpcion.Item(0).Value = True Then
                    .RecordSource = ("Select * from EmpleadosCapturaHoras where Fecha >= #" & DTPFecIni.Value & "# And Fecha <= #" & DTPFecFin.Value & "#")
                'FECHAS Y TURNO
                ElseIf OptOpcion.Item(1).Value = True Then
                    .RecordSource = ("Select * from EmpleadosCapturaHoras where Fecha >= #" & DTPFecIni.Value & "# And Fecha <= #" & DTPFecFin.Value & "# And Turno = '" & Txtbusqueda.Text & "'")
                'FECHAS Y LINEA
                ElseIf OptOpcion.Item(2).Value = True Then
                    .RecordSource = ("Select * from EmpleadosCapturaHoras where Fecha >= #" & DTPFecIni.Value & "# And Fecha <= #" & DTPFecFin.Value & "# And Linea = '" & Txtbusqueda.Text & "'")
                'FECHAS Y EMPLEADO
                ElseIf OptOpcion.Item(3).Value = True Then
                    .RecordSource = ("Select * from EmpleadosCapturaHoras where Fecha >= #" & DTPFecIni.Value & "# And Fecha <= #" & DTPFecFin.Value & "# And Empleado = '" & Txtbusqueda.Text & "'")
                End If
                    
                    
                .Refresh
                DBGridEmpleados.Refresh
        'SELECCIONAR TODOS LOS DATOS
        ElseIf Index = 1 Then
                .RecordSource = "Select * From EmpleadosCapturaHoras"
                .Refresh
                DBGridEmpleados.Refresh
        End If
    End With
        TabEmpleados.Tab = 1
End Sub


Private Sub DBGridConsultas_DblClick()
        If BTurno = True Then
            TxtTexto.Item(0).Text = DBGridConsultas.Columns(0).Text
            TxtTexto.Item(0).SetFocus
        ElseIf BLinea = True Then
            TxtTexto.Item(1).Text = DBGridConsultas.Columns(0).Text
            TxtTexto.Item(1).SetFocus
        ElseIf BEmpleado = True Then
            TxtTexto.Item(3).Text = DBGridConsultas.Columns(0).Text
            TxtTexto.Item(3).SetFocus
        End If
        FrameConsultas.Visible = False
End Sub

Private Sub DBGridConsultas_KeyPress(KeyAscii As Integer)
        If BTurno = True Then
            TxtTexto.Item(0).Text = DBGridConsultas.Columns(0).Text
            TxtTexto.Item(0).SetFocus
        ElseIf BLinea = True Then
            TxtTexto.Item(1).Text = DBGridConsultas.Columns(0).Text
            TxtTexto.Item(1).SetFocus
        ElseIf BEmpleado = True Then
            TxtTexto.Item(3).Text = DBGridConsultas.Columns(0).Text
            TxtTexto.Item(3).SetFocus
        End If
        FrameConsultas.Visible = False
End Sub

Private Sub DBgridempleados_HeadClick(ByVal ColIndex As Integer)
    DataEmpleados.RecordSource = ("Select * from EmpleadosCapturaHoras order by " & DBGridEmpleados.Columns(ColIndex).DataField)
    DataEmpleados.Refresh
    DBGridEmpleados.Refresh
End Sub


Private Sub Form_Load()
        'ASIGNA EL TIPO DE BASE DE DATOS YA QUE PUEDE SER ACCESS 97 O 2000
        DataEmpleados.Connect = GConnect
        DataConsultas.Connect = GConnect
        
        'ASIGNA LA RUTA DONDE SE ENCUENTRA LA BASE DE DATOS
        DataEmpleados.DatabaseName = BasedeDatos
        DataConsultas.DatabaseName = BasedeDatos

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

Private Sub MskFec_LostFocus()
        VDiasMes = UltimoDiaMes(MskFec.Text)
            
End Sub

Private Sub OptOpcion_Click(Index As Integer)
        If Index = 0 Then
            LblBusqueda.Caption = ""
            Txtbusqueda.Visible = False
        Else
            Txtbusqueda.Visible = True
            Txtbusqueda.SetFocus
        End If
        
        If Index = 1 Then
            LblBusqueda.Caption = "Turno"
        ElseIf Index = 2 Then
            LblBusqueda.Caption = "Linea"
        ElseIf Index = 3 Then
            LblBusqueda.Caption = "Empleado"
        End If
        
        
End Sub

Private Sub tabempleados_Click(PreviousTab As Integer)
        If TabEmpleados.Tab = 2 Then
            DTPFecIni.Value = Date
            DTPFecFin.Value = Date
        End If
End Sub

Private Sub TxtConsultas_Change()
    'Empleados
    If BEmpleado = True Then
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select * From Empleados Where Descripcion Like '" & TxtConsultas.Text & "*' Order By Descripcion"
            Else
                DataConsultas.RecordSource = "Select * From Empleados Where Descripcion Like '*" & TxtConsultas.Text & "*' Order By Descripcion"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select * From Empleados Where Codigo Like '" & TxtConsultas.Text & "*' Order By Descripcion"
            Else
                DataConsultas.RecordSource = "Select * From Empleados Where Codigo Like '*" & TxtConsultas.Text & "*' Order By Descripcion"
            End If
        End If
    'LINEA
    ElseIf BLinea = True Then
        If OptDes.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select * From Lineas Where Descrip Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select * From Lineas Where Descrip Like '*" & TxtConsultas.Text & "*'"
            End If
        ElseIf OptCod.Value = True Then
            If OptPalIni.Value = True Then
                DataConsultas.RecordSource = "Select * From Lineas Where Linea Like '" & TxtConsultas.Text & "*'"
            Else
                DataConsultas.RecordSource = "Select * From Lineas Where Linea Like '*" & TxtConsultas.Text & "*'"
            End If
        End If
    End If
    DataConsultas.Refresh
    DBGridConsultas.Refresh

End Sub

Private Sub TxtTexto_Change(Index As Integer)
        If Index = 1 Then
            Set RBuscaLineas = Db.OpenRecordset("Select Descrip From Lineas Where Linea = '" & TxtTexto.Item(1).Text & "'")
                If RBuscaLineas.RecordCount > 0 Then
                    LblLinea.Caption = RBuscaLineas!Descrip
                Else
                    LblLinea.Caption = ""
                End If
        ElseIf Index = 3 Then
            Set RBuscaEmpleado = Db.OpenRecordset("Select Descripcion From Empleados Where Codigo = '" & TxtTexto.Item(3).Text & "'")
                If RBuscaEmpleado.RecordCount > 0 Then
                    LblEmpleado.Caption = RBuscaEmpleado!Descripcion
                Else
                    LblEmpleado.Caption = ""
                End If
        End If
End Sub

Private Sub TxtTexto_DblClick(Index As Integer)
        If Index = 0 Then
                BTurno = True
                BLinea = False
                BEmpleado = False
                DataConsultas.RecordSource = "Select * From Turnos"
        ElseIf Index = 1 Then
                BTurno = False
                BLinea = True
                BEmpleado = False
                DataConsultas.RecordSource = "Select Linea, Descrip From Lineas"
        ElseIf Index = 3 Then
                BTurno = False
                BLinea = False
                BEmpleado = True
                DataConsultas.RecordSource = "Select Codigo, Descripcion From Empleados"
        End If
        
        If Index = 0 Or Index = 1 Or Index = 3 Then
                DataConsultas.Refresh
                DBGridConsultas.Refresh
                DBGridConsultas.Columns(1).Width = "4000"
                FrameConsultas.Visible = True
                TxtConsultas.SetFocus
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
                If Index = 0 Then
                        BTurno = True
                        BLinea = False
                        BEmpleado = False
                        DataConsultas.RecordSource = "Select * From Turnos"
                ElseIf Index = 1 Then
                        BTurno = False
                        BLinea = True
                        BEmpleado = False
                        DataConsultas.RecordSource = "Select Linea, Descrip From Lineas"
                ElseIf Index = 3 Then
                        BTurno = False
                        BLinea = False
                        BEmpleado = True
                        DataConsultas.RecordSource = "Select Codigo, Descripcion From Empleados"
                End If
                
                If Index = 0 Or Index = 1 Or Index = 3 Then
                        DataConsultas.Refresh
                        DBGridConsultas.Refresh
                        DBGridConsultas.Columns(1).Width = "4000"
                        FrameConsultas.Visible = True
                        TxtConsultas.SetFocus
                End If
        End If
End Sub

Private Sub Txttexto_LostFocus(Index As Integer)
        'EMPLEADO
        If Index = 3 Then
            Set RBuscaSueldoBase = Db.OpenRecordset("Select SueldoBase From Empleados Where Codigo = '" & TxtTexto.Item(3).Text & "'")
                If RBuscaSueldoBase.RecordCount > 0 Then
                    VSueldoBase = RBuscaSueldoBase!SueldoBase
                    VValorHoraLaboradaDiurna = ((VSueldoBase / VDiasMes) / VFHorasDiurnas)
                    VValorHoraLaboradaNocturna = ((VSueldoBase / VDiasMes) / VFHorasNocturnas)
                    VValorHoraExtraDiurna = VValorHoraLaboradaDiurna * VFPorcentajeDiurnas
                    VValorHoraExtraNocturna = VValorHoraLaboradaNocturna * VFPorcentajeNocturnas
                Else
                    VSueldoBase = "0"
                    VValorHoraLaboradaDiurna = "0"
                    VValorHoraLaboradaNocturna = "0"
                    VValorHoraExtraDiurna = "0"
                    VValorHoraExtraNocturna = "0"
                End If
        End If
        'MONTO LABORADAS DIURNAS
        If Index = 4 Then
                TxtTexto.Item(5).Text = Format(TxtTexto.Item(4).Text * VValorHoraLaboradaDiurna, "#,###,##0.00")
        End If
        'MONTO LABORADAS NOCTURNAS
        If Index = 6 Then
                TxtTexto.Item(7).Text = Format(TxtTexto.Item(6).Text * VValorHoraLaboradaNocturna, "#,###,##0.00")
        End If
        'MONTO HORAS EXTRAS DIURNAS
        If Index = 8 Then
                TxtTexto.Item(9).Text = Format(TxtTexto.Item(8).Text * VValorHoraExtraDiurna, "#,###,##0.00")
        End If
        'MONTO HORAS EXTRAS NOCTURNAS
        If Index = 10 Then
                TxtTexto.Item(11).Text = Format(TxtTexto.Item(10).Text * VValorHoraExtraNocturna, "#,###,##0.00")
        End If
        'MONTO HORAS EXTRAS DOBLES
        If Index = 12 Then
                TxtTexto.Item(13).Text = Format((TxtTexto.Item(12).Text * (VValorHoraExtraDiurna * 2)), "#,###,##0.00")
        End If
        
End Sub

