VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form EmpleadosProduccion 
   BackColor       =   &H000000FF&
   Caption         =   "Empleados de Produccion"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12015
   Icon            =   "EmpleadosProduccion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   12015
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6975
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   12000
      _ExtentX        =   21167
      _ExtentY        =   12303
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Vista Individual"
      TabPicture(0)   =   "EmpleadosProduccion.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FrameEmpleados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Vista General"
      TabPicture(1)   =   "EmpleadosProduccion.frx":0BE4
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DBGrid1"
      Tab(1).ControlCount=   1
      Begin VB.Frame FrameEmpleados 
         Caption         =   "Datos Empleados de Produccion"
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
         TabIndex        =   9
         Top             =   720
         Width           =   11835
         Begin VB.TextBox TxtSol 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "SoldadoraT1L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   38
            ToolTipText     =   " "
            Top             =   840
            Width           =   3000
         End
         Begin VB.TextBox TxtPal 
            Appearance      =   0  'Flat
            DataField       =   "PaletizadoraT1L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   37
            Top             =   1200
            Width           =   3000
         End
         Begin VB.TextBox TxtInsCal 
            Appearance      =   0  'Flat
            DataField       =   "InspectordeCalidadT1L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   36
            Top             =   1560
            Width           =   3000
         End
         Begin VB.TextBox TxtSupPro 
            Appearance      =   0  'Flat
            DataField       =   "SupervisorProduccionT1L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   35
            Top             =   1920
            Width           =   3000
         End
         Begin VB.TextBox TxtCer 
            Appearance      =   0  'Flat
            DataField       =   "CerradoraT1L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   34
            Top             =   2280
            Width           =   3000
         End
         Begin VB.TextBox TxtEmp 
            Appearance      =   0  'Flat
            DataField       =   "EmpacadorT1L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   33
            Top             =   2640
            Width           =   3000
         End
         Begin VB.TextBox TxtMecLin 
            Appearance      =   0  'Flat
            DataField       =   "MecanicodeLineaT1L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   31
            Top             =   3000
            Width           =   3000
         End
         Begin VB.TextBox TxtSol2 
            Appearance      =   0  'Flat
            DataField       =   "SoldadoraT2L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   30
            Top             =   840
            Width           =   3000
         End
         Begin VB.TextBox TxtPal2 
            Appearance      =   0  'Flat
            DataField       =   "PaletizadoraT2L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   29
            Top             =   1200
            Width           =   3000
         End
         Begin VB.TextBox TxtInsCal2 
            Appearance      =   0  'Flat
            DataField       =   "InspectordeCalidadT2L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   28
            Top             =   1560
            Width           =   3000
         End
         Begin VB.TextBox TxtSupPro2 
            Appearance      =   0  'Flat
            DataField       =   "SupervisorProduccionT2L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   27
            Top             =   1920
            Width           =   3000
         End
         Begin VB.TextBox TxtCer2 
            Appearance      =   0  'Flat
            DataField       =   "CerradoraT2L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   26
            Top             =   2280
            Width           =   3000
         End
         Begin VB.TextBox TxtEmp2 
            Appearance      =   0  'Flat
            DataField       =   "EmpacadorT2L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   25
            Top             =   2640
            Width           =   3000
         End
         Begin VB.TextBox TxtMecLin2 
            Appearance      =   0  'Flat
            DataField       =   "MecanicodeLineaT2L1"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   24
            Top             =   3000
            Width           =   3000
         End
         Begin VB.TextBox TxtOpeSolT1L2 
            Appearance      =   0  'Flat
            DataField       =   "SoldadoraT1L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   23
            Top             =   3480
            Width           =   3000
         End
         Begin VB.TextBox TxtOpePalT1L2 
            Appearance      =   0  'Flat
            DataField       =   "PaletizadoraT1L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   22
            Top             =   3840
            Width           =   3000
         End
         Begin VB.TextBox TxtInsCalT1L2 
            Appearance      =   0  'Flat
            DataField       =   "InspectordeCalidadT1L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   21
            Top             =   4200
            Width           =   3000
         End
         Begin VB.TextBox TxtSupProT1L2 
            Appearance      =   0  'Flat
            DataField       =   "SupervisorProduccionT1L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   20
            Top             =   4560
            Width           =   3000
         End
         Begin VB.TextBox TxtOpeCerT1L2 
            Appearance      =   0  'Flat
            DataField       =   "CerradoraT1L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   19
            Top             =   4920
            Width           =   3000
         End
         Begin VB.TextBox TxtEmpT1L2 
            Appearance      =   0  'Flat
            DataField       =   "EmpacadorT1L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   18
            Top             =   5280
            Width           =   3000
         End
         Begin VB.TextBox TxtMecT1L2 
            Appearance      =   0  'Flat
            DataField       =   "MecanicodeLineaT1L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   2640
            MaxLength       =   30
            TabIndex        =   17
            Top             =   5640
            Width           =   3000
         End
         Begin VB.TextBox TxtOpeSolT2L2 
            Appearance      =   0  'Flat
            DataField       =   "SoldadoraT2L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   16
            Top             =   3480
            Width           =   3000
         End
         Begin VB.TextBox TxtOpePalT2L2 
            Appearance      =   0  'Flat
            DataField       =   "PaletizadoraT2L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   15
            Top             =   3840
            Width           =   3000
         End
         Begin VB.TextBox TxtInsCalT2L2 
            Appearance      =   0  'Flat
            DataField       =   "InspectordeCalidadT2L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   14
            Top             =   4200
            Width           =   3000
         End
         Begin VB.TextBox TxtSupProT2L2 
            Appearance      =   0  'Flat
            DataField       =   "SupervisorProduccionT2L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   13
            Top             =   4560
            Width           =   3000
         End
         Begin VB.TextBox TxtOpeCerT2L2 
            Appearance      =   0  'Flat
            DataField       =   "CerradoraT2L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   12
            Top             =   4920
            Width           =   3000
         End
         Begin VB.TextBox TxtEmpT2L2 
            Appearance      =   0  'Flat
            DataField       =   "EmpacadorT2L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   11
            Top             =   5280
            Width           =   3000
         End
         Begin VB.TextBox TxtMecT2L2 
            Appearance      =   0  'Flat
            DataField       =   "MecanicodeLineaT2L2"
            DataSource      =   "DataEmpleados"
            Height          =   285
            Left            =   8640
            MaxLength       =   30
            TabIndex        =   10
            Top             =   5640
            Width           =   3000
         End
         Begin MSMask.MaskEdBox MskFecEmpPro 
            DataField       =   "FechaEmpPro"
            DataSource      =   "DataEmpleados"
            Height          =   255
            Left            =   2280
            TabIndex        =   32
            Top             =   360
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            Format          =   "dd/mm/yyyy"
            PromptChar      =   "_"
         End
         Begin VB.Label Label1 
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
            Left            =   1560
            TabIndex        =   69
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Operador Soldadora L1"
            Height          =   255
            Left            =   240
            TabIndex        =   68
            Top             =   840
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "Operador Paletizadora L1"
            Height          =   255
            Left            =   240
            TabIndex        =   67
            Top             =   1200
            Width           =   2055
         End
         Begin VB.Label Label4 
            Caption         =   "Inspector de Calidad L1"
            Height          =   255
            Left            =   240
            TabIndex        =   66
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label5 
            Caption         =   "Supervisor De Produccion L1"
            Height          =   255
            Left            =   240
            TabIndex        =   65
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label Label6 
            Caption         =   "Operador Cerradora L1"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   2280
            Width           =   2055
         End
         Begin VB.Label Label7 
            Caption         =   "Empacador L1"
            Height          =   255
            Left            =   240
            TabIndex        =   63
            Top             =   2640
            Width           =   2055
         End
         Begin VB.Label Label8 
            Caption         =   "Mecanico de Linea L1"
            Height          =   255
            Left            =   240
            TabIndex        =   62
            Top             =   3000
            Width           =   2055
         End
         Begin VB.Label Label9 
            Caption         =   "Operador Soldadora L1"
            Height          =   255
            Left            =   6000
            TabIndex        =   61
            Top             =   840
            Width           =   2055
         End
         Begin VB.Label Label10 
            Caption         =   "Operador Paletizadora L1"
            Height          =   255
            Left            =   6000
            TabIndex        =   60
            Top             =   1200
            Width           =   2175
         End
         Begin VB.Label Label11 
            Caption         =   "Inspector de Calidad L1"
            Height          =   255
            Left            =   6000
            TabIndex        =   59
            Top             =   1560
            Width           =   2055
         End
         Begin VB.Label Label12 
            Caption         =   "Supervisor de Produccion L1"
            Height          =   255
            Left            =   6000
            TabIndex        =   58
            Top             =   1920
            Width           =   2655
         End
         Begin VB.Label Label13 
            Caption         =   "Operador Cerradora L1"
            Height          =   255
            Left            =   6000
            TabIndex        =   57
            Top             =   2280
            Width           =   1935
         End
         Begin VB.Label Label14 
            Caption         =   "Empacador L1"
            Height          =   255
            Left            =   6000
            TabIndex        =   56
            Top             =   2640
            Width           =   1575
         End
         Begin VB.Label Label15 
            Caption         =   "Mecanico de Linea L1"
            Height          =   255
            Left            =   6000
            TabIndex        =   55
            Top             =   3000
            Width           =   1815
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   2
            Height          =   5295
            Left            =   120
            Top             =   720
            Width           =   5655
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   2
            Height          =   5295
            Left            =   5880
            Top             =   720
            Width           =   5895
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Turno 1 Grupo 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3600
            TabIndex        =   54
            Top             =   360
            Width           =   2325
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            Caption         =   "Turno 2 Grupo 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   9360
            TabIndex        =   53
            Top             =   360
            Width           =   2340
         End
         Begin VB.Label Label18 
            Caption         =   "Operador Soldadora L2"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label19 
            Caption         =   "Operador Paletizadora L2"
            Height          =   255
            Left            =   240
            TabIndex        =   51
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label20 
            Caption         =   "Inspector de Calidad L2"
            Height          =   255
            Left            =   240
            TabIndex        =   50
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label Label21 
            Caption         =   "Supervisor de Produccion L2"
            Height          =   255
            Left            =   240
            TabIndex        =   49
            Top             =   4560
            Width           =   2295
         End
         Begin VB.Label Label22 
            Caption         =   "Operador Cerradora L2"
            Height          =   255
            Left            =   240
            TabIndex        =   48
            Top             =   4920
            Width           =   2055
         End
         Begin VB.Label Label23 
            Caption         =   "Empacador L2"
            Height          =   255
            Left            =   240
            TabIndex        =   47
            Top             =   5280
            Width           =   1935
         End
         Begin VB.Label Label24 
            Caption         =   "Mecanico de Linea L2"
            Height          =   255
            Left            =   240
            TabIndex        =   46
            Top             =   5640
            Width           =   2055
         End
         Begin VB.Label Label25 
            Caption         =   "Operador Soldadora L2"
            Height          =   255
            Left            =   6000
            TabIndex        =   45
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label26 
            Caption         =   "Operador Paletizadora L2"
            Height          =   255
            Left            =   6000
            TabIndex        =   44
            Top             =   3840
            Width           =   1815
         End
         Begin VB.Label Label27 
            Caption         =   "Inspector de Calidad L2"
            Height          =   255
            Left            =   6000
            TabIndex        =   43
            Top             =   4200
            Width           =   1935
         End
         Begin VB.Label Label28 
            Caption         =   "Supervisor de Produccion L2"
            Height          =   255
            Left            =   6000
            TabIndex        =   42
            Top             =   4560
            Width           =   2175
         End
         Begin VB.Label Label29 
            Caption         =   "Operador Cerradora L2"
            Height          =   255
            Left            =   6000
            TabIndex        =   41
            Top             =   4920
            Width           =   1935
         End
         Begin VB.Label Label30 
            Caption         =   "Empacador L2"
            Height          =   255
            Left            =   6000
            TabIndex        =   40
            Top             =   5280
            Width           =   1935
         End
         Begin VB.Label Label31 
            Caption         =   "Mecanico de Linea L2"
            Height          =   255
            Left            =   6000
            TabIndex        =   39
            Top             =   5640
            Width           =   1935
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   120
            X2              =   5760
            Y1              =   3360
            Y2              =   3360
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   5880
            X2              =   11760
            Y1              =   3360
            Y2              =   3360
         End
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "EmpleadosProduccion.frx":1036
         Height          =   6105
         Left            =   -74880
         OleObjectBlob   =   "EmpleadosProduccion.frx":1052
         TabIndex        =   70
         Top             =   720
         Width           =   11745
      End
   End
   Begin VB.CommandButton CmdActualizar 
      Caption         =   "Actuali&zar"
      Height          =   800
      Left            =   8880
      Picture         =   "EmpleadosProduccion.frx":493D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7080
      Width           =   1400
   End
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "&Buscar"
      Height          =   800
      Left            =   7440
      Picture         =   "EmpleadosProduccion.frx":4C47
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7080
      Width           =   1400
   End
   Begin VB.Data DataEmpleados 
      BackColor       =   &H80000014&
      Caption         =   "Empleados"
      Connect         =   "Access"
      DatabaseName    =   "C:\Cucho\visualbasic\MetalEnvases\MetalEnvases.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EmpleadosProduccion"
      Top             =   8040
      Width           =   11865
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      Height          =   800
      Left            =   10320
      MouseIcon       =   "EmpleadosProduccion.frx":5089
      Picture         =   "EmpleadosProduccion.frx":54CB
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   " "
      Top             =   7080
      Width           =   1400
   End
   Begin VB.CommandButton CmdBorrar 
      Caption         =   "B&orrar"
      Height          =   800
      Left            =   6000
      MouseIcon       =   "EmpleadosProduccion.frx":590D
      Picture         =   "EmpleadosProduccion.frx":5D4F
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   " "
      Top             =   7080
      Width           =   1400
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   4560
      MouseIcon       =   "EmpleadosProduccion.frx":6281
      Picture         =   "EmpleadosProduccion.frx":66C3
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   " "
      Top             =   7080
      Width           =   1400
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   800
      Left            =   3120
      MouseIcon       =   "EmpleadosProduccion.frx":6BF5
      Picture         =   "EmpleadosProduccion.frx":7037
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " "
      Top             =   7080
      Width           =   1400
   End
   Begin VB.CommandButton CmdEditar 
      Caption         =   "&Editar"
      Height          =   800
      Left            =   1680
      MouseIcon       =   "EmpleadosProduccion.frx":7569
      Picture         =   "EmpleadosProduccion.frx":79AB
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " "
      Top             =   7080
      Width           =   1400
   End
   Begin VB.CommandButton CmdAgregar 
      Caption         =   "&Agregar"
      Height          =   800
      Left            =   240
      MouseIcon       =   "EmpleadosProduccion.frx":7EDD
      Picture         =   "EmpleadosProduccion.frx":831F
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " "
      Top             =   7080
      Width           =   1400
   End
End
Attribute VB_Name = "EmpleadosProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Bandera As Boolean
Dim mensaje As String
Dim buscar As String

Dim RBarniz As Recordset
Dim REmpleados As Recordset

Sub botones()
    If Bandera = True Then
         FrameEmpleados.Enabled = True
         CmdAgregar.Enabled = False
         CmdGrabar.Enabled = True
         CmdEditar.Enabled = False
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdActualizar.Enabled = False
         CmdSalida.Enabled = False
         MskFecEmpPro.SetFocus
         DataEmpleados.Visible = False
         
         DBGrid1.Visible = False
    Else
         FrameEmpleados.Enabled = False
         CmdAgregar.Enabled = True
         CmdGrabar.Enabled = False
         CmdEditar.Enabled = True
         CmdBorrar.Enabled = True
         CmdBuscar.Enabled = True
         CmdActualizar.Enabled = True
         CmdCancelar.Enabled = False
         CmdSalida.Enabled = True
         DataEmpleados.Visible = True
         
         DBGrid1.Visible = True
    End If
End Sub

Private Sub CboTurno_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub CmdActualizar_Click()
    DataEmpleados.RecordSource = ("Select * From EmpleadosProduccion")
    DataEmpleados.Refresh
    DBGrid1.Refresh
End Sub

Private Sub CmdAgregar_Click()
        
        Bandera = True
        botones
        DataEmpleados.Recordset.AddNew
        
        'SELECCIONA LOS ULTIMOS EMPLEADOS
        Set REmpleados = Db.OpenRecordset("Select * From EmpleadosProduccion")
        If REmpleados.RecordCount > 0 Then
            
            REmpleados.MoveLast
            
            'TURNO 1
            If Not IsNull(REmpleados(1)) Then
                TxtSol.Text = REmpleados(1)
            End If
            If Not IsNull(REmpleados(2)) Then
                TxtPal.Text = REmpleados(2)
            End If
            If Not IsNull(REmpleados(3)) Then
                TxtInsCal.Text = REmpleados(3)
            End If
            If Not IsNull(REmpleados(4)) Then
                TxtSupPro.Text = REmpleados(4)
            End If
            If Not IsNull(REmpleados(5)) Then
                TxtCer.Text = REmpleados(5)
            End If
            If Not IsNull(REmpleados(6)) Then
                TxtEmp.Text = REmpleados(6)
            End If
            If Not IsNull(REmpleados(7)) Then
                TxtMecLin.Text = REmpleados(7)
            End If
            
            'LINEA 2
            
            If Not IsNull(REmpleados(8)) Then
                TxtOpeSolT1L2.Text = REmpleados(8)
            End If
            If Not IsNull(REmpleados(9)) Then
                TxtOpePalT1L2.Text = REmpleados(9)
            End If
            If Not IsNull(REmpleados(10)) Then
                TxtInsCalT1L2.Text = REmpleados(10)
            End If
            If Not IsNull(REmpleados(11)) Then
                TxtSupProT1L2.Text = REmpleados(11)
            End If
            If Not IsNull(REmpleados(12)) Then
                TxtOpeCerT1L2.Text = REmpleados(12)
            End If
            If Not IsNull(REmpleados(13)) Then
                TxtEmpT1L2.Text = REmpleados(13)
            End If
            If Not IsNull(REmpleados(14)) Then
                TxtMecT1L2.Text = REmpleados(14)
            End If
            
            
            
            If Not IsNull(REmpleados(15)) Then
                TxtSol2.Text = REmpleados(15)
            End If
            If Not IsNull(REmpleados(16)) Then
                TxtPal2.Text = REmpleados(16)
            End If
            If Not IsNull(REmpleados(17)) Then
                TxtInsCal2.Text = REmpleados(17)
            End If
            If Not IsNull(REmpleados(18)) Then
                TxtSupPro2.Text = REmpleados(18)
            End If
            If Not IsNull(REmpleados(19)) Then
                TxtCer2.Text = REmpleados(19)
            End If
            If Not IsNull(REmpleados(20)) Then
                TxtEmp2.Text = REmpleados(20)
            End If
            If Not IsNull(REmpleados(21)) Then
                TxtMecLin2.Text = REmpleados(21)
            End If
            
            'LINEA 2
            
            If Not IsNull(REmpleados(22)) Then
                TxtOpeSolT2L2.Text = REmpleados(22)
            End If
            If Not IsNull(REmpleados(23)) Then
                TxtOpePalT2L2.Text = REmpleados(23)
            End If
            If Not IsNull(REmpleados(24)) Then
                TxtInsCalT2L2.Text = REmpleados(24)
            End If
            If Not IsNull(REmpleados(25)) Then
                TxtSupProT2L2.Text = REmpleados(25)
            End If
            If Not IsNull(REmpleados(26)) Then
                TxtOpeCerT2L2.Text = REmpleados(26)
            End If
            If Not IsNull(REmpleados(27)) Then
                TxtEmpT2L2.Text = REmpleados(27)
            End If
            If Not IsNull(REmpleados(28)) Then
                TxtMecT2L2.Text = REmpleados(28)
            End If
            
            
            
        End If
               
        
        MskFecEmpPro.SetFocus
        MskFecEmpPro.Text = Format(Date, "dd/mm/yyyy")
        
        
End Sub

Private Sub CmdBorrar_Click()
On Error Resume Next

            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            If mensaje = vbOK Then
                DataEmpleados.Recordset.Delete
                DataEmpleados.Recordset.MoveLast
            End If
  
            If DataEmpleados.Recordset.EOF Then
                DataEmpleados.Recordset.MoveLast
                If Err = 3021 Then
                    mensaje = MsgBox("ya no hay registros para borrar", vbInformation + vbOKOnly, "Informacion")
                End If
            End If
            
            
End Sub


Private Sub CmdBuscar_Click()
        buscar = InputBox("Ingrese Fecha A Buscar", "Busqueda")
        If Not IsDate(buscar) Then
            MsgBox "Fecha Invalida", vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
        
        DataEmpleados.RecordSource = ("Select * from EmpleadosProduccion Where FechaEmpPro = #" & Format(buscar, "mm/dd/yyyy") & "#")
        DataEmpleados.Refresh
        DBGrid1.Refresh
        
End Sub

Private Sub CmdCancelar_Click()
        Bandera = False
        botones
        DataEmpleados.Recordset.CancelUpdate
End Sub

Private Sub CmdEditar_Click()
        Bandera = True
        botones
        DataEmpleados.Recordset.Edit
        MskFecEmpPro.SetFocus
        
End Sub

Private Sub CmdGrabar_Click()
   On Error Resume Next
   
   
   DataEmpleados.Recordset.Update
   
   If Err = 3022 Then
      MsgBox "Con Esta Fecha Ya Existen Datos", vbOKOnly + vbInformation, "Informacion"
      MskFecEmpPro.SetFocus
   ElseIf Err <> 0 Then
      MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
   Else
      Bandera = False
      botones
  End If
      
   
      

End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub DBGrid1_HeadClick(ByVal ColIndex As Integer)
    DataEmpleados.RecordSource = ("Select * from EmpleadosProduccion order by " & DBGrid1.Columns(ColIndex).DataField)
    DataEmpleados.Refresh
    DBGrid1.Refresh
    
End Sub


Private Sub Form_Load()
    DataEmpleados.Connect = GConnect
    DataEmpleados.DatabaseName = BasedeDatos

End Sub



Private Sub mskFecEmpPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If
End Sub

Private Sub TxEmpT1L2_Change()

End Sub

Private Sub TxEmpT1L2_KeyPress(KeyAscii As Integer)

End Sub

Private Sub TxtCer_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtCer2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtEmp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtEmp2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtEmpT2L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtInsCal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub


Private Sub TxtInsCal2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtInsCalT1L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtInsCalT2L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtMecLin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtMecLin2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtMecLinT2L2_Change()

End Sub



Private Sub TxtMecT1L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtMecT2L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtOpeCerT1L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtOpeCerT2L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtOpePalT1L1_Change()

End Sub

Private Sub TxtOpePalT1L1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtOpePalT2L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtOpeSolT1L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtOpeSolT2L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtPal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtPal2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtSol_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtSol2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtSupPro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtSupPro2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtSupProT1L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub

Private Sub TxtSupProT2L2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   SendKeys "{tab}"
End If

End Sub
