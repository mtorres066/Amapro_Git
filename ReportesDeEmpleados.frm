VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReportesDeEmpleados 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes De Empleados"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10185
   Icon            =   "ReportesDeEmpleados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Height          =   5415
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   9975
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   4335
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   9735
         _ExtentX        =   17171
         _ExtentY        =   7646
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
         Left            =   1320
         TabIndex        =   6
         Top             =   600
         Width           =   3975
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   615
         Left            =   9120
         Picture         =   "ReportesDeEmpleados.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblbusqueda 
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
         TabIndex        =   11
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8760
      Picture         =   "ReportesDeEmpleados.frx":24B4
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8760
      Picture         =   "ReportesDeEmpleados.frx":2DE6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   120
      Width           =   1335
   End
   Begin TabDlg.SSTab TabReportes 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9551
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   1058
      BackColor       =   8438015
      TabCaption(0)   =   "Empleados"
      TabPicture(0)   =   "ReportesDeEmpleados.frx":3530
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblEmp"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblEmpDes"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FrameTipos"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "OptEmp(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "OptEmp(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "OptEmp(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "OptEmp(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TxtEmp"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "OptEmp(4)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "OptEmp(5)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Frame2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).ControlCount=   11
      TabCaption(1)   =   "Faltas"
      TabPicture(1)   =   "ReportesDeEmpleados.frx":3E0A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "DTPFalFecFin"
      Tab(1).Control(1)=   "DTPFalFecIni"
      Tab(1).Control(2)=   "FrameFaltas"
      Tab(1).Control(3)=   "OptFal(2)"
      Tab(1).Control(4)=   "OptFal(1)"
      Tab(1).Control(5)=   "OptFal(0)"
      Tab(1).Control(6)=   "TxtFal"
      Tab(1).Control(7)=   "LblFalDes"
      Tab(1).Control(8)=   "LblFecini(1)"
      Tab(1).Control(9)=   "LblFecini(0)"
      Tab(1).Control(10)=   "LblFal"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Cusos"
      TabPicture(2)   =   "ReportesDeEmpleados.frx":4124
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "TxtCur"
      Tab(2).Control(1)=   "DTPCurFecfin"
      Tab(2).Control(2)=   "DTPCurFecIni"
      Tab(2).Control(3)=   "OptCur(2)"
      Tab(2).Control(4)=   "OptCur(1)"
      Tab(2).Control(5)=   "Frame3"
      Tab(2).Control(6)=   "OptCur(0)"
      Tab(2).Control(7)=   "LblCurDes"
      Tab(2).Control(8)=   "LblCur"
      Tab(2).Control(9)=   "LblCurFecFin"
      Tab(2).Control(10)=   "LblCurFecIni"
      Tab(2).ControlCount=   11
      TabCaption(3)   =   "Niños"
      TabPicture(3)   =   "ReportesDeEmpleados.frx":4576
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "TxtHij"
      Tab(3).Control(1)=   "OptHij(0)"
      Tab(3).Control(2)=   "DTPHijFecFin"
      Tab(3).Control(3)=   "DTPHijFecIni"
      Tab(3).Control(4)=   "LblHorFecIni(0)"
      Tab(3).Control(5)=   "LblHorFecFin(1)"
      Tab(3).Control(6)=   "LblHij"
      Tab(3).Control(7)=   "LblHijDes"
      Tab(3).Control(8)=   "LblHor"
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "Habilidades"
      TabPicture(4)   =   "ReportesDeEmpleados.frx":4E50
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "OptHab(2)"
      Tab(4).Control(1)=   "OptHab(1)"
      Tab(4).Control(2)=   "TxtHab"
      Tab(4).Control(3)=   "OptHab(0)"
      Tab(4).Control(4)=   "LblHabDes"
      Tab(4).Control(5)=   "LblHab"
      Tab(4).ControlCount=   6
      TabCaption(5)   =   "Eficiencias"
      TabPicture(5)   =   "ReportesDeEmpleados.frx":6B5A
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "OptEfi(0)"
      Tab(5).Control(1)=   "OptEfi(1)"
      Tab(5).Control(2)=   "OptEfi(2)"
      Tab(5).Control(3)=   "TxtEfi"
      Tab(5).Control(4)=   "DtpEfiFecFin"
      Tab(5).Control(5)=   "DTPEfiFecIni"
      Tab(5).Control(6)=   "Label5"
      Tab(5).Control(7)=   "Label4"
      Tab(5).Control(8)=   "LblEfi"
      Tab(5).Control(9)=   "LblEfiDes"
      Tab(5).ControlCount=   10
      Begin VB.OptionButton OptEfi 
         Caption         =   "Empleado"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   74
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton OptEfi 
         Caption         =   "x Equipo"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   73
         Top             =   1440
         Width           =   1095
      End
      Begin VB.OptionButton OptEfi 
         Caption         =   "x Linea"
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   72
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox TxtEfi 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   69
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox TxtHij 
         Appearance      =   0  'Flat
         Height          =   288
         Left            =   -73080
         TabIndex        =   62
         Top             =   4200
         Width           =   1452
      End
      Begin VB.OptionButton OptHij 
         Caption         =   "Empleado y Fechas De Nacimiento"
         Height          =   195
         Index           =   0
         Left            =   -74160
         TabIndex        =   61
         Top             =   1800
         Width           =   3135
      End
      Begin VB.OptionButton OptHab 
         Caption         =   "Empleados Por Habilidades Del Puesto"
         Height          =   195
         Index           =   2
         Left            =   -74520
         TabIndex        =   60
         Top             =   2160
         Width           =   3375
      End
      Begin VB.OptionButton OptHab 
         Caption         =   "Habilidades Por Empleado"
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   57
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox TxtHab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72240
         TabIndex        =   56
         Top             =   3360
         Width           =   1695
      End
      Begin VB.OptionButton OptHab 
         Caption         =   "Habilidades Por Puesto"
         Height          =   195
         Index           =   0
         Left            =   -74520
         TabIndex        =   55
         Top             =   1200
         Width           =   2175
      End
      Begin VB.TextBox TxtCur 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   52
         Top             =   3360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPCurFecfin 
         Height          =   255
         Left            =   -70200
         TabIndex        =   50
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   37992
      End
      Begin MSComCtl2.DTPicker DTPCurFecIni 
         Height          =   255
         Left            =   -72840
         TabIndex        =   49
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   37992
      End
      Begin VB.OptionButton OptCur 
         Caption         =   "x Equipo"
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   47
         Top             =   1800
         Width           =   1575
      End
      Begin VB.OptionButton OptCur 
         Caption         =   "x Departamento"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   46
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo Reporte"
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
         Height          =   1215
         Left            =   -71400
         TabIndex        =   43
         Top             =   1080
         Width           =   2535
         Begin VB.OptionButton OptCurs 
            Caption         =   "x Equipo"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   45
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton OptCurs 
            Caption         =   "x Departamento"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.OptionButton OptCur 
         Caption         =   "Empleado"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   42
         Top             =   1080
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPFalFecFin 
         Height          =   255
         Left            =   -70320
         TabIndex        =   38
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17104899
         CurrentDate     =   37992
      End
      Begin MSComCtl2.DTPicker DTPFalFecIni 
         Height          =   255
         Left            =   -72840
         TabIndex        =   37
         Top             =   2760
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17104899
         CurrentDate     =   37992
      End
      Begin VB.Frame FrameFaltas 
         Caption         =   "Tipo Reporte"
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
         Height          =   1215
         Left            =   -71400
         TabIndex        =   34
         Top             =   1080
         Width           =   2535
         Begin VB.OptionButton OptFal2 
            Caption         =   "x Equipo"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton OptFal2 
            Caption         =   "x Departamento"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   35
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.OptionButton OptFal 
         Caption         =   "x Equipo"
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   33
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Caption         =   "Estado Empleado"
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
         Height          =   1095
         Left            =   6000
         TabIndex        =   27
         Top             =   1080
         Width           =   1935
         Begin VB.OptionButton OptEmp3 
            Caption         =   "Todos"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   30
            Top             =   840
            Width           =   975
         End
         Begin VB.OptionButton OptEmp3 
            Caption         =   "Baja"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   29
            Top             =   600
            Width           =   975
         End
         Begin VB.OptionButton OptEmp3 
            Caption         =   "Alta"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.OptionButton OptEmp 
         Caption         =   "Puesto"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   26
         Top             =   2880
         Width           =   1695
      End
      Begin VB.OptionButton OptEmp 
         Caption         =   "Escolaridad"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   25
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtEmp 
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   22
         ToolTipText     =   "signo + o doble click para ayuda"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.OptionButton OptEmp 
         Caption         =   "Departamento"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   21
         Top             =   1800
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.OptionButton OptEmp 
         Caption         =   "Equipo"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   20
         Top             =   2160
         Width           =   1695
      End
      Begin VB.OptionButton OptEmp 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   1440
         Width           =   1695
      End
      Begin VB.OptionButton OptEmp 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Frame FrameTipos 
         Caption         =   "Tipo Reporte"
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
         Height          =   2055
         Left            =   3600
         TabIndex        =   15
         Top             =   1080
         Width           =   2295
         Begin VB.OptionButton OptEmp2 
            Caption         =   "Ficha De Empleado"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.OptionButton OptEmp2 
            Caption         =   "Listado De Empleados"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   31
            Top             =   1080
            Width           =   1935
         End
         Begin VB.OptionButton OptEmp2 
            Caption         =   "X Departamento"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton OptEmp2 
            Caption         =   "X Equipo"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptFal 
         Caption         =   "x Departamento"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   14
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton OptFal 
         Caption         =   "Empleado"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox TxtFal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72840
         TabIndex        =   1
         Top             =   3360
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPHijFecFin 
         Height          =   255
         Left            =   -68160
         TabIndex        =   63
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17104897
         CurrentDate     =   37722
      End
      Begin MSComCtl2.DTPicker DTPHijFecIni 
         Height          =   255
         Left            =   -70560
         TabIndex        =   64
         Top             =   3480
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   17104897
         CurrentDate     =   37722
      End
      Begin MSComCtl2.DTPicker DtpEfiFecFin 
         Height          =   255
         Left            =   -70200
         TabIndex        =   70
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   37992
      End
      Begin MSComCtl2.DTPicker DTPEfiFecIni 
         Height          =   255
         Left            =   -72840
         TabIndex        =   71
         Top             =   2880
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         Format          =   17104897
         CurrentDate     =   37992
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -74040
         TabIndex        =   78
         Top             =   2880
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -71400
         TabIndex        =   77
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label LblEfi 
         Alignment       =   1  'Right Justify
         Caption         =   "Empleado"
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
         Left            =   -74880
         TabIndex        =   76
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label LblEfiDes 
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
         Left            =   -71400
         TabIndex        =   75
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Label LblHorFecIni 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   -71760
         TabIndex        =   68
         Top             =   3480
         Width           =   1065
      End
      Begin VB.Label LblHorFecFin 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   1
         Left            =   -69240
         TabIndex        =   67
         Top             =   3480
         Width           =   990
      End
      Begin VB.Label LblHij 
         Alignment       =   1  'Right Justify
         Caption         =   "Empleado"
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
         Left            =   -74520
         TabIndex        =   66
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label LblHijDes 
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
         Left            =   -71520
         TabIndex        =   65
         Top             =   4200
         Width           =   4575
      End
      Begin VB.Label LblHabDes 
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
         Left            =   -70440
         TabIndex        =   59
         Top             =   3360
         Width           =   3495
      End
      Begin VB.Label LblHab 
         Alignment       =   1  'Right Justify
         Caption         =   "Puesto"
         Height          =   255
         Left            =   -73680
         TabIndex        =   58
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label LblCurDes 
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
         Left            =   -71400
         TabIndex        =   54
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Label LblCur 
         Alignment       =   1  'Right Justify
         Caption         =   "Empleado"
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
         Left            =   -74880
         TabIndex        =   53
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label LblCurFecFin 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -71400
         TabIndex        =   51
         Top             =   2880
         Width           =   1005
      End
      Begin VB.Label LblCurFecIni 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -74040
         TabIndex        =   48
         Top             =   2880
         Width           =   1110
      End
      Begin VB.Label LblFalDes 
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
         Left            =   -71400
         TabIndex        =   41
         Top             =   3360
         Width           =   4335
      End
      Begin VB.Label LblFecini 
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
         Index           =   1
         Left            =   -71400
         TabIndex        =   40
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label LblFecini 
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
         Index           =   0
         Left            =   -74040
         TabIndex        =   39
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label LblEmpDes 
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
         Left            =   3600
         TabIndex        =   24
         Top             =   3720
         Width           =   3855
      End
      Begin VB.Label LblEmp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   840
         TabIndex        =   23
         Top             =   3720
         Width           =   840
         WordWrap        =   -1  'True
      End
      Begin VB.Label LblHor 
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
         Left            =   -74760
         TabIndex        =   12
         Top             =   3840
         Width           =   1335
      End
      Begin VB.Label LblFal 
         Alignment       =   1  'Right Justify
         Caption         =   "Empleado"
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
         Left            =   -74520
         TabIndex        =   2
         Top             =   3360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "ReportesDeEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim VDia As String
Dim VDia2 As String
Dim VMes As String
Dim VMes2 As String
Dim VAño As String
Dim VAño2 As String

Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaEquipo As New ADODB.Recordset
Dim RBuscaDepartamento As New ADODB.Recordset
Dim RBuscaEmpleado As New ADODB.Recordset
Dim RBuscaEscolaridad As New ADODB.Recordset
Dim RBuscaPuesto As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

'EMPLEADOS
Dim BEmpLinea As Boolean
Dim BEmpEquipo As Boolean
Dim BEmpDepartamento As Boolean
Dim BEmpEscolaridad As Boolean
Dim BEmpPuesto As Boolean
Dim BEmpEmpleado As Boolean

'FALTAS
Dim BFalEmpleado As Boolean
Dim BFalEquipo As Boolean
Dim BFalDepartamento As Boolean

'CURSOS
Dim BCurEmpleado As Boolean
Dim BCurEquipo As Boolean
Dim BCurDepartamento As Boolean

'HABILIDADES
Dim BHabPuesto As Boolean
Dim BHabEmpleado As Boolean

'EFICIENCIA
Dim BEfiEmpleado As Boolean
Dim BEfiEquipo As Boolean
Dim BEfiLinea As Boolean




Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11

                                'CrReportes.DiscardSavedData = True
                                GCriteriaReporte = ""

  'DEPARTAMENTOS
  If TabReportes.Tab = 0 Then
                                'SI ELIGE FICHA DE EMPLEADO VALIDA EL CODIGO DE EMPLEADO
                                If OptEmp2.Item(3).Value = True Then
                                    Set RBuscaEmpleado = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBuscaEmpleado, "Select * From Empleados Where Codigo = '" & TxtEmp.Text & "'")
                                        Else
                                            Call Abrir_Recordset(RBuscaEmpleado, "Select * From Empleados Where UPPER(Codigo) = '" & UCase(TxtEmp.Text) & "'")
                                        End If
                                        If RBuscaEmpleado.RecordCount > 0 Then
                                        Else
                                            MsgBox "Codigo Empleado No Existe", vbOKOnly + vbInformation, "Informacion"
                                            MousePointer = 0
                                            Exit Sub
                                        End If
                                End If
                                
                                Empleados
  ElseIf TabReportes.Tab = 1 Then
                                Faltas
  ElseIf TabReportes.Tab = 2 Then
                                Cursos
  ElseIf TabReportes.Tab = 3 Then
                                Hijos
  ElseIf TabReportes.Tab = 4 Then
  
                                If OptHab.Item(1).Value = True Then
                                    Set RBuscaEmpleado = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBuscaEmpleado, "Select * From Empleados Where Codigo = '" & TxtEmp.Text & "'")
                                        Else
                                            Call Abrir_Recordset(RBuscaEmpleado, "Select * From Empleados Where UPPER(Codigo) = '" & UCase(TxtEmp.Text) & "'")
                                        End If
                                        If RBuscaEmpleado.RecordCount > 0 Then
                                        Else
                                            MsgBox "Codigo Empleado No Existe", vbOKOnly + vbInformation, "Informacion"
                                            MousePointer = 0
                                            Exit Sub
                                        End If
                                End If
                                
                                Habilidades
  ElseIf TabReportes.Tab = 5 Then
  
                                If OptEfi.Item(0).Value = True Then
                                    Set RBuscaEmpleado = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBuscaEmpleado, "Select * From Empleados Where Codigo = '" & TxtEfi.Text & "'")
                                        Else
                                            Call Abrir_Recordset(RBuscaEmpleado, "Select * From Empleados Where UPPER(Codigo) = '" & UCase(TxtEfi.Text) & "'")
                                        End If
                                        If RBuscaEmpleado.RecordCount > 0 Then
                                        Else
                                            MsgBox "Codigo Empleado No Existe", vbOKOnly + vbInformation, "Informacion"
                                            MousePointer = 0
                                            Exit Sub
                                        End If
                                End If
                                
                                Eficiencias
  End If
  
'********************************************************************************************************************************************************************************************
                MousePointer = 0
                FrmReporte.Show
            
                
                If Err > 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                End If
  
End Sub

Private Sub CmdSalida_Click()
    Unload Me
    
End Sub


Private Sub CmdSalir_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
        'TAB DE EMPLEADOS
        If BEmpEmpleado = True Or BEmpEquipo = True Or BEmpDepartamento = True Or BEmpEscolaridad = True Or BEmpPuesto = True Then
                TxtEmp.Text = DBGridBusqueda.Columns(0)
                TxtEmp.SetFocus
        
        'FALTAS
        ElseIf BFalEquipo = True Or BFalDepartamento = True Or BFalEmpleado = True Then
                TxtFal.Text = DBGridBusqueda.Columns(0)
                TxtFal.SetFocus
        'CURSOS
        ElseIf BCurEquipo = True Or BCurDepartamento = True Or BCurEmpleado = True Then
                TxtCur.Text = DBGridBusqueda.Columns(0)
                TxtCur.SetFocus
        'PUESTOS
        ElseIf BHabEmpleado = True Or BHabPuesto = True Then
                TxtHab.Text = DBGridBusqueda.Columns(0)
                TxtHab.SetFocus
        'EFICIENCIA
        ElseIf BEfiEmpleado = True Or BEfiEquipo = True Or BEfiLinea = True Then
                TxtEfi.Text = DBGridBusqueda.Columns(0)
                TxtEfi.SetFocus
        End If
                FrameBusqueda.Visible = False
        
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        'TAB DE EMPLEADOS
        If BEmpEmpleado = True Or BEmpEquipo = True Or BEmpDepartamento = True Or BEmpEscolaridad = True Or BEmpPuesto = True Then
                TxtEmp.Text = DBGridBusqueda.Columns(0)
                TxtEmp.SetFocus
        'TAB DE HORAS
        
        'FALTAS
        ElseIf BFalEquipo = True Or BFalDepartamento = True Or BFalEmpleado = True Then
                TxtFal.Text = DBGridBusqueda.Columns(0)
                TxtFal.SetFocus
        'CURSOS
        ElseIf BCurEquipo = True Or BCurDepartamento = True Or BCurEmpleado = True Then
                TxtCur.Text = DBGridBusqueda.Columns(0)
                TxtCur.SetFocus
        'PUESTOS
        ElseIf BHabEmpleado = True Or BHabPuesto = True Then
                TxtHab.Text = DBGridBusqueda.Columns(0)
                TxtHab.SetFocus
        'EFICIENCIA
        ElseIf BEfiEmpleado = True Or BEfiEquipo = True Or BEfiLinea = True Then
                TxtEfi.Text = DBGridBusqueda.Columns(0)
                TxtEfi.SetFocus
        End If
                FrameBusqueda.Visible = False
        End If

End Sub

Private Sub OptBusqueda_Click(Index As Integer)
        If Index = 0 Then
            LblBusqueda.Caption = "Descripcion"
        ElseIf Index = 1 Then
            LblBusqueda.Caption = "Codigo"
        End If
            TxtBusqueda.SetFocus
        
End Sub

Private Sub OptCur_Click(Index As Integer)
        If Index = 0 Then
            LblCur.Caption = "Empleado"
        ElseIf Index = 1 Then
            LblCur.Caption = "Departamento"
        ElseIf Index = 2 Then
            LblCur.Caption = "Equipo"
        End If
        TxtCur.SetFocus

End Sub

Private Sub OptEfi_Click(Index As Integer)
        If Index = 0 Then
            LblEfi.Caption = "Empleado"
        ElseIf Index = 1 Then
            LblEfi.Caption = "Equipo"
        ElseIf Index = 2 Then
            LblEfi.Caption = "Linea"
        End If
End Sub

Private Sub OptEmp_Click(Index As Integer)
        If Index = 0 Then
            LblEmp.Caption = "Codigo"
        ElseIf Index = 1 Then
            LblEmp.Caption = "Descripcion"
        ElseIf Index = 2 Then
            LblEmp.Caption = "Departamento"
        ElseIf Index = 3 Then
            LblEmp.Caption = "Equipo"
        End If
            TxtEmp.SetFocus
            
        'SI ELIGE OPCION DE CODIGO
        If Index = 0 Then
            'SI TIENE ACCESO A LA CONFIGURACION SI PUEDE SACAR LA FICHA
            If GConfiguracionEmpleados = True Then
                OptEmp2.Item(3).Visible = True
            Else
                OptEmp2.Item(3).Visible = False
            End If
        Else
            OptEmp2.Item(3).Visible = False
            OptEmp2.Item(1).Value = True
        End If
        
        
End Sub

Private Sub OptEmp2_Click(Index As Integer)
            TxtEmp.SetFocus
End Sub

Private Sub OptEmp3_Click(Index As Integer)
    TxtEmp.SetFocus
End Sub

Private Sub OptFal_Click(Index As Integer)
        If Index = 0 Then
            LblFal.Caption = "Empleado"
        ElseIf Index = 1 Then
            LblFal.Caption = "Departamento"
        ElseIf Index = 2 Then
            LblFal.Caption = "Equipo"
        End If
        TxtFal.SetFocus
End Sub

Private Sub OptFal2_Click(Index As Integer)
        TxtFal.SetFocus
End Sub

Private Sub OptHab_Click(Index As Integer)
        If Index = 0 Then
            LblHab.Caption = "Puesto"
        ElseIf Index = 1 Then
            LblHab.Caption = "Empleado"
        ElseIf Index = 2 Then
            LblHab.Caption = "Puesto"
        End If
        
        TxtHab.SetFocus
End Sub


Private Sub TabReportes_Click(PreviousTab As Integer)
        
                If TabReportes.Tab = 0 Then
                        OptEmp.Item(0).Value = True
                ElseIf TabReportes.Tab = 1 Then
                        OptFal.Item(0).Value = True
                        DTPFalFecIni.Value = Date
                        DTPFalFecFin.Value = Date
                ElseIf TabReportes.Tab = 2 Then
                        OptCur.Item(0).Value = True
                        DTPCurFecIni.Value = Date
                        DTPCurFecfin.Value = Date
                ElseIf TabReportes.Tab = 3 Then
                        OptHij.Item(0).Value = True
                        DTPHijFecIni.Value = Date
                        DTPHijFecFin.Value = Date
                ElseIf TabReportes.Tab = 4 Then
                        OptHab.Item(0).Value = True
                ElseIf TabReportes.Tab = 5 Then
                        OptEfi.Item(0).Value = True
                        DTPEfiFecIni.Value = Date
                        DtpEfiFecFin.Value = Date
                End If

End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'BUSCA EQUIPO
            If (BEmpEquipo = True Or BFalEquipo = True Or BCurEquipo = True Or BEfiEquipo = True) Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                            
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos Where Codigo Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosGrupos Where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            End If
            
             'BUSCA DEPARTAMENTO
            If (BEmpDepartamento = True Or BFalDepartamento = True Or BCurDepartamento = True) Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos Where Codigo Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosDepartamentos Where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        
                    End If
            End If
            
            'BUSCA ESCOLARIDADES
            If BEmpEscolaridad = True Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad Where Codigo Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from EmpleadosEscolaridad Where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                        
                    End If
            End If
            
            'BUSCA PUESTOS
            If (BEmpPuesto = True Or BHabPuesto = True) Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos Where CodigoPuesto Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion from EmpleadosPuestos Where UPPER(CodigoPuesto) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            End If
            
            
             'BUSCA EMPLEADO
            If BEmpEmpleado = True Or BFalEmpleado = True Or BCurEmpleado = True Or BHabEmpleado = True Or BEfiEmpleado = True Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Empleados Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Empleados Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Empleados Where Codigo Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Empleados Where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            End If
            
             'BUSCA LINEA
            If (BEmpLinea = True Or BEfiLinea = True) Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where Descrip Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where Linea Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            End If
            
                                
                                Set DBGridBusqueda.DataSource = RBusqueda
                                DBGridBusqueda.Columns(1).Width = "4000"
                            

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

Public Sub Faltas()
                        VDia = Day(DTPFalFecIni.Value)
                        VMes = Month(DTPFalFecIni.Value)
                        VAño = Year(DTPFalFecIni.Value)
                        VDia2 = Day(DTPFalFecFin.Value)
                        VMes2 = Month(DTPFalFecFin.Value)
                        VAño2 = Year(DTPFalFecFin.Value)
                                                
                                                
                       
                       'EMPLEADOS
                       If OptFal.Item(0).Value = True Then
                            GCriteriaReporte = "{EmpleadosCapturaFaltas.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EmpleadosCapturaFaltas.Empleado} Like '" & TxtFal.Text & "*'"
                            GTituloReporte = "Desde " & DTPFalFecIni.Value & " Hasta " & DTPFalFecFin.Value & " Por Empleado " & TxtFal.Text
                       'DEPARTAMENTO
                       ElseIf OptFal.Item(1).Value = True Then
                            GCriteriaReporte = "{EmpleadosCapturaFaltas.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EmpleadosCapturaFaltas.Empleado} = {Empleados.Codigo} And {Empleados.Departamento} Like '" & TxtFal.Text & "*'"
                            GTituloReporte = "Desde " & DTPFalFecIni.Value & " Hasta " & DTPFalFecFin.Value & " Por Departamento " & TxtFal.Text
                       'EQUIPO
                       ElseIf OptFal.Item(2).Value = True Then
                            GCriteriaReporte = "{EmpleadosCapturaFaltas.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EmpleadosCapturaFaltas.Empleado} = {Empleados.Codigo} And {Empleados.Grupo} Like '" & TxtFal.Text & "*'"
                            GTituloReporte = "Desde " & DTPFalFecIni.Value & " Hasta " & DTPFalFecFin.Value & " Por Equipo " & TxtFal.Text
                       End If
                       
                       'TIPO DE REPORTE
                       If OptFal2.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosFaltasDepartamentos.rpt"
                            Else
                                GNombreReporte = "EmpleadosFaltasDepartamentosO.rpt"
                            End If
                       ElseIf OptFal2.Item(1).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosFaltasEquipos.rpt"
                            Else
                                GNombreReporte = "EmpleadosFaltasEquiposO.rpt"
                            End If
                       End If
                       
End Sub





Private Sub TxtCur_Change()
        'EQUIPO
        If OptCur.Item(2).Value = True Then
                Set RBuscaEquipo = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtCur.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where UPPER(Codigo) = '" & UCase(TxtCur.Text) & "'")
                    End If
                    If RBuscaEquipo.RecordCount > 0 Then
                        LblCurDes.Caption = RBuscaEquipo!Descripcion
                    Else
                        LblCurDes.Caption = ""
                    End If
        'DEPARTAMENTO
        ElseIf OptCur.Item(1).Value = True Then
                Set RBuscaDepartamento = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where Codigo = '" & TxtCur.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where UPPER(Codigo) = '" & UCase(TxtCur.Text) & "'")
                    End If
                    If RBuscaDepartamento.RecordCount > 0 Then
                        LblCurDes.Caption = RBuscaDepartamento!Descripcion
                    Else
                        LblCurDes.Caption = ""
                    End If
        'EMPLEADO
        ElseIf OptCur.Item(0).Value = True Then
                Set RBuscaEmpleado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtCur.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where UPPER(Codigo) = '" & UCase(TxtCur.Text) & "'")
                    End If
                    If RBuscaEmpleado.RecordCount > 0 Then
                        LblCurDes.Caption = RBuscaEmpleado!Descripcion
                    Else
                        LblCurDes.Caption = ""
                    End If
       End If

End Sub

Private Sub TxtCur_DblClick()
        Set RBusqueda = New ADODB.Recordset
        'EMPLEADOS
        If OptCur.Item(0).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = True
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
        'DEPARTAMENTO
        ElseIf OptCur.Item(1).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = True
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosDepartamentos")
        'EQUIPO
        ElseIf OptCur.Item(2).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = True
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
        End If
        
        If (OptCur.Item(0).Value = True Or OptCur.Item(1).Value = True Or OptCur.Item(2).Value = True) Then
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If

End Sub

Private Sub TxtCur_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        
        If KeyAscii = 43 Then
                'EMPLEADOS
                If OptCur.Item(0).Value = True Then
                            BEmpEquipo = False
                            BEmpDepartamento = False
                            BEmpLinea = False
                            BEmpEmpleado = False
                            BEmpEscolaridad = False
                            BEmpPuesto = False
                            BFalEmpleado = False
                            BFalEquipo = False
                            BFalDepartamento = False
                            BCurEmpleado = True
                            BCurEquipo = False
                            BCurDepartamento = False
                            BHabEmpleado = False
                            BHabPuesto = False
                            BEfiEmpleado = False
                            BEfiEquipo = False
                            BEfiLinea = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
                'DEPARTAMENTO
                ElseIf OptCur.Item(1).Value = True Then
                            BEmpEquipo = False
                            BEmpDepartamento = False
                            BEmpLinea = False
                            BEmpEmpleado = False
                            BEmpEscolaridad = False
                            BEmpPuesto = False
                            BFalEmpleado = False
                            BFalEquipo = False
                            BFalDepartamento = False
                            BCurEmpleado = False
                            BCurEquipo = False
                            BCurDepartamento = True
                            BHabEmpleado = False
                            BHabPuesto = False
                            BEfiEmpleado = False
                            BEfiEquipo = False
                            BEfiLinea = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosDepartamentos")
                'EQUIPO
                ElseIf OptCur.Item(2).Value = True Then
                            BEmpEquipo = False
                            BEmpDepartamento = False
                            BEmpLinea = False
                            BEmpEmpleado = False
                            BEmpEscolaridad = False
                            BEmpPuesto = False
                            BFalEmpleado = False
                            BFalEquipo = False
                            BFalDepartamento = False
                            BCurEmpleado = False
                            BCurEquipo = True
                            BCurDepartamento = False
                            BHabEmpleado = False
                            BHabPuesto = False
                            BEfiEmpleado = False
                            BEfiEquipo = False
                            BEfiLinea = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
                End If
                
                
                If (OptCur.Item(0).Value = True Or OptCur.Item(1).Value = True Or OptCur.Item(2).Value = True) Then
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                End If
                
        End If

End Sub

Private Sub TxtEfi_Change()
        'EQUIPO
        If OptEfi.Item(1).Value = True Then
                Set RBuscaEquipo = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtEfi.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where UPPER(Codigo) = '" & UCase(TxtEfi.Text) & "'")
                    End If
                    If RBuscaEquipo.RecordCount > 0 Then
                        LblEfiDes.Caption = RBuscaEquipo!Descripcion
                    Else
                        LblEfiDes.Caption = ""
                    End If
        'LINEAS
        ElseIf OptEfi.Item(2).Value = True Then
                Set RBuscaLinea = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtEfi.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtEfi.Text) & "'")
                    End If
                    If RBuscaLinea.RecordCount > 0 Then
                        LblEfiDes.Caption = RBuscaLinea!Descrip
                    Else
                        LblEfiDes.Caption = ""
                    End If
        'EMPLEADO
        ElseIf OptEfi.Item(0).Value = True Then
                Set RBuscaEmpleado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtEfi.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where UPPER(Codigo) = '" & UCase(TxtEfi.Text) & "'")
                    End If
                    If RBuscaEmpleado.RecordCount > 0 Then
                        LblEfiDes.Caption = RBuscaEmpleado!Descripcion
                    Else
                        LblEfiDes.Caption = ""
                    End If
       End If

End Sub

Private Sub TxtEfi_DblClick()
        Set RBusqueda = New ADODB.Recordset
        'EMPLEADO
        If OptEfi.Item(0).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = True
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
        
        'EQUIPO
        ElseIf OptEfi.Item(1).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = True
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
        'LINEA
        ElseIf OptEfi.Item(2).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = True
                    Call Abrir_Recordset(RBusqueda, "Select linea, Descrip From Lineas")
        End If
        
        If (OptEfi.Item(0).Value = True Or OptEfi.Item(1).Value = True Or OptEfi.Item(2).Value = True) Then
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If

End Sub

Private Sub TxtEfi_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                    Set RBusqueda = New ADODB.Recordset
                    'EMPLEADO
                    If OptEfi.Item(0).Value = True Then
                                BEmpEquipo = False
                                BEmpDepartamento = False
                                BEmpLinea = False
                                BEmpEmpleado = False
                                BEmpEscolaridad = False
                                BEmpPuesto = False
                                BFalEmpleado = False
                                BFalEquipo = False
                                BFalDepartamento = False
                                BCurEmpleado = False
                                BCurEquipo = False
                                BCurDepartamento = False
                                BHabEmpleado = False
                                BHabPuesto = False
                                BEfiEmpleado = True
                                BEfiEquipo = False
                                BEfiLinea = False
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
                    
                    'EQUIPO
                    ElseIf OptEfi.Item(1).Value = True Then
                                BEmpEquipo = False
                                BEmpDepartamento = False
                                BEmpLinea = False
                                BEmpEmpleado = False
                                BEmpEscolaridad = False
                                BEmpPuesto = False
                                BFalEmpleado = False
                                BFalEquipo = False
                                BFalDepartamento = False
                                BCurEmpleado = False
                                BCurEquipo = False
                                BCurDepartamento = False
                                BHabEmpleado = False
                                BHabPuesto = False
                                BEfiEmpleado = False
                                BEfiEquipo = True
                                BEfiLinea = False
                                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
                    'LINEA
                    ElseIf OptEfi.Item(2).Value = True Then
                                BEmpEquipo = False
                                BEmpDepartamento = False
                                BEmpLinea = False
                                BEmpEmpleado = False
                                BEmpEscolaridad = False
                                BEmpPuesto = False
                                BFalEmpleado = False
                                BFalEquipo = False
                                BFalDepartamento = False
                                BCurEmpleado = False
                                BCurEquipo = False
                                BCurDepartamento = False
                                BHabEmpleado = False
                                BHabPuesto = False
                                BEfiEmpleado = False
                                BEfiEquipo = False
                                BEfiLinea = True
                                Call Abrir_Recordset(RBusqueda, "Select linea, Descrip From Lineas")
                    End If
                    
                    If (OptEfi.Item(0).Value = True Or OptEfi.Item(1).Value = True Or OptEfi.Item(2).Value = True) Then
                                Set DBGridBusqueda.DataSource = RBusqueda
                                DBGridBusqueda.Columns(1).Width = "4000"
                                FrameBusqueda.Visible = True
                                TxtBusqueda.SetFocus
                    End If
        End If

End Sub

Private Sub TxtEmp_Change()
        'EQUIPO
        If OptEmp.Item(3).Value = True Then
                Set RBuscaEquipo = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtEmp.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where UPPER(Codigo) = '" & UCase(TxtEmp.Text) & "'")
                    End If
                    If RBuscaEquipo.RecordCount > 0 Then
                        LblEmpDes.Caption = RBuscaEquipo!Descripcion
                    Else
                        LblEmpDes.Caption = ""
                    End If
        'DEPARTAMENTO
        ElseIf OptEmp.Item(2).Value = True Then
                Set RBuscaDepartamento = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where Codigo = '" & TxtEmp.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where UPPER(Codigo) = '" & UCase(TxtEmp.Text) & "'")
                    End If
                    If RBuscaDepartamento.RecordCount > 0 Then
                        LblEmpDes.Caption = RBuscaDepartamento!Descripcion
                    Else
                        LblEmpDes.Caption = ""
                    End If
        'EMPLEADO
        ElseIf OptEmp.Item(0).Value = True Then
                Set RBuscaEmpleado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtEmp.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where UPPER(Codigo) = '" & UCase(TxtEmp.Text) & "'")
                    End If
                    If RBuscaEmpleado.RecordCount > 0 Then
                        LblEmpDes.Caption = RBuscaEmpleado!Descripcion
                    Else
                        LblEmpDes.Caption = ""
                    End If
        'ESCOLARIDAD
        ElseIf OptEmp.Item(4).Value = True Then
                Set RBuscaEscolaridad = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEscolaridad, "Select Descripcion From EmpleadosEscolaridad Where Codigo = '" & TxtEmp.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEscolaridad, "Select Descripcion From EmpleadosEscolaridad Where UPPER(Codigo) = '" & UCase(TxtEmp.Text) & "'")
                    End If
                    If RBuscaEscolaridad.RecordCount > 0 Then
                        LblEmpDes.Caption = RBuscaEscolaridad!Descripcion
                    Else
                        LblEmpDes.Caption = ""
                    End If
        'PUESTO
        ElseIf OptEmp.Item(5).Value = True Then
                Set RBuscaPuesto = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaPuesto, "Select Descripcion From EmpleadosPuestos Where CodigoPuesto = '" & TxtEmp.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaPuesto, "Select Descripcion From EmpleadosPuestos Where UPPER(CodigoPuesto) = '" & UCase(TxtEmp.Text) & "'")
                    End If
                    If RBuscaPuesto.RecordCount > 0 Then
                        LblEmpDes.Caption = RBuscaPuesto!Descripcion
                    Else
                        LblEmpDes.Caption = ""
                    End If
        
        
        End If
End Sub

Private Sub TxtEmp_DblClick()
        Set RBusqueda = New ADODB.Recordset
        'EQUIPO
        If OptEmp.Item(3).Value = True Then
                    BEmpEquipo = True
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
        'EMPLEADO
        ElseIf OptEmp.Item(0).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = True
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
        'DEPARTAMENTO
        ElseIf OptEmp.Item(2).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = True
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosDepartamentos")
        'ESCOLARIDAD
        ElseIf OptEmp.Item(4).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = True
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosEscolaridad")
        'PUESTO
        ElseIf OptEmp.Item(5).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = True
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion From EmpleadosPuestos")
                
        End If
        
        If (OptEmp.Item(0).Value = True Or OptEmp.Item(2).Value = True Or OptEmp.Item(3).Value = True Or OptEmp.Item(4).Value = True Or OptEmp.Item(5).Value = True) Then
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If
End Sub

Private Sub TxtEmp_GotFocus()
        TxtEmp.SelStart = 0
        TxtEmp.SelLength = Len(TxtEmp.Text)
End Sub

Private Sub TxtEmp_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                'EQUIPO
                If OptEmp.Item(3).Value = True Then
                    BEmpEquipo = True
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
                'EMPLEADO
                ElseIf OptEmp.Item(0).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = True
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
        
                'DEPARTAMENTO
                ElseIf OptEmp.Item(2).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = True
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosDepartamentos")
                'ESCOLARIDAD
                ElseIf OptEmp.Item(4).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = True
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosEscolaridad")
                'PUESTO
                ElseIf OptEmp.Item(5).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = True
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion From EmpleadosPuestos")
                End If
                
                If (OptEmp.Item(0).Value = True Or OptEmp.Item(2).Value = True Or OptEmp.Item(3).Value = True Or OptEmp.Item(4).Value = True Or OptEmp.Item(5).Value = True) Then
                            
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                End If
        End If
End Sub

Private Sub TxtFal_Change()
        'EQUIPO
        If OptFal.Item(2).Value = True Then
                Set RBuscaEquipo = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtFal.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where UPPER(Codigo) = '" & UCase(TxtFal.Text) & "'")
                    End If
                    If RBuscaEquipo.RecordCount > 0 Then
                        LblFalDes.Caption = RBuscaEquipo!Descripcion
                    Else
                        LblFalDes.Caption = ""
                    End If
        'DEPARTAMENTO
        ElseIf OptFal.Item(1).Value = True Then
                Set RBuscaDepartamento = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where Codigo = '" & TxtFal.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaDepartamento, "Select Descripcion From EmpleadosDepartamentos Where UPPER(Codigo) = '" & UCase(TxtFal.Text) & "'")
                    End If
                    If RBuscaDepartamento.RecordCount > 0 Then
                        LblFalDes.Caption = RBuscaDepartamento!Descripcion
                    Else
                        LblFalDes.Caption = ""
                    End If
        'EMPLEADO
        ElseIf OptFal.Item(0).Value = True Then
                Set RBuscaEmpleado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtFal.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where UPPER(Codigo) = '" & UCase(TxtFal.Text) & "'")
                    End If
                    If RBuscaEmpleado.RecordCount > 0 Then
                        LblFalDes.Caption = RBuscaEmpleado!Descripcion
                    Else
                        LblFalDes.Caption = ""
                    End If
       End If

End Sub

Private Sub TxtFal_DblClick()
        Set RBusqueda = New ADODB.Recordset
        'EMPLEADO
        If OptFal.Item(0).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = True
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
        
        'DEPARTAMENTO
        ElseIf OptFal.Item(1).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = True
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosDepartamentos")
        'EQUIPO
        ElseIf OptFal.Item(2).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = True
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
        End If
        
        If (OptFal.Item(0).Value = True Or OptFal.Item(1).Value = True Or OptFal.Item(2).Value = True) Then
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If

End Sub

Private Sub TxtFal_GotFocus()
        TxtFal.SelStart = 0
        TxtFal.SelLength = Len(TxtFal.Text)
End Sub

Private Sub TxtFal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                'EMPLEADO
                If OptFal.Item(0).Value = True Then
                            BEmpEquipo = False
                            BEmpDepartamento = False
                            BEmpLinea = False
                            BEmpEmpleado = False
                            BEmpEscolaridad = False
                            BEmpPuesto = False
                            BFalEmpleado = True
                            BFalEquipo = False
                            BFalDepartamento = False
                            BCurEmpleado = False
                            BCurEquipo = False
                            BCurDepartamento = False
                            BHabEmpleado = False
                            BHabPuesto = False
                            BEfiEmpleado = False
                            BEfiEquipo = False
                            BEfiLinea = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
                'DEPARTAMENTO
                ElseIf OptFal.Item(1).Value = True Then
                            BEmpEquipo = False
                            BEmpDepartamento = False
                            BEmpLinea = False
                            BEmpEmpleado = False
                            BEmpEscolaridad = False
                            BEmpPuesto = False
                            BFalEmpleado = False
                            BFalEquipo = False
                            BFalDepartamento = True
                            BCurEmpleado = False
                            BCurEquipo = False
                            BCurDepartamento = False
                            BHabEmpleado = False
                            BHabPuesto = False
                            BEfiEmpleado = False
                            BEfiEquipo = False
                            BEfiLinea = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosDepartamentos")
                'EQUIPO
                ElseIf OptFal.Item(2).Value = True Then
                            BEmpEquipo = False
                            BEmpDepartamento = False
                            BEmpLinea = False
                            BEmpEmpleado = False
                            BEmpEscolaridad = False
                            BEmpPuesto = False
                            BFalEmpleado = False
                            BFalEquipo = True
                            BFalDepartamento = False
                            BCurEmpleado = False
                            BCurEquipo = False
                            BCurDepartamento = False
                            BHabEmpleado = False
                            BHabPuesto = False
                            BEfiEmpleado = False
                            BEfiEquipo = False
                            BEfiLinea = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
                End If
                
                If (OptFal.Item(0).Value = True Or OptFal.Item(1).Value = True Or OptFal.Item(2).Value = True) Then
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                End If
        End If

End Sub



Private Sub TxtHab_Change()
        'EMPLEADO
        If OptHab.Item(1).Value = True Then
                Set RBuscaEmpleado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtHab.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where UPPER(Codigo) = '" & UCase(TxtHab.Text) & "'")
                    End If
                    If RBuscaEmpleado.RecordCount > 0 Then
                        LblHabDes.Caption = RBuscaEmpleado!Descripcion
                    Else
                        LblHabDes.Caption = ""
                    End If
        'PUESTO
        ElseIf (OptHab.Item(0).Value = True Or OptHab.Item(2).Value = True) Then
                Set RBuscaPuesto = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaPuesto, "Select Descripcion From EmpleadosPuestos Where CodigoPuesto = '" & TxtHab.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaPuesto, "Select Descripcion From EmpleadosPuestos Where UPPER(CodigoPuesto) = '" & UCase(TxtHab.Text) & "'")
                    End If
                    If RBuscaPuesto.RecordCount > 0 Then
                        LblHabDes.Caption = RBuscaPuesto!Descripcion
                    Else
                        LblHabDes.Caption = ""
                    End If
        End If
End Sub

Private Sub TxtHab_DblClick()
        Set RBusqueda = New ADODB.Recordset
        'PUESTO
        If (OptHab.Item(0).Value = True Or OptHab.Item(2).Value = True) Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = False
                    BHabPuesto = True
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion From EmpleadosPuestos")
        'EMPLEADO
        ElseIf OptHab.Item(1).Value = True Then
                    BEmpEquipo = False
                    BEmpDepartamento = False
                    BEmpLinea = False
                    BEmpEmpleado = False
                    BEmpEscolaridad = False
                    BEmpPuesto = False
                    BFalEmpleado = False
                    BFalEquipo = False
                    BFalDepartamento = False
                    BCurEmpleado = False
                    BCurEquipo = False
                    BCurDepartamento = False
                    BHabEmpleado = True
                    BHabPuesto = False
                    BEfiEmpleado = False
                    BEfiEquipo = False
                    BEfiLinea = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
        End If
        
        If (OptHab.Item(0).Value = True Or OptHab.Item(1).Value = True Or OptHab.Item(2).Value = True) Then
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If

End Sub

Private Sub TxtHab_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        
        If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                'PUESTO
                If (OptHab.Item(0).Value = True Or OptHab.Item(2).Value = True) Then
                            BEmpEquipo = False
                            BEmpDepartamento = False
                            BEmpLinea = False
                            BEmpEmpleado = False
                            BEmpEscolaridad = False
                            BEmpPuesto = False
                            BFalEmpleado = False
                            BFalEquipo = False
                            BFalDepartamento = False
                            BCurEmpleado = False
                            BCurEquipo = False
                            BCurDepartamento = False
                            BHabEmpleado = False
                            BHabPuesto = True
                            BEfiEmpleado = False
                            BEfiEquipo = False
                            BEfiLinea = False
                            Call Abrir_Recordset(RBusqueda, "Select CodigoPuesto, Descripcion From EmpleadosPuestos")
                'EMPLEADO
                ElseIf OptHab.Item(1).Value = True Then
                            BEmpEquipo = False
                            BEmpDepartamento = False
                            BEmpLinea = False
                            BEmpEmpleado = False
                            BEmpEscolaridad = False
                            BEmpPuesto = False
                            BFalEmpleado = False
                            BFalEquipo = False
                            BFalDepartamento = False
                            BCurEmpleado = False
                            BCurEquipo = False
                            BCurDepartamento = False
                            BHabEmpleado = True
                            BHabPuesto = False
                            BEfiEmpleado = False
                            BEfiEquipo = False
                            BEfiLinea = False
                            Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Empleados")
                End If
                
                If (OptHab.Item(0).Value = True Or OptHab.Item(1).Value = True Or OptHab.Item(2).Value = True) Then
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                End If
        End If

End Sub


Public Sub Cursos()
                        VDia = Day(DTPCurFecIni.Value)
                        VMes = Month(DTPCurFecIni.Value)
                        VAño = Year(DTPCurFecIni.Value)
                        VDia2 = Day(DTPCurFecfin.Value)
                        VMes2 = Month(DTPCurFecfin.Value)
                        VAño2 = Year(DTPCurFecfin.Value)
                                                
                       'EMPLEADOS
                       If OptCur.Item(0).Value = True Then
                            GCriteriaReporte = "{EmpleadosCapturaCursos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EmpleadosCapturaCursos.Empleado} Like '" & TxtCur.Text & "*'"
                       'DEPARTAMENTO
                       ElseIf OptCur.Item(1).Value = True Then
                            GCriteriaReporte = "{EmpleadosCapturaCursos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EmpleadosCapturaCursos.Empleado} = {Empleados.Codigo} And {Empleados.Departamento} Like '" & TxtCur.Text & "*'"
                       'EQUIPO
                       ElseIf OptCur.Item(2).Value = True Then
                            GCriteriaReporte = "{EmpleadosCapturaCursos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EmpleadosCapturaCursos.Empleado} = {Empleados.Codigo} And {Empleados.Grupo} Like '" & TxtCur.Text & "*'"
                       End If
                       
                       'TIPO DE REPORTE
                       If OptCur.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosCursosEquipos.rpt"
                            Else
                                GNombreReporte = "EmpleadosCursosEquiposO.rpt"
                            End If
                       ElseIf OptCur.Item(1).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosCursosEquipos.rpt"
                            Else
                                GNombreReporte = "EmpleadosCursosEquiposO.rpt"
                            End If
                       End If
End Sub

Public Sub Empleados()
                        'CrReportes.DiscardSavedData = True
                        GCriteriaReporte = ""
                            
                        
                        'SI PIDE LA FICHA DEL EMPLEADO
                        If OptEmp2.Item(3).Value = True Then
                            GCriteriaReporte = "{Empleados.Codigo} = '" & TxtEmp.Text & "'"
                        Else
                            'CODIGO
                            If OptEmp.Item(0).Value = True Then
                                 GCriteriaReporte = "{Empleados.Codigo} Like '" & TxtEmp.Text & "*'"
                            'DESCRIPCION
                            ElseIf OptEmp.Item(1).Value = True Then
                                 GCriteriaReporte = "{Empleados.Descripcion} Like '" & TxtEmp.Text & "*'"
                            'DEPARTAMENTO
                            ElseIf OptEmp.Item(2).Value = True Then
                                 GCriteriaReporte = "{Empleados.Departamento} Like '" & TxtEmp.Text & "*'"
                            'EQUIPO
                            ElseIf OptEmp.Item(3).Value = True Then
                                 GCriteriaReporte = "{Empleados.Grupo} Like '" & TxtEmp.Text & "*'"
                             'ESCOLARIDAD
                            ElseIf OptEmp.Item(4).Value = True Then
                                 GCriteriaReporte = "{Empleados.Escolaridad} Like '" & TxtEmp.Text & "*'"
                             'PUESTO
                            ElseIf OptEmp.Item(5).Value = True Then
                                 GCriteriaReporte = "{Empleados.Puesto} Like '" & TxtEmp.Text & "*'"
                            End If
                        End If
                              
                       'AGRUPADO POR EQUIPO
                       If OptEmp2.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosxEquipo.rpt"
                            Else
                                GNombreReporte = "EmpleadosxEquipoO.rpt"
                            End If
                       'AGRUPADO POR DEPARTAMENTO
                       ElseIf OptEmp2.Item(1).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosxDepartamento.rpt"
                            Else
                                GNombreReporte = "EmpleadosxDepartamentoO.rpt"
                            End If
                       'LISTADO DE EMPLEADOS
                       ElseIf OptEmp2.Item(2).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosListado.rpt"
                            Else
                                GNombreReporte = "EmpleadosListadoO.rpt"
                            End If
                       'FICHA DE EMPLEADO
                       ElseIf OptEmp2.Item(3).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosFicha.rpt"
                            Else
                                GNombreReporte = "EmpleadosFichaO.rpt"
                            End If
                            'CrReportes.SubreportToChange = "Hijos"
                            'CrReportes.ConnectionString = "pwd=metal"
                            'CrReportes.SubreportToChange = "Faltas"
                            'CrReportes.ConnectionString = "pwd=metal"
                            'CrReportes.SubreportToChange = "Cusros"
                            'CrReportes.ConnectionString = "pwd=metal"
                            'CrReportes.SubreportToChange = "Aumentos"
                            'CrReportes.ConnectionString = "pwd=metal"
                                                   
                       End If
                       
                       'SI ELIGE FICHA DE EMPLEADO
                       If OptEmp2.Item(3).Value = True Then
                       Else
                                'ESTADO DEL EMPLEADO
                                'ALTA
                                If OptEmp3.Item(0).Value = True Then
                                     GCriteriaReporte = GCriteriaReporte & " And {Empleados.Estado} = 'ALTA'"
                                'BAJA
                                ElseIf OptEmp3.Item(1).Value = True Then
                                     GCriteriaReporte = GCriteriaReporte & " And {Empleados.Estado} = 'BAJA'"
                                'TODOS
                                ElseIf OptEmp3.Item(2).Value = True Then
                                
                                End If
                        End If
                       
End Sub

Public Sub Habilidades()
            'HABILIDADES POR PUESTO
            If OptHab.Item(0).Value = True Then
                    GCriteriaReporte = "{EmpleadosPuestos.CodigoPuesto} Like '" & TxtHab.Text & "*'"
            'HABILIDADES POR EMPLEADO
            ElseIf OptHab.Item(1).Value = True Then
                    GCriteriaReporte = "{Empleados.Codigo} Like '" & TxtHab.Text & "*'"
            'EMPLEADOS POR HABILIDADES DEL PUESTO
            ElseIf OptHab.Item(2).Value = True Then
                    GCriteriaReporte = "{EmpleadosPuestos.CodigoPuesto} = '" & TxtHab.Text & "'"
            End If
            
                        
            If OptHab.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "EmpleadosHabilidadesDelPuesto.rpt"
                    Else
                        GNombreReporte = "EmpleadosHabilidadesDelPuestoO.rpt"
                    End If
            ElseIf OptHab.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "EmpleadosHabilidadesDelEmpleado.rpt"
                    Else
                        GNombreReporte = "EmpleadosHabilidadesDelEmpleadoO.rpt"
                    End If
            ElseIf OptHab.Item(2).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "EmpleadosHabilidadesPorPuesto.rpt"
                    Else
                        GNombreReporte = "EmpleadosHabilidadesPorPuestoO.rpt"
                    End If
            End If
            
                       
End Sub

Public Sub Hijos()

                        VDia = Day(DTPHijFecIni.Value)
                        VMes = Month(DTPHijFecIni.Value)
                        VAño = Year(DTPHijFecIni.Value)
                        VDia2 = Day(DTPHijFecFin.Value)
                        VMes2 = Month(DTPHijFecFin.Value)
                        VAño2 = Year(DTPHijFecFin.Value)
                                                
                       'EMPLEADOS
                        If OptHij.Item(0).Value = True Then
                            GTituloReporte = "Por Fechas De Nacimiento Desde " & DTPHijFecIni.Value & " Hasta " & DTPHijFecFin.Value
                            GCriteriaReporte = "{EmpleadosHijos.FechaNacimiento} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EmpleadosHijos.Codigo} = {Empleados.Codigo} And {Empleados.Estado} = 'ALTA' And {Empleados.Codigo} Like '" & TxtHij.Text & "*'"
                        End If
                        
                        If OptHij.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    GNombreReporte = "EmpleadosHijos.rpt"
                                Else 'ORACLE
                                    GNombreReporte = "EmpleadosHijosO.rpt"
                                End If
                        End If
            

End Sub

Private Sub TxtHij_Change()
        'LINEA
        If OptHij.Item(0).Value = True Then
                Set RBuscaEmpleado = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where Codigo = '" & TxtHij.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaEmpleado, "Select Descripcion From Empleados Where UPPER(Codigo) = '" & UCase(TxtHij.Text) & "'")
                    End If
                    If RBuscaEmpleado.RecordCount > 0 Then
                        LblHijDes.Caption = RBuscaEmpleado!Descripcion
                    Else
                        LblHijDes.Caption = ""
                    End If
        End If

End Sub

Private Sub TxtHij_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
                
        

End Sub

Public Sub Eficiencias()
                        VDia = Day(DTPEfiFecIni.Value)
                        VMes = Month(DTPEfiFecIni.Value)
                        VAño = Year(DTPEfiFecIni.Value)
                        VDia2 = Day(DtpEfiFecFin.Value)
                        VMes2 = Month(DtpEfiFecFin.Value)
                        VAño2 = Year(DtpEfiFecFin.Value)
                                                
                       'EMPLEADO
                       If OptEfi.Item(0).Value = True Then
                            GTituloReporte = "Desde " & DTPEfiFecIni.Value & " Hasta " & DtpEfiFecFin.Value & " Por Empleado & " & TxtEfi.Text & " " & LblEfiDes.Caption
                            GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoCapturaParos.Documento} = {DetalleEmpleados.Documento} And {DetalleEmpleados.Empleado} Like '" & TxtEfi.Text & "*'"
                       'EQUIPO
                       ElseIf OptEfi.Item(1).Value = True Then
                            GTituloReporte = "Desde " & DTPEfiFecIni.Value & " Hasta " & DtpEfiFecFin.Value & " Por Equipo & " & TxtEfi.Text & " " & LblEfiDes.Caption
                            GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoCapturaParos.Documento} = {DetalleEmpleados.Documento} And {EncabezadoCapturaParos.Grupo} Like '" & TxtEfi.Text & "*'"
                       'MAQUINA
                       ElseIf OptEfi.Item(2).Value = True Then
                            GTituloReporte = "Desde " & DTPEfiFecIni.Value & " Hasta " & DtpEfiFecFin.Value & " Por Linea & " & TxtEfi.Text & " " & LblEfiDes.Caption
                            GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoCapturaParos.Documento} = {DetalleEmpleados.Documento} And {EncabezadoCapturaParos.Linea} Like '" & TxtEfi.Text & "*'"
                       End If
                       
                       'TIPO DE REPORTE
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "EmpleadosEficienciasxEmpleado.rpt"
                            Else
                                GNombreReporte = "EmpleadosEficienciasxEmpleadoO.rpt"
                            End If
                       

End Sub
