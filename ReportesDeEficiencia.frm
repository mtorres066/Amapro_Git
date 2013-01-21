VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReportesDeEficiencia 
   BackColor       =   &H000080FF&
   Caption         =   "Reportes De Eficiencia"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9840
   Icon            =   "ReportesDeEficiencia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBusqueda 
      Height          =   6615
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   9615
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   5535
         Left            =   120
         TabIndex        =   53
         Top             =   960
         Width           =   9375
         _ExtentX        =   16536
         _ExtentY        =   9763
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
         TabIndex        =   52
         Top             =   600
         Width           =   3975
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   1
         Left            =   1560
         TabIndex        =   51
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.CommandButton CmdSalir 
         Height          =   615
         Left            =   8760
         Picture         =   "ReportesDeEficiencia.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   54
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
         TabIndex        =   61
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.CommandButton CmdSalida 
      Caption         =   "&Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8400
      Picture         =   "ReportesDeEficiencia.frx":24B4
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   8400
      Picture         =   "ReportesDeEficiencia.frx":2DE6
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   120
      Width           =   1335
   End
   Begin TabDlg.SSTab TabReportes 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   11668
      _Version        =   393216
      Tab             =   2
      TabHeight       =   1058
      BackColor       =   33023
      TabCaption(0)   =   "Ficha De Paros"
      TabPicture(0)   =   "ReportesDeEficiencia.frx":35CB
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "LblParos"
      Tab(0).Control(1)=   "LblGruPar"
      Tab(0).Control(2)=   "OptPar(0)"
      Tab(0).Control(3)=   "TxtPar"
      Tab(0).Control(4)=   "OptPar(1)"
      Tab(0).Control(5)=   "OptPar(2)"
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Captura De Paros"
      TabPicture(1)   =   "ReportesDeEficiencia.frx":38E5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblfecini"
      Tab(1).Control(1)=   "LblFecFin"
      Tab(1).Control(2)=   "LblCapPar"
      Tab(1).Control(3)=   "LblCapParLin"
      Tab(1).Control(4)=   "OptCapPar(0)"
      Tab(1).Control(5)=   "OptCapPar(2)"
      Tab(1).Control(6)=   "TxtCapPar"
      Tab(1).Control(7)=   "DTPFecIniPar"
      Tab(1).Control(8)=   "DTPFecFinPar"
      Tab(1).Control(9)=   "OptCapPar(1)"
      Tab(1).Control(10)=   "FrameParos"
      Tab(1).Control(11)=   "OptCapPar(3)"
      Tab(1).Control(12)=   "Frame3"
      Tab(1).Control(13)=   "OptCapPar(4)"
      Tab(1).ControlCount=   14
      TabCaption(2)   =   "Eficiencia"
      TabPicture(2)   =   "ReportesDeEficiencia.frx":55EF
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label2(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label2(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "LblEfiLin"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label2(3)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label2(4)"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label2(5)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "LblEfiGru"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "Label2(6)"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "DTPFecEfi"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "TxtLinea"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "FrameTipos"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "DTPFecEfiFin"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "FrameEfiTipRep"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "DTPFecEfiAcu"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "DTPFecEfiFinAcu"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "DTPVentas"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "TxtEfiGru"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "DbGridGrupos"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "DTPInventario"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).ControlCount=   20
      Begin MSComCtl2.DTPicker DTPInventario 
         Height          =   255
         Left            =   6720
         TabIndex        =   32
         Top             =   4320
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51904515
         CurrentDate     =   37914
      End
      Begin MSDataGridLib.DataGrid DbGridGrupos 
         Height          =   3375
         Left            =   2400
         TabIndex        =   67
         Top             =   840
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   5953
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
         ColumnCount     =   3
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
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
               Locked          =   -1  'True
               ColumnWidth     =   464.882
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   3660.095
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   764.787
            EndProperty
         EndProperty
      End
      Begin VB.OptionButton OptCapPar 
         Caption         =   "Fechas Y Equipo"
         Height          =   195
         Index           =   4
         Left            =   -74640
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TxtEfiGru 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   39
         Top             =   6240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPVentas 
         Height          =   255
         Left            =   6720
         TabIndex        =   33
         Top             =   4680
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51904515
         CurrentDate     =   37914
      End
      Begin MSComCtl2.DTPicker DTPFecEfiFinAcu 
         Height          =   255
         Left            =   6720
         TabIndex        =   37
         Top             =   5400
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51904515
         CurrentDate     =   37914
      End
      Begin MSComCtl2.DTPicker DTPFecEfiAcu 
         Height          =   255
         Left            =   3120
         TabIndex        =   36
         Top             =   5400
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51904515
         CurrentDate     =   37914
      End
      Begin VB.Frame Frame3 
         Caption         =   "Opciones Reporte"
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
         Height          =   1095
         Left            =   -72360
         TabIndex        =   62
         Top             =   840
         Width           =   1815
         Begin VB.OptionButton OptFicPar 
            Caption         =   "Todos"
            Height          =   192
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   972
         End
         Begin VB.OptionButton OptFicPar 
            Caption         =   "Primeros 10"
            Height          =   192
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin VB.OptionButton OptPar 
         Caption         =   "Grupo"
         Height          =   195
         Index           =   2
         Left            =   -71400
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton OptCapPar 
         Caption         =   "Fechas Y Grupo"
         Height          =   195
         Index           =   3
         Left            =   -74640
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Frame FrameParos 
         Caption         =   "Tipo De Reporte"
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
         Height          =   3855
         Left            =   -70080
         TabIndex        =   43
         Top             =   720
         Width           =   3135
         Begin VB.OptionButton OptResumenTurno 
            Caption         =   "Resumen x Linea y Turno"
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   1080
            Width           =   2655
         End
         Begin VB.OptionButton OptResumenEquipo 
            Caption         =   "Resumen x Equipo"
            Height          =   195
            Left            =   120
            TabIndex        =   17
            Top             =   1800
            Width           =   2655
         End
         Begin VB.OptionButton OptResumenProduccion 
            Caption         =   "Produccion x Linea Cuadricula"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   3600
            Width           =   2895
         End
         Begin VB.OptionButton OptResumenEficienciaEquipo 
            Caption         =   "Eficiencias x Equipo"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   3120
            Width           =   1935
         End
         Begin VB.OptionButton OptResumenEficienciaLinea 
            Caption         =   "Eficiencias x Linea"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   2880
            Width           =   1935
         End
         Begin VB.OptionButton OptResumenCuadriculaMesPS 
            Caption         =   "Resumen Horas Programadas y Horas De Paros y Grafica"
            Height          =   435
            Left            =   120
            TabIndex        =   18
            Top             =   2280
            Width           =   2895
         End
         Begin VB.OptionButton OptResumenCuadriculaMes 
            Caption         =   "Resumen x Linea Cuadricula y Mes"
            Height          =   195
            Left            =   120
            TabIndex        =   16
            Top             =   1560
            Width           =   2895
         End
         Begin VB.OptionButton OptResumenEmpleado 
            Caption         =   "Resumen x Linea y Empleado"
            Height          =   195
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   2415
         End
         Begin VB.OptionButton OptResumenLinea 
            Caption         =   "Resumen x Linea"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1575
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detalle Reporte Produccion"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   360
            Value           =   -1  'True
            Width           =   2775
         End
      End
      Begin VB.Frame FrameEfiTipRep 
         Caption         =   "Tipo De Reporte"
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
         Height          =   3015
         Left            =   120
         TabIndex        =   48
         Top             =   1560
         Width           =   2175
         Begin VB.OptionButton OptEfiLinResEmp 
            Caption         =   "x Maquina Resumen y Empleados"
            Height          =   435
            Left            =   120
            TabIndex        =   69
            Top             =   840
            Width           =   1815
         End
         Begin VB.OptionButton OptEfiGruResEmp 
            Caption         =   "x Equipos Resumen y Empleados"
            ForeColor       =   &H00FF0000&
            Height          =   435
            Left            =   120
            TabIndex        =   68
            Top             =   1920
            Width           =   1935
         End
         Begin VB.OptionButton OptRepEje 
            Caption         =   "Reporte Ejecutivo"
            ForeColor       =   &H000000FF&
            Height          =   192
            Left            =   120
            TabIndex        =   31
            Top             =   2640
            Width           =   1692
         End
         Begin VB.OptionButton OptEfiGruRes 
            Caption         =   "x Equipos Resumen"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   30
            Top             =   1680
            Width           =   1935
         End
         Begin VB.OptionButton OptEfiGru 
            Caption         =   "x Equipos Detalle"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Left            =   120
            TabIndex        =   29
            Top             =   1440
            Width           =   1695
         End
         Begin VB.OptionButton OptEfiEfiPar 
            Caption         =   "x Maquina Resumen"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton OptEfiEfi 
            Caption         =   "x Maquina Detalle"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.OptionButton OptPar 
         Caption         =   "Tipo De Paro"
         Height          =   195
         Index           =   1
         Left            =   -73080
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton OptCapPar 
         Caption         =   "Fechas Y Linea"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   7
         Top             =   1680
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker DTPFecEfiFin 
         Height          =   255
         Left            =   6720
         TabIndex        =   35
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51904515
         CurrentDate     =   37129
      End
      Begin VB.Frame FrameTipos 
         Caption         =   "Opciones de Reporte"
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
         Height          =   735
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   2175
         Begin VB.OptionButton OptEfi 
            Caption         =   "Grupo"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton OptEfi 
            Caption         =   "Linea"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox TxtLinea 
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
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   38
         ToolTipText     =   "signo + o doble click para ayuda"
         Top             =   5880
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker DTPFecEfi 
         Height          =   255
         Left            =   4080
         TabIndex        =   34
         Top             =   5040
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51904515
         CurrentDate     =   37126
      End
      Begin MSComCtl2.DTPicker DTPFecFinPar 
         Height          =   255
         Left            =   -73080
         TabIndex        =   23
         Top             =   4920
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51904515
         CurrentDate     =   37123
      End
      Begin MSComCtl2.DTPicker DTPFecIniPar 
         Height          =   255
         Left            =   -73080
         TabIndex        =   22
         Top             =   4560
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   51904515
         CurrentDate     =   37123
      End
      Begin VB.TextBox TxtCapPar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73080
         TabIndex        =   24
         Top             =   5280
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptCapPar 
         Caption         =   "Documento"
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   9
         Top             =   2400
         Width           =   1215
      End
      Begin VB.OptionButton OptCapPar 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox TxtPar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -73080
         TabIndex        =   4
         Top             =   2280
         Width           =   1335
      End
      Begin VB.OptionButton OptPar 
         Caption         =   "Codigo Paro"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   1
         Top             =   1320
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Inventario"
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
         Index           =   6
         Left            =   5760
         TabIndex        =   70
         Top             =   4320
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.Label LblEfiGru 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo 2"
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
         TabIndex        =   66
         Top             =   6240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Ventas"
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
         Index           =   5
         Left            =   6000
         TabIndex        =   65
         Top             =   4680
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Acumulado Fecha Final"
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
         Index           =   4
         Left            =   4560
         TabIndex        =   64
         Top             =   5400
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Acumulado Fecha Inicial"
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
         Index           =   3
         Left            =   960
         TabIndex        =   63
         Top             =   5400
         Visible         =   0   'False
         Width           =   2160
      End
      Begin VB.Label LblGruPar 
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
         Left            =   -71640
         TabIndex        =   60
         Top             =   2280
         Width           =   4575
      End
      Begin VB.Label LblCapParLin 
         Alignment       =   1  'Right Justify
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
         Left            =   -71640
         TabIndex        =   59
         Top             =   5280
         Width           =   4695
      End
      Begin VB.Label LblEfiLin 
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
         Left            =   4200
         TabIndex        =   58
         Top             =   5880
         Width           =   3855
      End
      Begin VB.Label Label2 
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
         Index           =   2
         Left            =   5640
         TabIndex        =   57
         Top             =   5040
         Width           =   1005
      End
      Begin VB.Label Label2 
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
         Index           =   1
         Left            =   2880
         TabIndex        =   56
         Top             =   5040
         Width           =   1110
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   1440
         TabIndex        =   55
         Top             =   5880
         Width           =   840
         WordWrap        =   -1  'True
      End
      Begin VB.Label LblCapPar 
         Caption         =   "Documento"
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
         Left            =   -74280
         TabIndex        =   46
         Top             =   5280
         Width           =   975
      End
      Begin VB.Label LblFecFin 
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
         Left            =   -74280
         TabIndex        =   45
         Top             =   4920
         Width           =   1095
      End
      Begin VB.Label lblfecini 
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
         Left            =   -74280
         TabIndex        =   44
         Top             =   4560
         Width           =   1215
      End
      Begin VB.Label LblParos 
         AutoSize        =   -1  'True
         Caption         =   "Codigo De Paro"
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
         Left            =   -74520
         TabIndex        =   42
         Top             =   2280
         Width           =   1350
      End
   End
End
Attribute VB_Name = "ReportesDeEficiencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text 'PARA QUE NO DISTINGA MINUSCULAS DE MAYUSCULAS

Dim VDia As String
Dim VDia2 As String
Dim VMes As String
Dim VMes2 As String
Dim VAño As String
Dim VAño2 As String

Dim RBusqueda As New ADODB.Recordset

'EFICIENCIA
Dim RTiempoProgramadoD As New ADODB.Recordset
Dim VTiempoProgramadoD As Single

Dim RTiempoProgramadoN As New ADODB.Recordset
Dim VTiempoProgramadoN As Single

'PAROS QUE NO AFECTAN LA PRODUCCION
Dim RBuscaParosNoAfectanD As New ADODB.Recordset
Dim VParosND As Single

Dim RBuscaParosNoAfectanN As New ADODB.Recordset
Dim VParosNN As Single

'PARO QUE SI AFECTAN LA PRODUCCION
Dim RBuscaParosSiAfectanD As New ADODB.Recordset
Dim VParosSD As Single

Dim RBuscaParosSiAfectanN As New ADODB.Recordset
Dim VParosSN As Single

Dim VParosCFD As Single
Dim VParosMPD As Single
Dim VParosCFN As Single
Dim VParosMPN As Single

'PRODUCCION
Dim RBuscaProduccionD As New ADODB.Recordset
Dim VProduccionD As Single

Dim RBuscaProduccionN As New ADODB.Recordset
Dim VProduccionN As Single


Dim RGrupos As New ADODB.Recordset

'TIEMPO REAL DE PRODUCCION
Dim VTiempoRealProducidoD As Single
Dim VTiempoRealProducidoN As Single

Dim VHorasProducidasDN As Single
Dim VParosDN As Single

Dim VPCD As Single
Dim VPNCD As Single
Dim VPDD As Single

Dim VPCN As Single
Dim VPNCN As Single
Dim VPDN As Single

Dim RProduccion As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset

Dim VTotalProduccion As Single
Dim VTotalProduccionD As Single
Dim VTotalProduccionN As Single

Dim VVelocidadPromedio As Integer
Dim VVelocidadPromedioD As Integer
Dim VVelocidadPromedioN As Integer

Dim VVelocidadTeoricaLinea As Integer
Dim VVelocidadRealLinea As Integer

'CALCULOS DE EFICIENCIA
Dim VVelocidadTeoricaDia As Integer
Dim VVelocidadTeoricaNoche As Integer
Dim VVelocidadRealDia As Integer
Dim VVelocidadRealNoche As Integer

Dim VFactor1D As Single
Dim VFactor2D As Single
Dim VFactor3D As Single
Dim VFactor4D As Single
Dim VFactor5D As Single

Dim VFactor1N As Single
Dim VFactor2N As Single
Dim VFactor3N As Single
Dim VFactor4N As Single
Dim VFactor5N As Single

Dim VFactor1DN As Single
Dim VFactor2DN As Single
Dim VFactor3DN As Single
Dim VFactor4DN As Single
Dim VFactor5DN As Single

Dim VFactor1TDN As Single
Dim VFactor2TDN As Single
Dim VFactor3TDN As Single
Dim VFactor4TDN As Single
Dim VFactor5TDN As Single

Dim RLineas As New ADODB.Recordset

Dim VEficienciaRealD As Single
Dim VEficienciaRealN As Single

Dim RReporteEficiencia As New ADODB.Recordset

Dim VLinea As String
Dim VFechaInicial As Date
Dim VFechaFinal As Date

Dim RSeleccionaLineas As New ADODB.Recordset

Dim VPorcentajeLinea As Single
Dim VPorcentajeRechazo As Single
Dim VPorcentajeDesperdicio As Single
Dim VPorcentajeRechazoD As Single
Dim VPorcentajeDesperdicioD As Single
Dim VPorcentajeRechazoN As Single
Dim VPorcentajeDesperdicioN As Single

Dim Cont As Integer
'Dim VFactorUno As Double
Dim RBuscaDescripcionLinea As New ADODB.Recordset
Dim RBuscaGrupo As New ADODB.Recordset
Dim RBuscaEquipo As New ADODB.Recordset
Dim RBuscaEmpleado As New ADODB.Recordset

'VARIABLES PARA BUSQUEDA DE DATOS
Dim BGrupos As Boolean
Dim BParos As Boolean
Dim BGrupos2 As Boolean
Dim BEficiencia As Boolean
Dim BEficiencia2 As Boolean
Dim BGrupoParo As Boolean
Dim BCliente As Boolean
Dim BEquipo As Boolean

Dim RBuscaOrden As New ADODB.Recordset
Dim RCapturaParos As New ADODB.Recordset
Dim RBuscaCliente As New ADODB.Recordset

Dim BMenorHorasD As Boolean
Dim BMenorHorasN As Boolean

Dim VGrupoDia As String
Dim VGrupoNoche As String

'REPORTE EJECUTIVO DIA
Dim VOrdenDetalle As String
Dim VProduccionPC As Long
Dim VProduccionPNC As Long
Dim VProduccionDes As Long
Dim VFichaTecnicaOrden As String
Dim VPesoFichaTecnica As Single

Dim VToneladasPC As Single
Dim VToneladasPNC As Single
Dim VToneladasDes As Single
Dim VToneladasCalculoPC As Single
Dim VToneladasCalculoPNC As Single
Dim VToneladasCalculoDes As Single
Dim VUnidadesPC As Single
Dim VUnidadesPNC As Single
Dim VUnidadesDes As Single


Dim RBuscaDetalleProduccion As New ADODB.Recordset
Dim RBuscaFichaTecnica As New ADODB.Recordset

'REPORTE EJECUTIVO ACUMULADO
Dim VAOrdenDetalle As String
Dim VAProduccionPC As Long
Dim VAProduccionPNC As Long
Dim VAProduccionDes As Long
Dim VAFichaTecnicaOrden As String
Dim VAPesoFichaTecnica As Single

Dim VAToneladasPC As Single
Dim VAToneladasPNC As Single
Dim VAToneladasDes As Single
Dim VAToneladasCalculoPC As Single
Dim VAToneladasCalculoPNC As Single
Dim VAToneladasCalculoDes As Single
Dim VAUnidadesPC As Single
Dim VAUnidadesPNC As Single
Dim VAUnidadesDes As Single


Dim RABuscaDetalleProduccion As New ADODB.Recordset
Dim RABuscaFichaTecnica As New ADODB.Recordset
Dim RABuscaOrden As New ADODB.Recordset

Dim RInventario As New ADODB.Recordset
Dim RReporteEjecutivoInventario As New ADODB.Recordset
Dim RVentas As New ADODB.Recordset
Dim RReporteEjecutivoVentas As New ADODB.Recordset

Dim VTexto As String

Dim RBuscaEficiencia As New ADODB.Recordset

Dim RBuscaTipo As New ADODB.Recordset
Dim RBuscaMetaMensual As New ADODB.Recordset
Dim RBuscaDiasVenta As New ADODB.Recordset
Dim VFechaMeta As Date
Dim VMetaDolares As Currency
Dim VMetaCantidad As Currency
Dim VDiasVenta As Integer

' SE agrega cambios para reporte ejec. en Toneladas
Dim VMetaToneladas As Currency

Dim RClientes As New ADODB.Recordset
Dim RFichaTecnica As New ADODB.Recordset

Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11

    

  'MATERIAS PRIMAS
  If TabReportes.Tab = 0 Then
                                Paros
  ElseIf TabReportes.Tab = 1 Then
                                CapturaParos
  ElseIf TabReportes.Tab = 2 Then
                                'CrReportes.DiscardSavedData = True
                                GCriteriaReporte = ""
                                If (OptEfiEfi.Value = True Or OptEfiEfiPar.Value = True Or OptEfiLinResEmp.Value = True) Then
                                    Conexion.Execute ("Delete from ReporteEficiencia")
                                    EficienciaPorLinea
                                ElseIf (OptEfiGru.Value = True Or OptEfiGruRes.Value = True Or OptEfiGruResEmp.Value = True) Then
                                    Conexion.Execute ("Delete from ReporteEficienciaGrupos")
                                    EficienciaPorGrupo
                                ElseIf OptRepEje.Value = True Then
                                    Conexion.Execute ("Delete from ReporteEjecutivoDia")
                                    Conexion.Execute ("Delete from ReporteEjecutivoAcumulado")
                                    Conexion.Execute ("Delete from ReporteEjecutivoDia2")
                                    Conexion.Execute ("Delete from ReporteEjecutivoAcumulado2")
                                    Conexion.Execute ("Delete From ReporteEjecutivoVentas")
                                    Conexion.Execute ("Delete From ReporteEjecutivoInventario")
                                    Conexion.Execute ("Delete From ReporteEjecutivoVentasNuevas")
                                    Conexion.Execute ("Delete From ReporteEjecutivoProduccionPorDia")
                                    Conexion.Execute ("Delete From ReporteEjecutivoProduccionPorDiaGrafica1")
                                    Conexion.Execute ("Delete From ReporteEjecutivoProduccionPorDiaGrafica2")
                                    Conexion.Execute ("Delete From ReporteEjecutivoProduccionPorDiaGrafica3")
                                    Conexion.Execute ("Delete From ReporteEjecutivoProduccionPorDiaGrafica4")
                                    Conexion.Execute ("Delete From ReporteEjecutivoProduccionPorDiaGrafica5")
                                    ReporteEjecutivo
                                End If
  End If
  
'********************************************************************************************************************************************************************************************
                MousePointer = 0
                FrmReporte.Show
                
                
                If Err > 0 Then
                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                        Err.Clear
                End If
  
End Sub

Private Sub CmdSalida_Click()
    Unload Me
    
End Sub


Private Sub CmdSalir_Click()
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_DblClick()
        If BGrupoParo = True Then
                TxtPar.Text = DBGridBusqueda.Columns(0)
                TxtPar.SetFocus
        ElseIf BEficiencia = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(0)
                TxtLinea.SetFocus
        ElseIf BGrupos = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(2)
                TxtLinea.SetFocus
        ElseIf BEficiencia2 = True Then
                TxtCapPar.Text = DBGridBusqueda.Columns(0)
                TxtCapPar.SetFocus
        ElseIf BGrupos2 = True Then
                TxtCapPar.Text = DBGridBusqueda.Columns(2)
                TxtCapPar.SetFocus
        ElseIf BEquipo = True Then
                TxtCapPar.Text = DBGridBusqueda.Columns(0)
                TxtCapPar.SetFocus
        End If
                FrameBusqueda.Visible = False
        
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
        If BGrupoParo = True Then
                TxtPar.Text = DBGridBusqueda.Columns(0)
                TxtPar.SetFocus
        ElseIf BEficiencia = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(0)
                TxtLinea.SetFocus
        ElseIf BGrupos = True Then
                TxtLinea.Text = DBGridBusqueda.Columns(2)
                TxtLinea.SetFocus
        ElseIf BEficiencia2 = True Then
                TxtCapPar.Text = DBGridBusqueda.Columns(0)
                TxtCapPar.SetFocus
        ElseIf BGrupos2 = True Then
                TxtCapPar.Text = DBGridBusqueda.Columns(2)
                TxtCapPar.SetFocus
        ElseIf BEquipo = True Then
                TxtCapPar.Text = DBGridBusqueda.Columns(0)
                TxtCapPar.SetFocus
        End If
                FrameBusqueda.Visible = False
    End If

End Sub

Private Sub Form_Load()
            
            'SI TIENE ACTIVA LA OPCION PARA VER INVENTARIO, VENTAS Y REPORTE EJEUCTIVO
            If GInvVenRepEje = True Then
                    OptRepEje.Visible = True
                    OptEfiGruResEmp.Visible = True
                    OptEfiLinResEmp.Visible = True
            Else
                    OptRepEje.Visible = False
                    OptEfiGruResEmp.Visible = False
                    OptEfiLinResEmp.Visible = False
            End If
            
            DTPFecEfi.Value = Now
            DTPFecEfiAcu.Value = Now
            DTPFecEfiFin.Value = Now
            DTPFecEfiFinAcu.Value = Now
            DTPFecFinPar.Value = Now
            DTPFecIniPar.Value = Now
            DTPInventario.Value = Now
            DTPVentas.Value = Now
            
End Sub


Private Sub OptBusqueda_Click(Index As Integer)
        If Index = 0 Then
            LblBusqueda.Caption = "Descripcion"
        ElseIf Index = 1 Then
            LblBusqueda.Caption = "Codigo"
        End If
            TxtBusqueda.SetFocus
        
End Sub

Private Sub OptCapPar_Click(Index As Integer)
    'FECHAS
    If Index = 0 Then
        DTPFecIniPar.Visible = True
        DTPFecFinPar.Visible = True
        LblFecIni.Visible = True
        LblFecFin.Visible = True
        TxtCapPar.Visible = False
        LblCapPar.Caption = ""
    'LINEA
    ElseIf Index = 1 Then
        DTPFecIniPar.Visible = True
        DTPFecFinPar.Visible = True
        LblFecIni.Visible = True
        LblFecFin.Visible = True
        TxtCapPar.Visible = True
        LblCapPar.Caption = "Linea"
        TxtCapPar.SetFocus
    'DOCUMENTO
    ElseIf Index = 2 Then
        DTPFecIniPar.Visible = False
        DTPFecFinPar.Visible = False
        LblFecIni.Visible = False
        LblFecFin.Visible = False
        TxtCapPar.Visible = True
        LblCapPar.Caption = "Documento"
        TxtCapPar.SetFocus
    'GRUPO
    ElseIf Index = 3 Then
        DTPFecIniPar.Visible = True
        DTPFecFinPar.Visible = True
        LblFecIni.Visible = True
        LblFecFin.Visible = True
        TxtCapPar.Visible = True
        LblCapPar.Caption = "Grupo"
        TxtCapPar.SetFocus
    'EQUIPO
    ElseIf Index = 4 Then
        DTPFecIniPar.Visible = True
        DTPFecFinPar.Visible = True
        LblFecIni.Visible = True
        LblFecFin.Visible = True
        TxtCapPar.Visible = True
        LblCapPar.Caption = "Equipo"
        TxtCapPar.SetFocus
    
    
    End If

End Sub

Private Sub OptDetalle_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = False
        
End Sub

Private Sub OptEfi_Click(Index As Integer)
        If OptEfi.Item(0).Value = True Then
            Label2.Item(0).Caption = "Linea"
            DbGridGrupos.Visible = False
        Else
            Label2.Item(0).Caption = "Grupo"
            DbGridGrupos.Visible = True
            Set RGrupos = New ADODB.Recordset
                Call Abrir_Recordset(RGrupos, "Select Linea, Descrip, Grupo From Lineas")
                Set DbGridGrupos.DataSource = RGrupos
        End If
        
            TxtLinea.SetFocus
End Sub

Private Sub OptEfiEfi_Click()
        Label2.Item(3).Visible = False
        Label2.Item(4).Visible = False
        DTPFecEfiAcu.Visible = False
        DTPFecEfiFinAcu.Visible = False
        DTPVentas.Visible = False
        DTPInventario.Visible = False
        Label2.Item(5).Visible = False
        Label2.Item(6).Visible = False
        TxtEfiGru.Visible = False
        LblEfiGru.Visible = False
End Sub

Private Sub OptEfiEfiPar_Click()
        Label2.Item(3).Visible = False
        Label2.Item(4).Visible = False
        DTPFecEfiAcu.Visible = False
        DTPFecEfiFinAcu.Visible = False
        DTPVentas.Visible = False
        DTPInventario.Visible = False
        Label2.Item(5).Visible = False
        Label2.Item(6).Visible = False
        TxtEfiGru.Visible = False
        LblEfiGru.Visible = False
End Sub

Private Sub OptEfiGru_Click()
        Label2.Item(3).Visible = False
        Label2.Item(4).Visible = False
        DTPFecEfiAcu.Visible = False
        DTPFecEfiFinAcu.Visible = False
        DTPVentas.Visible = False
        DTPInventario.Visible = False
        Label2.Item(5).Visible = False
        Label2.Item(6).Visible = False
        TxtEfiGru.Visible = False
        LblEfiGru.Visible = False
End Sub

Private Sub OptEfiGruRes_Click()
        Label2.Item(3).Visible = False
        Label2.Item(4).Visible = False
        DTPFecEfiAcu.Visible = False
        DTPFecEfiFinAcu.Visible = False
        DTPVentas.Visible = False
        DTPInventario.Visible = False
        Label2.Item(5).Visible = False
        Label2.Item(6).Visible = False
        TxtEfiGru.Visible = False
        LblEfiGru.Visible = False
End Sub


Private Sub OptPar_Click(Index As Integer)
        If Index = 0 Then
            LblParos.Caption = "Codigo Paro"
        ElseIf Index = 1 Then
            LblParos.Caption = "Tipo De Paro"
        ElseIf Index = 2 Then
            LblParos.Caption = "Grupo"
            
        End If
            TxtPar.SetFocus
End Sub

Private Sub OptRepEje_Click()
        Label2.Item(3).Visible = True
        Label2.Item(4).Visible = True
        Label2.Item(5).Visible = True
        Label2.Item(6).Visible = True
        DTPFecEfiAcu.Visible = True
        DTPFecEfiFinAcu.Visible = True
        DTPVentas.Visible = True
        DTPInventario.Visible = True
        TxtEfiGru.Visible = True
        LblEfiGru.Visible = True
        DTPVentas.SetFocus
End Sub


Private Sub OptResumenCuadriculaMes_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = True
End Sub

Private Sub OptResumenCuadriculaMesPS_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = False

End Sub

Private Sub OptResumenEficienciaEquipo_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = False

End Sub

Private Sub OptResumenEficienciaLinea_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = False

End Sub

Private Sub OptResumenEmpleado_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = True
End Sub

Private Sub OptResumenEquipo_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = True
End Sub

Private Sub OptResumenLinea_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = True
End Sub

Private Sub OptResumenTurno_Click()
        OptFicPar.Item(0).Visible = True
        OptFicPar.Item(1).Visible = True

End Sub

Private Sub TabReportes_Click(PreviousTab As Integer)
        
If TabReportes.Tab = 0 Then
        OptPar.Item(0).Value = True
ElseIf TabReportes.Tab = 1 Then
        OptCapPar.Item(0).Value = True
        DTPFecIniPar.Value = Date
        DTPFecFinPar.Value = Date
ElseIf TabReportes.Tab = 2 Then
        OptEfi.Item(0).Value = True
        DTPFecEfi.Value = Date
        DTPFecEfiFin.Value = Date
        DTPFecEfiAcu.Value = Date
        DTPFecEfiFinAcu.Value = Date
        DTPVentas.Value = Date
        DTPInventario.Value = Date
End If

End Sub

Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'BUSCA LINEA
            If (BEficiencia = True Or BEficiencia2 = True Or BGrupos = True Or BGrupos2 = True) Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo from Lineas Where Descrip Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo from Lineas Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo from Lineas Where Linea Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo from Lineas Where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            End If
            
            'BUSCA EQUIPO
            If BEquipo = True Then
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
            
            
            'GRUPOS DE PARO
            If BGrupoParo = True Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoGrupo, Descripcion from ParosGrupos Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoGrupo, Descripcion from ParosGrupos Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoGrupo, Descripcion from ParosGrupos Where CodigoGrupo Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoGrupo, Descripcion from ParosGrupos Where UPPER(CodigoGrupo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            End If
            
            'CLIENTES
            If BCliente = True Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where CodigoCliente Like '%" & TxtBusqueda.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion from Clientes Where UPPER(CodigoCliente) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                            End If
                    End If
            End If
                            Set DBGridBusqueda.DataSource = RBusqueda
                            

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

Private Sub TxtCapPar_Change()
    'OPCION DE LINEA
    If OptCapPar.Item(1).Value = True Then
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtCapPar.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtCapPar.Text) & "'")
            End If
            
            If RBuscaLinea.RecordCount > 0 Then
                LblCapParLin.Caption = RBuscaLinea!Descrip
            Else
                LblCapParLin.Caption = ""
            End If
    'X EQUIPO
    ElseIf OptCapPar.Item(4).Value = True Then
        Set RBuscaEquipo = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where Codigo = '" & TxtCapPar.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaEquipo, "Select Descripcion From EmpleadosGrupos Where UPPER(Codigo) = '" & UCase(TxtCapPar.Text) & "'")
            End If
            If RBuscaEquipo.RecordCount > 0 Then
                LblCapParLin.Caption = RBuscaEquipo!Descripcion
            Else
                LblCapParLin.Caption = ""
            End If
    End If

End Sub

Private Sub TxtCapPar_DblClick()
            Set RBusqueda = New ADODB.Recordset
                    If OptCapPar.Item(1).Value = True Then
                        BEficiencia = False
                        BEficiencia2 = True
                        BGrupos = False
                        BGrupos2 = False
                        BGrupoParo = False
                        BCliente = False
                        BEquipo = False
                        BParos = False
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                    ElseIf OptCapPar.Item(3).Value = True Then
                        BGrupos = False
                        BGrupos2 = True
                        BEficiencia = False
                        BEficiencia2 = False
                        BGrupoParo = False
                        BCliente = False
                        BEquipo = False
                        BParos = False
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                    ElseIf OptCapPar.Item(4).Value = True Then
                        BGrupos = False
                        BGrupos2 = False
                        BEficiencia = False
                        BEficiencia2 = False
                        BGrupoParo = False
                        BCliente = False
                        BEquipo = True
                        BParos = False
                        Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
                    End If
                        FrameBusqueda.Visible = True
                        Set DBGridBusqueda.DataSource = RBusqueda
                        DBGridBusqueda.Columns(1).Width = "4000"
                        DBGridBusqueda.SetFocus

End Sub

Private Sub TxtCapPar_GotFocus()
        TxtCapPar.SelStart = 0
        TxtCapPar.SelLength = Len(TxtCapPar.Text)
End Sub

Private Sub TxtCapPar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
    
        If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                If OptCapPar.Item(1).Value = True Then
                    BEficiencia = False
                    BEficiencia2 = True
                    BGrupos = False
                    BGrupos2 = False
                    BGrupoParo = False
                    BCliente = False
                    BEquipo = False
                    BParos = False
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                ElseIf OptCapPar.Item(3).Value = True Then
                    BGrupos = False
                    BGrupos2 = True
                    BEficiencia = False
                    BEficiencia2 = False
                    BGrupoParo = False
                    BCliente = False
                    BEquipo = False
                    BParos = False
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                ElseIf OptCapPar.Item(4).Value = True Then
                    BGrupos = False
                    BGrupos2 = False
                    BEficiencia = False
                    BEficiencia2 = False
                    BGrupoParo = False
                    BCliente = False
                    BEquipo = True
                    BParos = False
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From EmpleadosGrupos")
                End If
                    FrameBusqueda.Visible = True
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    DBGridBusqueda.SetFocus
        End If
End Sub

Private Sub TxtLinea_Change()
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLinea.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
                LblEfiLin.Caption = RBuscaLinea!Descrip
            Else
                LblEfiLin.Caption = ""
            End If
End Sub

Private Sub TxtLinea_DblClick()
            Set RBusqueda = New ADODB.Recordset
            If OptEfi.Item(0).Value = True Then
                BEficiencia = True
                BEficiencia2 = False
                BGrupos = False
                BGrupos2 = False
                BGrupoParo = False
                BCliente = False
                BEquipo = False
            ElseIf OptEfi.Item(1).Value = True Then
                BGrupos = True
                BGrupos2 = False
                BEficiencia = False
                BEficiencia = False
                BGrupoParo = False
                BCliente = False
                BEquipo = False
            End If
                BParos = False
                FrameBusqueda.Visible = True
                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                DBGridBusqueda.SetFocus
End Sub

Private Sub TxtLinea_GotFocus()
        TxtLinea.SelStart = 0
        TxtLinea.SelLength = Len(TxtLinea.Text)
End Sub

Private Sub TxtLinea_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If

            If KeyAscii = 43 Then
                    Set RBusqueda = New ADODB.Recordset
                    If OptEfi.Item(0).Value = True Then
                        BEficiencia = True
                        BEficiencia2 = False
                        BGrupos = False
                        BGrupos2 = False
                        BGrupoParo = False
                        BCliente = False
                        BEquipo = False
                    ElseIf OptEfi.Item(1).Value = True Then
                        BGrupos = True
                        BGrupos2 = False
                        BEficiencia = False
                        BEficiencia2 = False
                        BGrupoParo = False
                        BCliente = False
                        BEquipo = False
                    End If
                        BParos = False
                        FrameBusqueda.Visible = True
                        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                        Set DBGridBusqueda.DataSource = RBusqueda
                        DBGridBusqueda.Columns(1).Width = "4000"
                        DBGridBusqueda.SetFocus
            End If

End Sub



Public Sub Paros()
On Error Resume Next
                       If OptPar.Item(0).Value = True Then
                            GCriteriaReporte = "{Paros.CodigoParo} Like '" & TxtPar.Text & "*'"
                       ElseIf OptPar.Item(1).Value = True Then
                            GCriteriaReporte = "{Paros.Tipo} = '" & TxtPar.Text & "'"
                       ElseIf OptPar.Item(2).Value = True Then
                            GCriteriaReporte = "{Paros.Grupo} = '" & TxtPar.Text & "'"
                       End If
                         
                       If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "FichaTecnicaParos.rpt"
                       Else
                            GNombreReporte = "FichaTecnicaParosO.rpt"
                       End If

End Sub

Public Sub CapturaParos()
On Error Resume Next
                        VDia = Day(DTPFecIniPar.Value)
                        VMes = Month(DTPFecIniPar.Value)
                        VAño = Year(DTPFecIniPar.Value)
                        VDia2 = Day(DTPFecFinPar.Value)
                        VMes2 = Month(DTPFecFinPar.Value)
                        VAño2 = Year(DTPFecFinPar.Value)
                        
                        
                        'FECHAS
                        If OptCapPar.Item(0).Value = True Then
                                           GTituloReporte = "Por Fechas Desde " & DTPFecIniPar.Value & " Hasta " & DTPFecFinPar.Value
                                           'GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                           GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} >= #" & Format(DTPFecIniPar.Value, "mm/dd/yyyy") & "# And {EncabezadoCapturaParos.Fecha} <= #" & Format(DTPFecFinPar.Value) & "#"
                        'FECHAS Y LINEA
                        ElseIf OptCapPar.Item(1).Value = True Then
                                           GTituloReporte = "Por Fechas Desde " & DTPFecIniPar.Value & " Hasta " & DTPFecFinPar.Value & " Y Linea " & TxtCapPar.Text
                                           GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoCapturaParos.Linea} = '" & TxtCapPar.Text & "'"
                        'DOCUMENTO
                        ElseIf OptCapPar.Item(2).Value = True Then
                                           GTituloReporte = "Por Documento " & TxtCapPar.Text
                                           GCriteriaReporte = "{EncabezadoCapturaParos.Documento} = " & TxtCapPar.Text
                        'FECHAS Y GRUPO
                        ElseIf OptCapPar.Item(3).Value = True Then
                                           GTituloReporte = "Por Fechas Desde " & DTPFecIniPar.Value & " Hasta " & DTPFecFinPar.Value & " Y Grupo " & TxtCapPar.Text
                                           GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoCapturaParos.Linea} = {Lineas.Linea} And {Lineas.Grupo} = '" & TxtCapPar.Text & "'"
                        'FECHAS Y EQUIPO
                        ElseIf OptCapPar.Item(4).Value = True Then
                                           GTituloReporte = "Por Fechas Desde " & DTPFecIniPar.Value & " Hasta " & DTPFecFinPar.Value & " Y Equipo " & TxtCapPar.Text
                                           GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {EncabezadoCapturaParos.Grupo} = '" & TxtCapPar.Text & "'"
                        End If
                         
                         
                         
    'ELIGE EL REPORTE ------------------------------------------------------------------------
                        'TODOS
                        If OptDetalle.Value = True Then
                            'TODOS
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParos.rpt"
                                Else
                                     GNombreReporte = "CapturaParosO.rpt"
                                End If
                        'LINEA
                        ElseIf OptResumenLinea.Value = True Then
                            'TODOS
                            If OptFicPar.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLinea.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLineaO.rpt"
                                End If
                            '10 PRIMEROS
                            Else
                                GTituloReporte = GTituloReporte & " 10 Mas Altos"
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLinea10Primeros.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLinea10PrimerosO.rpt"
                                End If
                            End If
                        'X EQUIPO
                        ElseIf OptResumenEquipo.Value = True Then
                            'TODOS
                            If OptFicPar.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenEquipo.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenEquipoO.rpt"
                                End If
                            '10 PRIMEROS
                            Else
                                GTituloReporte = GTituloReporte & " 10 Mas Altos"
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenEquipo10Primeros.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenEquipo10PrimerosO.rpt"
                                End If
                            End If
                        ElseIf OptResumenEmpleado.Value = True Then
                            'TODOS
                            If OptFicPar.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLineaEmpleado.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLineaEmpleadoO.rpt"
                                End If
                            '10 PRIMEROS
                            Else
                                GTituloReporte = GTituloReporte & " 10 Mas Altos"
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLineaEmpleado10Primeros.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLineaEmpleado10PrimerosO.rpt"
                                End If
                            End If
                        'RESUMEN CUADRICULA MES
                        ElseIf OptResumenCuadriculaMes.Value = True Then
                            'TODOS
                            If OptFicPar.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLineaCuadriculaPorMes.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLineaCuadriculaPorMesO.rpt"
                                End If
                            '10 PRIMEROS
                            Else
                                GTituloReporte = GTituloReporte & " 10 Mas Altos"
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLineaCuadricula10PrimerosPorMes.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLineaCuadricula10PrimerosPorMesO.rpt"
                                End If
                            End If
                        'RESUMEN CUADRICULA MES HORAS PROGRAMADAS Y PAROS S
                        ElseIf OptResumenCuadriculaMesPS.Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLineaCuadriculaPorMesPS.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLineaCuadriculaPorMesPSO.rpt"
                                End If
                        'EFICIENCIAS X LINEA
                        ElseIf OptResumenEficienciaLinea.Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosEficienciaLinea.rpt"
                                Else
                                     GNombreReporte = "CapturaParosEficienciaLineaO.rpt"
                                End If
                        'EFICIENCIAS X EQUIPO
                        ElseIf OptResumenEficienciaEquipo.Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosEficienciaEquipo.rpt"
                                Else
                                     GNombreReporte = "CapturaParosEficienciaEquipoO.rpt"
                                End If
                        'PRODUCCION X LINEA
                        ElseIf OptResumenProduccion.Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosProduccion.rpt"
                                Else
                                     GNombreReporte = "CapturaParosProduccionO.rpt"
                                End If
                        'LINEA Y TURNO
                        ElseIf OptResumenTurno.Value = True Then
                            'TODOS
                            If OptFicPar.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLineaTurno.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLineaTurnoO.rpt"
                                End If
                            '10 PRIMEROS
                            Else
                                GTituloReporte = GTituloReporte & " 10 Mas Altos"
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "CapturaParosResumenLineaTurno10Primeros.rpt"
                                Else
                                     GNombreReporte = "CapturaParosResumenLineaTurno10PrimerosO.rpt"
                                End If
                            End If
                        
                        End If
                        
                       ' FrmReporte.Show
  
  
End Sub

Public Sub EficienciaPorLinea()
On Error Resume Next
            
            'SELECCIONA LAS LINEAS DE ACUERDO A LA OPCION
            If OptEfi.Item(0).Value = True = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Linea = '" & TxtLinea.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                            End If
            ElseIf OptEfi.Item(1).Value = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Grupo = '" & TxtLinea.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Grupo) = '" & UCase(TxtLinea.Text) & "'")
                            End If
            End If
             
             If RSeleccionaLineas.RecordCount > 0 Then
             Else
                    MsgBox "Linea O Lineas No Existen ", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
             End If
  
  
  
  'CREA UN CICLO CON LAS LINEAS POSIBLES DE ACUERDO A LA OPCION ELEGIDA
  Do Until RSeleccionaLineas.EOF
                        
                    'ASIGNA LA LINEA QUE ES SELECCIONADA
                    VLinea = RSeleccionaLineas!Linea
                            
                    'FECHA DE INICIO DEL RANGO
                    VFechaInicial = DTPFecEfi.Value
                    'FECHA DEL FINAL DEL RANGO
                    VFechaFinal = DTPFecEfiFin.Value
                        
                
                Do Until VFechaInicial > VFechaFinal
                        
                        
'VERIFICA SI HAY DATOS EN LA PRESENTE FECHA Y SI NO HAY PASA A LA SIGUIENTE FECHA
'ESTO NOS SIRVE PARA CUANDO SAQUEMOS EL REPORTE DE EFICIENCIA NO TOME EN CUENTA LOS DIAS QUE
'NO SE TRABAJO PORQUE AFECTA LA EFICIENCIA DE LINEA Y PLANTA
                
                Set RCapturaParos = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Linea = '" & VLinea & "'")
                    Else
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                    End If
                    
                    If RCapturaParos.RecordCount > 0 Then
                        'NO HACE NADA SI HAY DATOS ESTA BIEN
                    
                                
                  
                  '*******  TIEMPO PROGRAMADO Y VELOCIDAD **************************************************
  
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE DIA
                         Set RTiempoProgramadoD = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoD.RecordCount > 0 Then
                                VTiempoProgramadoD = RTiempoProgramadoD!HorasProgramadas
                                VVelocidadTeoricaDia = RTiempoProgramadoD!VelocidadTeorica
                                VVelocidadRealDia = RTiempoProgramadoD!VelocidadReal
                                VGrupoDia = RTiempoProgramadoD!Grupo
                                    VParosND = RTiempoProgramadoD!ParoN / 60
                                    VParosSD = RTiempoProgramadoD!Paros / 60
                                    VProduccionD = RTiempoProgramadoD!ParoP / 60
                                        VPCD = RTiempoProgramadoD!ProductoConforme
                                        VPNCD = RTiempoProgramadoD!ProductoNoConforme
                                        VPDD = RTiempoProgramadoD!Desperdicio
                                            VEficienciaRealD = RTiempoProgramadoD!Eficiencia
                                Else
                                VTiempoProgramadoD = 0
                                VVelocidadTeoricaDia = 0
                                VVelocidadRealDia = 0
                                VGrupoDia = ""
                                    VParosND = 0
                                    VParosSD = 0
                                    VProduccionD = 0
                                        VPCD = 0
                                        VPNCD = 0
                                        VPDD = 0
                                            VEficienciaRealD = 0
                             End If
                                               
                        
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE NOCHE
                         Set RTiempoProgramadoN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoN.RecordCount > 0 Then
                                VTiempoProgramadoN = RTiempoProgramadoN!HorasProgramadas
                                VVelocidadTeoricaNoche = RTiempoProgramadoN!VelocidadTeorica
                                VVelocidadRealNoche = RTiempoProgramadoN!VelocidadReal
                                VGrupoNoche = RTiempoProgramadoN!Grupo
                                    VParosNN = RTiempoProgramadoN!ParoN / 60
                                    VParosSN = RTiempoProgramadoN!Paros / 60
                                    VProduccionN = RTiempoProgramadoN!ParoP / 60
                                        VPCN = RTiempoProgramadoN!ProductoConforme
                                        VPNCN = RTiempoProgramadoN!ProductoNoConforme
                                        VPDN = RTiempoProgramadoN!Desperdicio
                                            VEficienciaRealN = RTiempoProgramadoN!Eficiencia
                             Else
                                VTiempoProgramadoN = 0
                                VVelocidadTeoricaNoche = 0
                                VVelocidadRealNoche = 0
                                VGrupoNoche = ""
                                    VParosNN = 0
                                    VParosSN = 0
                                    VProduccionN = 0
                                        VPCN = 0
                                        VPNCN = 0
                                        VPDN = 0
                                            VEficienciaRealN = 0
                             End If
                                                
                                                                                                
                 '********  PAROS QUE NO AFECTAN 'N' **************************************************************
                        
                        'BUSCAR PAROS QUE NO AFECTAN DEL TURNO DE DIA
                 '       Set RBuscaParosNoAfectanD = New ADODB.Recordset
                 '           If GOrigenDeDatos = "AmaproAccess" Then
                 '               Call Abrir_Recordset(RBuscaParosNoAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '1' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N' And EP.Linea = '" & VLinea & "'")
                 '           Else
                 '               Call Abrir_Recordset(RBuscaParosNoAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '1' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'N' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                 '           End If
                 '
                 '           If RBuscaParosNoAfectanD.RecordCount > 0 Then
                 '               If IsNull(RBuscaParosNoAfectanD(0)) Then
                 '                   VParosND = 0
                 '               Else
                 '                   VParosND = RBuscaParosNoAfectanD(0) / 60
                 '               End If
                 '           Else
                 '               VParosND = 0
                 '           End If
                 '
                        'BUSCAR PAROS QUE NO AFECTAN DEL TURNO DE NOCHE
                 '       Set RBuscaParosNoAfectanN = New ADODB.Recordset
                 '           If GOrigenDeDatos = "AmaproAccess" Then
                 '               Call Abrir_Recordset(RBuscaParosNoAfectanN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '2' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N' And EP.Linea = '" & VLinea & "'")
                 '           Else
                 '               Call Abrir_Recordset(RBuscaParosNoAfectanN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '2' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'N' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                 '           End If
                 '           If RBuscaParosNoAfectanN.RecordCount > 0 Then
                 '               If IsNull(RBuscaParosNoAfectanN(0)) Then
                 '                   VParosNN = 0
                 '               Else
                 '                   VParosNN = RBuscaParosNoAfectanN(0) / 60
                 '               End If
                 '           Else
                 '                   VParosNN = 0
                 '           End If
                                            
                '********  PAROS QUE SI AFECTAN 'S' **************************************************************
                                            
                                            
                        'BUSCAR PAROS QUE AFECTAN DEL TURNO DE DIA
                        
                 '       Set RBuscaParosSiAfectanD = New ADODB.Recordset
                 '           If GOrigenDeDatos = "AmaproAccess" Then
                 '               Call Abrir_Recordset(RBuscaParosSiAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '1' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S' And EP.Linea = '" & VLinea & "'")
                 '           Else
                 '               Call Abrir_Recordset(RBuscaParosSiAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '1' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'S' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                 '           End If
                 '
                 '           If RBuscaParosSiAfectanD.RecordCount > 0 Then
                 '               If IsNull(RBuscaParosSiAfectanD(0)) Then
                 '                   VParosSD = 0
                 '               Else
                 '                   VParosSD = RBuscaParosSiAfectanD(0) / 60
                 '               End If
                 '           Else
                 '               VParosSD = 0
                 '           End If
                                            
                 '       'BUSCAR PAROS QUE AFECTAN DEL TURNO DE NOCHE
                 '       Set RBuscaParosSiAfectanN = New ADODB.Recordset
                 '           If GOrigenDeDatos = "AmaproAccess" Then
                 '               Call Abrir_Recordset(RBuscaParosSiAfectanN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '2' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S' And EP.Linea = '" & VLinea & "'")
                 '           Else
                 '               Call Abrir_Recordset(RBuscaParosSiAfectanN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '2' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'S' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                 '           End If
                 '
                 '           If RBuscaParosSiAfectanN.RecordCount > 0 Then
                 '               If IsNull(RBuscaParosSiAfectanN(0)) Then
                 '                   VParosSN = 0
                 '               Else
                 '                   VParosSN = RBuscaParosSiAfectanN(0) / 60
                 '               End If
                 '           Else
                 '                   VParosSN = 0
                 '           End If
                           
                '********  PRODUCCION **************************************************************
                                            
                                            
                        'BUSCAR PRODUCCION DIA
                 '       Set RBuscaProduccionD = New ADODB.Recordset
                 '           If GOrigenDeDatos = "AmaproAccess" Then
                 '              Call Abrir_Recordset(RBuscaProduccionD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '1' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'P' And EP.Linea = '" & VLinea & "'")
                 '           Else
                 '               Call Abrir_Recordset(RBuscaProduccionD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '1' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'P' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                 '           End If
                                            
                 '           If RBuscaProduccionD.RecordCount > 0 Then
                 '               If IsNull(RBuscaProduccionD(0)) Then
                 '                   VProduccionD = 0
                 '               Else
                 '                   VProduccionD = RBuscaProduccionD(0) / 60
                 '               End If
                 '           Else
                 '               VProduccionD = 0
                 '           End If
                 '
                        'BUSCAR PRODUCCION NOCHE
                  '      Set RBuscaProduccionN = New ADODB.Recordset
                  '          If GOrigenDeDatos = "AmaproAccess" Then
                  '              Call Abrir_Recordset(RBuscaProduccionN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '2' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'P' And EP.Linea = '" & VLinea & "'")
                  '          Else
                  '              Call Abrir_Recordset(RBuscaProduccionN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '2' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'P' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                  '          End If
                  '          If RBuscaProduccionN.RecordCount > 0 Then
                  '              If IsNull(RBuscaProduccionN(0)) Then
                  '                  VProduccionN = 0
                  '              Else
                  '                  VProduccionN = RBuscaProduccionN(0) / 60
                  '              End If
                  '          Else
                  '                  VProduccionN = 0
                  '          End If
                                            
    '***************************************************************************************************************
                        
                        
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE DIA
                        VTiempoRealProducidoD = VTiempoProgramadoD - VParosND
                                                    
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE NOCHE
                        If VTiempoProgramadoN = 0 Then
                            VTiempoRealProducidoN = 0
                        Else
                            VTiempoRealProducidoN = VTiempoProgramadoN - VParosNN
                        End If
                                                    
                        'HORAS PRODUCIDAS POR LOS 2 TURNOS
                            VHorasProducidasDN = Format(VTiempoRealProducidoD + VTiempoRealProducidoN, "#,###,##0.00")
                                                    
                        'TOTAL DE PAROS S "NO AFECTAN"
                            VParosDN = Format(VParosND + VParosNN, "#,###,##0.00")
                        
    'DIA _______________________________________________________________________________________________________
                        
             'PRODUCTO CONFORME
                        'BUSCA EL TOTAL DE ENVASES DE ACUERDO A LA FECHA DEL TURNO DE DIA
                   '     Set RProduccion = New ADODB.Recordset
                   '         If GOrigenDeDatos = "AmaproAccess" Then
                   '             Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                   '         Else
                   '             Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                   '         End If
                   '
                   '         If RProduccion.RecordCount > 0 Then
                   '             If IsNull(RProduccion(0)) Then
                   '                 VPCD = 0
                   '             Else
                   '                 VPCD = RProduccion(0)
                   '             End If
                   '         Else
                   '             VPCD = 0
                   '         End If
                   '
            'PRODUCTO NO CONFORME
                        'BUSCA EL TOTAL DE ENVASES DE ACUERDO A LA FECHA DEL TURNO DE DIA
                   '     Set RProduccion = New ADODB.Recordset
                   '         If GOrigenDeDatos = "AmaproAccess" Then
                   '             Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                   '         Else
                   '             Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                   '         End If
                   '
                   '         If RProduccion.RecordCount > 0 Then
                   '             If IsNull(RProduccion(0)) Then
                   '                 VPNCD = 0
                   '             Else
                   '                 VPNCD = RProduccion(0)
                   '             End If
                   '         Else
                   '                 VPNCD = 0
                   '         End If
                   '
            'DESPERDICIO
                   '     Set RProduccion = New ADODB.Recordset
                   '         If GOrigenDeDatos = "AmaproAccess" Then
                   '             Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                   '         Else
                   '             Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                   '         End If
                   '
                   '         If RProduccion.RecordCount > 0 Then
                   '             If IsNull(RProduccion(0)) Then
                   '                 VPDD = 0
                   '             Else
                   '                 VPDD = RProduccion(0)
                   '             End If
                   '         Else
                   '             VPDD = 0
                   '         End If
                            
                            
                            
                  '          Set RProduccion = New ADODB.Recordset
                  '          If GOrigenDeDatos = "AmaproAccess" Then
                  '              Call Abrir_Recordset(RProduccion, "Select ProductoConforme, ProductoNoConforme, Desperdicio From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                  '          Else
                  '              Call Abrir_Recordset(RProduccion, "Select ProductoConforme, ProductoNoConforme, Desperdicio From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                  '          End If
                  '
                  '         If RProduccion.RecordCount > 0 Then
                  '            If IsNull(RProduccion(0)) Then
                  '                  VPCD = 0
                  '            Else
                  '                  VPCD = RProduccion(0)
                  '            End If
                  '            If IsNull(RProduccion(1)) Then
                  '                  VPNCD = 0
                  '            Else
                  '                  VPNCD = RProduccion(1)
                  '            End If
                  '            If IsNull(RProduccion(2)) Then
                  '                  VPDD = 0
                  '            Else
                  '                  VPDD = RProduccion(2)
                  '            End If
                  '         Else
                  '                  VPCD = 0
                  '                  VPNCD = 0
                  '                  VPDD = 0
                  '         End If
    'NOCHE _______________________________________________________________________________________________________
                                                             
                                                             
                                                             
            'PRODUCTO CONFORME
                       'Set RProduccion = New ADODB.Recordset
                       '     If GOrigenDeDatos = "AmaproAccess" Then
                       '         Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                       '     Else
                       '         Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                       '     End If
                       '
                       '    If RProduccion.RecordCount > 0 Then
                       '       If IsNull(RProduccion(0)) Then
                       '             VPCN = 0
                       '       Else
                       '             VPCN = RProduccion(0)
                       '       End If
                       '    Else
                       '             VPCN = 0
                       '    End If
                                                             
            'PRODUCTO NO CONFORME
                       ' Set RProduccion = New ADODB.Recordset
                       '     If GOrigenDeDatos = "AmaproAccess" Then
                       '         Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                       '     Else
                       '         Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Fecha = To_date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                       '     End If
                       '
                       '     If RProduccion.RecordCount > 0 Then
                       '         If IsNull(RProduccion(0)) Then
                       '             VPNCN = 0
                       '         Else
                       '             VPNCN = RProduccion(0)
                       '         End If
                       '     Else
                       '             VPNCN = 0
                       '     End If
                       '
            'DESPERDICIO
                       ' Set RProduccion = New ADODB.Recordset
                       '     If GOrigenDeDatos = "AmaproAccess" Then
                       '         Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                       '     Else
                       '         Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                       '     End If
                       '
                       '     If RProduccion.RecordCount > 0 Then
                       '         If IsNull(RProduccion(0)) Then
                       '             VPDN = 0
                       '         Else
                       '             VPDN = RProduccion(0)
                       '         End If
                       '     Else
                       '             VPDN = 0
                       '     End If
                                                            
                                                            
                  '     Set RProduccion = New ADODB.Recordset
                  '          If GOrigenDeDatos = "AmaproAccess" Then
                  '             Call Abrir_Recordset(RProduccion, "Select ProductoConforme, ProductoNoConforme, Desperdicio From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                  '          Else
                  '              Call Abrir_Recordset(RProduccion, "Select ProductoConforme, ProductoNoConforme, Desperdicio From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                  '          End If
                  '
                  '         If RProduccion.RecordCount > 0 Then
                  '            If IsNull(RProduccion(0)) Then
                  '                  VPCN = 0
                  '            Else
                  '                  VPCN = RProduccion(0)
                  '            End If
                  '            If IsNull(RProduccion(1)) Then
                  '                  VPNCN = 0
                  '            Else
                  '                  VPNCN = RProduccion(1)
                  '            End If
                  '            If IsNull(RProduccion(2)) Then
                  '                  VPDN = 0
                  '            Else
                  '                  VPDN = RProduccion(2)
                  '            End If
                  '         Else
                  '                  VPCN = 0
                  '                  VPNCN = 0
                  '                  VPDN = 0
                  '         End If
                  '
'________________________________________________________________________________________________________________________
'________________________________________________________________________________________________________________________
                        
                        
                        
                        'EL TOTAL DE LA PRODUCCION ES LA SUMA DEL PRODUCTO CONFORME Y NO CONFORME NO INCLUYE EL DESPERDICIO
                        'TOTAL PRODUCCION
                        VTotalProduccion = VPCD + VPNCD + VPCN + VPNCN
                        'TOTAL PRODUCCION DE DIA
                        VTotalProduccionD = VPCD + VPNCD
                        'TOTAL PRODUCCION DE NOCHE
                        VTotalProduccionN = VPCN + VPNCN
                        
                        
                        'SELECCIONA LA VELOCIDAD TEORICA DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaNoche
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaDia
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = ((VVelocidadTeoricaDia + VVelocidadTeoricaNoche) / 2)
                        ElseIf (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = 0
                        End If
                        
                        'SELECCIONA LA VELOCIDAD REAL DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadRealDia <= 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = VVelocidadRealNoche
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = VVelocidadRealDia
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = ((VVelocidadRealDia + VVelocidadRealNoche) / 2)
                        ElseIf (VVelocidadRealDia <= 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = 0
                        End If
                        
                                                                                                                                                                        
                                                                            
  'CONVIERTE LAS VARIABLES DE HORAS A MINUTOS PARA CALCULOS ________________________________________________________
                        'DIA
                            VTiempoProgramadoD = VTiempoProgramadoD * 60
                            VParosSD = VParosSD * 60
                            VParosND = VParosND * 60
                        'NOCHE
                            VTiempoProgramadoN = VTiempoProgramadoN * 60
                            VParosSN = VParosSN * 60
                            VParosNN = VParosNN * 60
                                                                                       
                                        
'____________________________________________________________ FACTORES __________________________________________________
                                       
                'DIA
                'FACTOR 1______________________________________________________________________
                '            VFactor1D = VTiempoProgramadoD - VParosND
                '            VFactor1D = VFactor1D - VParosSD
                '            VFactor1D = VFactor1D * VVelocidadRealDia
                '            If VFactor1D = 0 Then
                '            Else
                '               VFactor1D = VTotalProduccionD / VFactor1D
                '            End If
                                        
                'FACTOR 2______________________________________________________________________
                '            If VPNCD > VTotalProduccionD Then
                '                VFactor2D = 0
                '            Else
                '                    VFactor2D = VTotalProduccionD - VPNCD
                '                    If VTotalProduccionD = 0 Then
                '                    Else
                '                       VFactor2D = VFactor2D / VTotalProduccionD
                '                    End If
                '            End If
                '
                ''FACTOR 3______________________________________________________________________
                '            If VPDD > VTotalProduccionD Then
                '                VFactor3D = 0
                '            Else
                '                    VFactor3D = VTotalProduccionD - VPDD
                '                    If VTotalProduccionD = 0 Then
                '                    Else
                '                      VFactor3D = VFactor3D / VTotalProduccionD
                '                    End If
                '            End If
                                        
                'FACTOR 4______________________________________________________________________
                '            VFactor4D = VTiempoProgramadoD - VParosND
                '            VFactor4D = VFactor4D - VParosSD
                '            If (VTiempoProgramadoD - VParosND) = 0 Then
                '            Else
                '                VFactor4D = (VFactor4D / (VTiempoProgramadoD - VParosND))
                '            End If
                '
                                        
                'FACTOR 5______________________________________________________________________
                '            If VVelocidadTeoricaDia = 0 Then
                '                VFactor5D = 0
                '            Else
                '                VFactor5D = VVelocidadRealDia / VVelocidadTeoricaDia
                '            End If
                                                
                
                'EFICIENCIA REAL DEL TURNO DE DIA
                '             VEficienciaRealD = VFactor1D * VFactor2D * VFactor3D * VFactor4D * VFactor5D * 100
                             
                             
                'Set RBuscaEficiencia = New ADODB.Recordset
                '            If GOrigenDeDatos = "AmaproAccess" Then
                '                Call Abrir_Recordset(RBuscaEficiencia, "Select Eficiencia From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                '            Else
                '                Call Abrir_Recordset(RBuscaEficiencia, "Select Eficiencia From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                '            End If
                '
                '           If RBuscaEficiencia.RecordCount > 0 Then
                '                If IsNull(RBuscaEficiencia(0)) Then
                '                    VEficienciaRealD = 0
                '                Else
                '                    VEficienciaRealD = RBuscaEficiencia(0)
                '                End If
                '           Else
                '                VEficienciaRealD = 0
                '           End If
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
                                        
                'NOCHE
                'SI EL TIEMPO PROGRAMADO ES CER0 NO HACE NADA
                If VTiempoProgramadoN > 0 Then
                                         
                            'EFICIENCIA DEL TURNO DE NOCHE
                                                            
                            'FACTOR 1___________________________________________________________
                   '                 VFactor1N = VTiempoProgramadoN - VParosNN
                   '                 VFactor1N = VFactor1N - VParosSN
                   '                 VFactor1N = VFactor1N * VVelocidadRealNoche
                   '                 If VFactor1N = 0 Then
                   '                 Else
                   '                    VFactor1N = VTotalProduccionN / VFactor1N
                   '                 End If
                                                            
                            'FACTOR 2__________________________________________________________
                   '                 If VPNCN > VTotalProduccionN Then
                   '                     VFactor2N = 0
                   '                 Else
                   '                     VFactor2N = VTotalProduccionN - VPNCN
                   '                     If VTotalProduccionN = 0 Then
                   '                     Else
                   '                         VFactor2N = (VFactor2N / VTotalProduccionN)
                   '                     End If
                   '                 End If
                   '
                   '         'FACTOR 3___________________________________________________________
                   '                 If VPDN > VTotalProduccionN Then
                   '                     VFactor3N = 0
                   '                 Else
                   '                     VFactor3N = VTotalProduccionN - VPDN
                   '                     If VTotalProduccionN = 0 Then
                   '                     Else
                   '                         VFactor3N = VFactor3N / VTotalProduccionN
                   '                     End If
                   '                 End If
                   '
                   '         'FACTOR 4___________________________________________________________
                  '                  VFactor4N = VTiempoProgramadoN - VParosNN
                   '                 VFactor4N = VFactor4N - VParosSN
                   '                 If (VTiempoProgramadoN - VParosNN) = 0 Then
                   '                 Else
                   '                     VFactor4N = (VFactor4N / (VTiempoProgramadoN - VParosNN))
                   '                 End If
                   '
                   '         'FACTOR 5___________________________________________________________
                   '                 If VVelocidadTeoricaNoche = 0 Then
                   '                     VFactor5N = 0
                   '                 Else
                   '                     VFactor5N = VVelocidadRealNoche / VVelocidadTeoricaNoche
                   '                 End If
                   
                '    Set RBuscaEficiencia = New ADODB.Recordset
                '            If GOrigenDeDatos = "AmaproAccess" Then
                '                Call Abrir_Recordset(RBuscaEficiencia, "Select Eficiencia From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                '            Else
                '                Call Abrir_Recordset(RBuscaEficiencia, "Select Eficiencia From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                '            End If
                '
                '           If RBuscaEficiencia.RecordCount > 0 Then
                '                If IsNull(RBuscaEficiencia(0)) Then
                '                    VEficienciaRealN = 0
                '                Else
                '                    VEficienciaRealN = RBuscaEficiencia(0)
                '                End If
                '           Else
                '                VEficienciaRealN = 0
                '           End If
                Else
                   '     VFactor1N = 0
                   '     VFactor2N = 0
                   '     VFactor3N = 0
                   '     VFactor4N = 0
                   '     VFactor5N = 0
                End If
                           
                           'EFICIENCIA REAL DEL TURNO DE NOCHE
                   '         VEficienciaRealN = VFactor1N * VFactor2N * VFactor3N * VFactor4N * VFactor5N * 100
                                                                                        
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
                        
                '*********** QUITAMOS ESTA PARTE YA QUE NO SE UTILIZA Y ASI ES MAS RAPIDO *****************
                
                '% DE LINEA___________________________________________________________________________________________
                                                        
                            'FACTOR 1_________________________________________________________________________________
                                    VFactor1DN = (VTiempoProgramadoD - (VParosND + VParosSD)) + (VTiempoProgramadoN - (VParosNN + VParosSN))
                                    VFactor1DN = VFactor1DN * VVelocidadRealLinea
                                    If VFactor1DN = 0 Then
                                    Else
                                        VFactor1DN = VTotalProduccion / VFactor1DN
                                    End If
                
                                    'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                                    If VFactor1DN = 0 Then
                                        VFactor1DN = 1
                                    End If
                
                            'FACTOR 2_________________________________________________________________________________
                                    If (VPNCD + VPNCN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor2DN = VTotalProduccion - (VPNCD + VPNCN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor2DN = VFactor2DN / VTotalProduccion
                                        End If
                                    End If
                
                            'FACTOR 3_________________________________________________________________________________
                                    If (VPDD + VPDN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor3DN = VTotalProduccion - (VPDD + VPDN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor3DN = VFactor3DN / VTotalProduccion
                                        End If
                                    End If
                
                            'FACTOR 4_________________________________________________________________________________
                                    VFactor4DN = (((VTiempoProgramadoD - VParosND) - VParosSD) + ((VTiempoProgramadoN - VParosNN) - VParosSN))
                                    If (VTiempoProgramadoD + VTiempoProgramadoN) = 0 Then
                                    ElseIf ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)) = 0 Then
                                    Else
                                        VFactor4DN = (VFactor4DN / ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)))
                                    End If
                
                            'FACTOR 5___________________________________________________________
                                    If VVelocidadTeoricaLinea = 0 Then
                                        VFactor5DN = 0
                                    Else
                                        VFactor5DN = VVelocidadRealLinea / VVelocidadTeoricaLinea
                                    End If
                
                                    'SI EL FACTOR 5 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 5
                                    If VFactor5DN = 0 Then
                                        VFactor5DN = 1
                                    End If
                                                                                
                            'EFICIENCIA DE LINEA______________________________________________________________________
                                   VPorcentajeLinea = VFactor1DN * VFactor2DN * VFactor3DN * VFactor4DN * VFactor5DN * 100
                                    
                                   
                '% DE RECHAZO__________________________________________________________________________________________
                '______________________________________________________________________________________________________
                                    VPorcentajeRechazo = VPNCD + VPNCN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeRechazo = VPorcentajeRechazo / VTotalProduccion
                                        VPorcentajeRechazo = VPorcentajeRechazo * 100
                                    End If
                        
                        
                '% DE DESPERDICIO______________________________________________________________________________________
                '______________________________________________________________________________________________________
                
                                    VPorcentajeDesperdicio = VPDD + VPDN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio / VTotalProduccion
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio * 100
                                    End If
                                    
'_______________________________________________________________________________________________________________________
'************************************************* HORAS A MINUTOS *****************************************************
'_______________________________________________________________________________________________________________________

                                                
                'CONVIERTE LAS VARIABLES A HORAS PARA IMPRIMIR LOS DATOS
                                        
                            'DIA______________________________________________________________________
                                    If VTiempoProgramadoD = 0 Then
                                        VTiempoProgramadoD = 0
                                    Else
                                            VTiempoProgramadoD = VTiempoProgramadoD / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE DIA
                                    If VParosSD = 0 Then
                                        VParosSD = 0
                                    Else
                                        VParosSD = VParosSD / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE DIA
                                    If VParosND = 0 Then
                                        VParosND = 0
                                    Else
                                        VParosND = VParosND / 60
                                    End If
                                    
                            'NOCHE______________________________________________________________________
                                    If VTiempoProgramadoN = 0 Then
                                        VTiempoProgramadoN = 0
                                    Else
                                            VTiempoProgramadoN = VTiempoProgramadoN / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE NOCHE
                                    If VParosSN = 0 Then
                                        VParosSN = 0
                                    Else
                                        VParosSN = VParosSN / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE NOCHE
                                    If VParosNN = 0 Then
                                        VParosNN = 0
                                    Else
                                        VParosNN = VParosNN / 60
                                    End If
                                    
                                        
                                    If Err > 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                        Err.Clear
                                    End If
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________

                        
                        'BUSCA LA DESCRIPCION DE LA LINEA
                        Set RBuscaDescripcionLinea = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip From Lineas Where Linea = '" & VLinea & "'")
                            Else
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(VLinea) & "'")
                            End If
                           
                         
                         If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Insert Into ReporteEficiencia (Linea, Fecha, TiempoProgramado, TiempoRealProducidoD, TiempoRealProducidoN, ParosAfectanD, ParosAfectanN, TiempoRealProducido, ParosNoAfectanD, ParosNoAfectanN, ProductoConformeD, ProductoNoConformeD, DesperdicioD, ProductoConformeN, ProductoNoConformeN, DesperdicioN, TotalProduccion, EficienciaD, EficienciaN, PorcentajeLinea, PorcentajeRechazo, PorcentajeDesperdicio, Factor1, Factor5, VelocidadTeoricaDia, VelocidadRealDia, VelocidadTeoricaNoche, VelocidadRealNoche) " _
                                    & "VALUES('" & RBuscaDescripcionLinea(0) & "', #" & Format(VFechaInicial, "mm/dd/yyyy") & "#, " & (VTiempoProgramadoD + VTiempoProgramadoN) & ", " & VProduccionD & ", " & VProduccionN & ", " & VParosSD & ", " & VParosSN & ", " & (VProduccionD + VProduccionN) & ", " & VParosND & ", " & VParosNN & ", " & VPCD & ", " & VPNCD & ", " & VPDD & ", " & VPCN & ", " & VPNCN & ", " & VPDN & ", " & VTotalProduccion & ", " & VEficienciaRealD & ", " & VEficienciaRealN & ", " & VPorcentajeLinea & ", " & VPorcentajeRechazo & ", " & VPorcentajeDesperdicio & ", " & VFactor1DN & ", " & VFactor5DN & ", " & VVelocidadTeoricaDia & ", " & VVelocidadRealDia & ", " & VVelocidadTeoricaNoche & ", " & VVelocidadRealNoche & ")"
                         Else
                                Conexion.Execute "Insert Into ReporteEficiencia (Linea, Fecha, TiempoProgramado, TiempoRealProducidoD, TiempoRealProducidoN, ParosAfectanD, ParosAfectanN, TiempoRealProducido, ParosNoAfectanD, ParosNoAfectanN, ProductoConformeD, ProductoNoConformeD, DesperdicioD, ProductoConformeN, ProductoNoConformeN, DesperdicioN, TotalProduccion, EficienciaD, EficienciaN, PorcentajeLinea, PorcentajeRechazo, PorcentajeDesperdicio, Factor1, Factor5, VelocidadTeoricaDia, VelocidadRealDia, VelocidadTeoricaNoche, VelocidadRealNoche) " _
                                    & "VALUES('" & RBuscaDescripcionLinea(0) & "', To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & ", " & (VTiempoProgramadoD + VTiempoProgramadoN) & ", " & VProduccionD & ", " & VProduccionN & ", " & VParosSD & ", " & VParosSN & ", " & (VProduccionD + VProduccionN) & ", " & VParosND & ", " & VParosNN & ", " & VPCD & ", " & VPNCD & ", " & VPDD & ", " & VPCN & ", " & VPNCN & ", " & VPDN & ", " & VTotalProduccion & ", " & VEficienciaRealD & ", " & VEficienciaRealN & ", " & VPorcentajeLinea & ", " & VPorcentajeRechazo & ", " & VPorcentajeDesperdicio & ", " & VFactor1DN & ", " & VFactor5DN & ", " & VVelocidadTeoricaDia & ", " & VVelocidadRealDia & ", " & VVelocidadTeoricaNoche & ", " & VVelocidadRealNoche & ")"
                         End If
                                    
                                    If Err > 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                            Err.Clear
                                    End If
                                    
                            VFechaInicial = VFechaInicial + 1
                                    
                    Else ' IF DE CONDICION DE LA QUE NO ENCUENTRA DATOS EN RANGO DE FECHAS
                            
                            'A LA FECHA ACTUAL LE SUMA 1 PARA INCREMENTAR LA FECHA
                            'LA EFICIENCIA SE SACA POR DIA
                            VFechaInicial = VFechaInicial + 1
                    End If
                
                        
                Loop
                
                
                'SIGUE AL SIGUIENTE REGISTRO DE LINEA
                RSeleccionaLineas.MoveNext
  Loop
                
                    GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "                        Reporte del " & DTPFecEfi.Value & " Al " & DTPFecEfiFin.Value
             
                'TIPO DE REPORTE
                'EFICIENCIA Y PAROS
                If OptEfiEfiPar.Value = True Then
                
                        
                        'REPORTE DE EFICIENCIA CON PAROS
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "EficienciaResumen.rpt"
                                Else
                                     GNombreReporte = "EficienciaResumenO.rpt"
                                End If
                'SOLO EFICIENCIA
                ElseIf OptEfiEfi.Value = True Then
                        'REPORTE DE EFICIENCIA
                        'SELECCIONA TODA LAS CAPTURAS DE PAROS ENTRE  FECHAS Y LINEA
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "Eficiencia.rpt"
                                Else
                                     GNombreReporte = "EficienciaO.rpt"
                                End If
                ElseIf OptEfiLinResEmp.Value = True Then
                                If OptEfi.Item(0).Value = True Then
                                    GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} >= #" & Format(DTPFecEfi.Value, "mm/dd/yyyy") & "# And {EncabezadoCapturaParos.Fecha} <= #" & Format(DTPFecEfiFin.Value) & "# And {EncabezadoCapturaParos.Linea} = '" & TxtLinea.Text & "'"
                                ElseIf OptEfi.Item(1).Value = True Then
                                    GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} >= #" & Format(DTPFecEfi.Value, "mm/dd/yyyy") & "# And {EncabezadoCapturaParos.Fecha} <= #" & Format(DTPFecEfiFin.Value) & "# And {EncabezadoCapturaParos.Linea} = {Lineas.Linea} And {Lineas.Grupo} = '" & TxtLinea.Text & "'"
                                End If
                        
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "EficienciaEnLineasEmpleados.rpt"
                                Else
                                     GNombreReporte = "EficienciaEnLineasEmpleadosO.rpt"
                                End If
                
                End If

End Sub



Private Sub TxtPar_Change()
    If OptPar.Item(2).Value = True Then
        'BUSCA GRUPO
        Set RBuscaGrupo = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From ParosGrupos Where CodigoGrupo = '" & TxtPar.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaGrupo, "Select Descripcion From ParosGrupos Where UPPER(CodigoGrupo) = '" & UCase(TxtPar.Text) & "'")
            End If
            If RBuscaGrupo.RecordCount > 0 Then
                LblGruPar.Caption = RBuscaGrupo!Descripcion
            Else
                LblGruPar.Caption = ""
            End If
    End If
End Sub

Private Sub TxtPar_DblClick()
    'GRUPO DE PARO
    If OptPar.Item(2).Value = True Then
        BEficiencia = False
        BEficiencia2 = False
        BGrupos = False
        BGrupos2 = False
        BGrupoParo = True
        BParos = False
        BCliente = False
        BEquipo = False
        FrameBusqueda.Visible = True
        Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select CodigoGrupo, Descripcion From ParosGrupos")
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "4000"
        DBGridBusqueda.SetFocus
    End If

End Sub

Private Sub TxtPar_GotFocus()
        TxtPar.SelStart = 0
        TxtPar.SelLength = Len(TxtPar.Text)
End Sub

Private Sub TxtPar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
            'GRUPO DE PARO
            If OptPar.Item(2).Value = True Then
                BEficiencia = False
                BEficiencia2 = False
                BGrupos = False
                BGrupos2 = False
                BGrupoParo = True
                BParos = False
                BCliente = False
                BEquipo = False
                FrameBusqueda.Visible = True
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select CodigoGrupo, Descripcion From ParosGrupos")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                DBGridBusqueda.SetFocus
            End If
        End If
End Sub


Public Sub EficienciaPorGrupo()
On Error Resume Next
            Set RSeleccionaLineas = New ADODB.Recordset
            'SELECCIONA LAS LINEAS DE ACUERDO A LA OPCION
            If OptEfi.Item(0).Value = True = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Linea = '" & TxtLinea.Text & "'")
                    Else
                        Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                    End If
            ElseIf OptEfi.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Grupo = '" & TxtLinea.Text & "'")
                    Else
                        Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Grupo) = '" & UCase(TxtLinea.Text) & "'")
                    End If
            End If
             
             If RSeleccionaLineas.RecordCount > 0 Then
             Else
                    MsgBox "Linea O Lineas No Existen ", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
             End If
  
  
  
  'CREA UN CICLO CON LAS LINEAS POSIBLES DE ACUERDO A LA OPCION ELEGIDA
  Do Until RSeleccionaLineas.EOF
                        
                    'ASIGNA LA LINEA QUE ES SELECCIONADA
                    VLinea = RSeleccionaLineas!Linea
                                                
                    'FECHA DE INICIO DEL RANGO
                    VFechaInicial = DTPFecEfi.Value
                    'FECHA DEL FINAL DEL RANGO
                    VFechaFinal = DTPFecEfiFin.Value
                        
                
                Do Until VFechaInicial > VFechaFinal
                        
                        
'VERIFICA SI HAY DATOS EN LA PRESENTE FECHA Y SI NO HAY PASA A LA SIGUIENTE FECHA
'ESTO NOS SIRVE PARA CUANDO SAQUEMOS EL REPORTE DE EFICIENCIA NO TOME EN CUENTA LOS DIAS QUE
'NO SE TRABAJO PORQUE AFECTA LA EFICIENCIA DE LINEA Y PLANTA

                Set RCapturaParos = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Linea = '" & VLinea & "'")
                    Else
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                    End If
                    If RCapturaParos.RecordCount > 0 Then
                        'NO HACE NADA SI HAY DATOS ESTA BIEN
                    
                                
                  
                  '*******  TIEMPO PROGRAMADO Y VELOCIDAD **************************************************
  
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE DIA
                         Set RTiempoProgramadoD = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoD.RecordCount > 0 Then
                                VTiempoProgramadoD = RTiempoProgramadoD!HorasProgramadas
                                VVelocidadTeoricaDia = RTiempoProgramadoD!VelocidadTeorica
                                VVelocidadRealDia = RTiempoProgramadoD!VelocidadReal
                                VGrupoDia = RTiempoProgramadoD!Grupo
                                    VParosND = RTiempoProgramadoD!ParoN / 60
                                    VParosSD = RTiempoProgramadoD!Paros / 60
                                    VProduccionD = RTiempoProgramadoD!ParoP / 60
                                        VPCD = RTiempoProgramadoD!ProductoConforme
                                        VPNCD = RTiempoProgramadoD!ProductoNoConforme
                                        VPDD = RTiempoProgramadoD!Desperdicio
                                            VEficienciaRealD = RTiempoProgramadoD!Eficiencia
                                Else
                                VTiempoProgramadoD = 0
                                VVelocidadTeoricaDia = 0
                                VVelocidadRealDia = 0
                                VGrupoDia = ""
                                    VParosND = 0
                                    VParosSD = 0
                                    VProduccionD = 0
                                        VPCD = 0
                                        VPNCD = 0
                                        VPDD = 0
                                            VEficienciaRealD = 0
                             End If
                                               
                        
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE NOCHE
                         Set RTiempoProgramadoN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoN.RecordCount > 0 Then
                                VTiempoProgramadoN = RTiempoProgramadoN!HorasProgramadas
                                VVelocidadTeoricaNoche = RTiempoProgramadoN!VelocidadTeorica
                                VVelocidadRealNoche = RTiempoProgramadoN!VelocidadReal
                                VGrupoNoche = RTiempoProgramadoN!Grupo
                                    VParosNN = RTiempoProgramadoN!ParoN / 60
                                    VParosSN = RTiempoProgramadoN!Paros / 60
                                    VProduccionN = RTiempoProgramadoN!ParoP / 60
                                        VPCN = RTiempoProgramadoN!ProductoConforme
                                        VPNCN = RTiempoProgramadoN!ProductoNoConforme
                                        VPDN = RTiempoProgramadoN!Desperdicio
                                            VEficienciaRealN = RTiempoProgramadoN!Eficiencia
                             Else
                                VTiempoProgramadoN = 0
                                VVelocidadTeoricaNoche = 0
                                VVelocidadRealNoche = 0
                                VGrupoNoche = ""
                                    VParosNN = 0
                                    VParosSN = 0
                                    VProduccionN = 0
                                        VPCN = 0
                                        VPNCN = 0
                                        VPDN = 0
                                            VEficienciaRealN = 0
                             End If
                        
                                               
                                                                                                
                 '********  PAROS QUE NO AFECTAN 'N' **************************************************************
                        
                        'BUSCAR PAROS QUE NO AFECTAN DEL TURNO DE DIA
                        'Set RBuscaParosNoAfectanD = New ADODB.Recordset
                        '    If GOrigenDeDatos = "AmaproAccess" Then
                        '        Call Abrir_Recordset(RBuscaParosNoAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '1' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N' And EP.Linea = '" & VLinea & "'")
                        '    Else
                        '        Call Abrir_Recordset(RBuscaParosNoAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '1' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'N' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                        '    End If
                        '
                        '    If RBuscaParosNoAfectanD.RecordCount > 0 Then
                        '        If IsNull(RBuscaParosNoAfectanD(0)) Then
                        '            VParosND = 0
                        '        Else
                        '            VParosND = RBuscaParosNoAfectanD(0) / 60
                        '        End If
                        '    Else
                        '        VParosND = 0
                        '    End If
                                            
                        'BUSCAR PAROS QUE NO AFECTAN DEL TURNO DE NOCHE
                        'Set RBuscaParosNoAfectanN = New ADODB.Recordset
                        '    If GOrigenDeDatos = "AmaproAccess" Then
                        '        Call Abrir_Recordset(RBuscaParosNoAfectanN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '2' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'N' And EP.Linea = '" & VLinea & "'")
                        '    Else
                        '        Call Abrir_Recordset(RBuscaParosNoAfectanN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '2' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'N' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                        '    End If
                        '
                        '    If RBuscaParosNoAfectanN.RecordCount > 0 Then
                        '        If IsNull(RBuscaParosNoAfectanN(0)) Then
                        '            VParosNN = 0
                        '        Else
                        '            VParosNN = RBuscaParosNoAfectanN(0) / 60
                        '        End If
                        '    Else
                        '            VParosNN = 0
                        '    End If
                                            
                '********  PAROS QUE SI AFECTAN 'S' **************************************************************
                                            
                                            
                        'BUSCAR PAROS QUE AFECTAN DEL TURNO DE DIA
                        'Set RBuscaParosSiAfectanD = New ADODB.Recordset
                        '    If GOrigenDeDatos = "AmaproAccess" Then
                        '        Call Abrir_Recordset(RBuscaParosSiAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '1' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S' And EP.Linea = '" & VLinea & "'")
                        '    Else
                        '        Call Abrir_Recordset(RBuscaParosSiAfectanD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '1' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'S' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                        '    End If
                        '
                        '    If RBuscaParosSiAfectanD.RecordCount > 0 Then
                        '        If IsNull(RBuscaParosSiAfectanD(0)) Then
                        '            VParosSD = 0
                        '        Else
                        '            VParosSD = RBuscaParosSiAfectanD(0) / 60
                        '        End If
                        '    Else
                        '        VParosSD = 0
                        '    End If
                                            
                        'BUSCAR PAROS QUE AFECTAN DEL TURNO DE NOCHE
                        'Set RBuscaParosSiAfectanN = New ADODB.Recordset
                        '    If GOrigenDeDatos = "AmaproAccess" Then
                        '        Call Abrir_Recordset(RBuscaParosSiAfectanN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '2' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'S' And EP.Linea = '" & VLinea & "'")
                        '    Else
                        '        Call Abrir_Recordset(RBuscaParosSiAfectanN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '2' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'S' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                        '    End If
                        '
                        '    If RBuscaParosSiAfectanN.RecordCount > 0 Then
                        '        If IsNull(RBuscaParosSiAfectanN(0)) Then
                        '            VParosSN = 0
                        '        Else
                        '            VParosSN = RBuscaParosSiAfectanN(0) / 60
                        '        End If
                        '    Else
                        '            VParosSN = 0
                        '    End If
                           
                '********  PRODUCCION **************************************************************
                                            
                                            
                        'BUSCAR PRODUCCION DIA
                        'Set RBuscaProduccionD = New ADODB.Recordset
                        '    If GOrigenDeDatos = "AmaproAccess" Then
                        '        Call Abrir_Recordset(RBuscaProduccionD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '1' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'P' And EP.Linea = '" & VLinea & "'")
                        '    Else
                        '        Call Abrir_Recordset(RBuscaProduccionD, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '1' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'P' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                        '    End If
                        '
                        '    If RBuscaProduccionD.RecordCount > 0 Then
                        '        If IsNull(RBuscaProduccionD(0)) Then
                        '            VProduccionD = 0
                        '        Else
                        '            VProduccionD = RBuscaProduccionD(0) / 60
                        '        End If
                        '    Else
                        '        VProduccionD = 0
                        '    End If
                                            
                        'BUSCAR PRODUCCION NOCHE
                        'Set RBuscaProduccionN = New ADODB.Recordset
                        '    If GOrigenDeDatos = "AmaproAccess" Then
                        '        Call Abrir_Recordset(RBuscaProduccionN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Turno = '2' And EP.Documento = DP.Documento And DP.Paro = P.CodigoParo And P.Tipo = 'P' And EP.Linea = '" & VLinea & "'")
                        '    Else
                        '        Call Abrir_Recordset(RBuscaProduccionN, "Select Sum(DP.Minutos) From DetalleCapturaParos DP, EncabezadoCapturaParos EP, Paros P Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Turno) = '2' And EP.Documento = DP.Documento And UPPER(DP.Paro) = UPPER(P.CodigoParo) And UPPER(P.Tipo) = 'P' And UPPER(EP.Linea) = '" & UCase(VLinea) & "'")
                        '    End If
                        '
                        '    If RBuscaProduccionN.RecordCount > 0 Then
                        '        If IsNull(RBuscaProduccionN(0)) Then
                        '            VProduccionN = 0
                        '        Else
                        '            VProduccionN = RBuscaProduccionN(0) / 60
                        '        End If
                        '    Else
                        '            VProduccionN = 0
                        '    End If
                                            
    '***************************************************************************************************************
                        
                        
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE DIA
                        VTiempoRealProducidoD = VTiempoProgramadoD - VParosND
                                                    
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE NOCHE
                        If VTiempoProgramadoN = 0 Then
                            VTiempoRealProducidoN = 0
                        Else
                            VTiempoRealProducidoN = VTiempoProgramadoN - VParosNN
                        End If
                                                    
                        'HORAS PRODUCIDAS POR LOS 2 TURNOS
                            VHorasProducidasDN = Format(VTiempoRealProducidoD + VTiempoRealProducidoN, "#,###,##0.00")
                                                    
                        'TOTAL DE PAROS S "NO AFECTAN"
                            VParosDN = Format(VParosND + VParosNN, "#,###,##0.00")
                        
    'DIA _______________________________________________________________________________________________________
                        
             'PRODUCTO CONFORME
                        'BUSCA EL TOTAL DE ENVASES DE ACUERDO A LA FECHA DEL TURNO DE DIA
             '           Set RProduccion = New ADODB.Recordset
             '               If GOrigenDeDatos = "AmaproAccess" Then
             '                   Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
             '               Else
             '                   Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
             '               End If
             '
             '               If RProduccion.RecordCount > 0 Then
             '                   If IsNull(RProduccion(0)) Then
             '                       VPCD = 0
             '                   Else
             '                       VPCD = RProduccion(0)
             '                   End If
             '               Else
             '                   VPCD = 0
             '               End If
             '
            'PRODUCTO NO CONFORME
             '           'BUSCA EL TOTAL DE ENVASES DE ACUERDO A LA FECHA DEL TURNO DE DIA
             '           Set RProduccion = New ADODB.Recordset
             '               If GOrigenDeDatos = "AmaproAccess" Then
             '                   Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
             '               Else
             '                   Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
             '               End If
             '
             '               If RProduccion.RecordCount > 0 Then
             '                   If IsNull(RProduccion(0)) Then
             '                       VPNCD = 0
             '                   Else
             '                       VPNCD = RProduccion(0)
             '                   End If
             '               Else
             '                       VPNCD = 0
             '               End If
             '
            ''DESPERDICIO
             '           Set RProduccion = New ADODB.Recordset
             '               If GOrigenDeDatos = "AmaproAccess" Then
             '                   Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
             '               Else
             '                   Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
             '               End If
             '
             '               If RProduccion.RecordCount > 0 Then
             '                   If IsNull(RProduccion(0)) Then
             '                       VPDD = 0
             '                   Else
             '                       VPDD = RProduccion(0)
             '                   End If
             '               Else
             '                   VPDD = 0
             '               End If
                            
                   '     Set RProduccion = New ADODB.Recordset
                   '         If GOrigenDeDatos = "AmaproAccess" Then
                   '             Call Abrir_Recordset(RProduccion, "Select ProductoConforme, ProductoNoConforme, Desperdicio From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                   '         Else
                   '             Call Abrir_Recordset(RProduccion, "Select ProductoConforme, ProductoNoConforme, Desperdicio From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                   '         End If
                                                            
                   '        If RProduccion.RecordCount > 0 Then
                   '           If IsNull(RProduccion(0)) Then
                   '                 VPCD = 0
                   '           Else
                   '                 VPCD = RProduccion(0)
                   '           End If
                   '           If IsNull(RProduccion(1)) Then
                   '                 VPNCD = 0
                   '           Else
                   '                 VPNCD = RProduccion(1)
                   '           End If
                   '           If IsNull(RProduccion(2)) Then
                   '                 VPDD = 0
                   '           Else
                   '                 VPDD = RProduccion(2)
                   '           End If
                   '        Else
                   '                 VPCD = 0
                   '                 VPNCD = 0
                   '                 VPDD = 0
                   '        End If
                            
                            
    'NOCHE _______________________________________________________________________________________________________
                                                             
                                                             
                                                             
            'PRODUCTO CONFORME
            '           Set RProduccion = New ADODB.Recordset
            '                If GOrigenDeDatos = "AmaproAccess" Then
            '                    Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
            '                Else
            '                    Call Abrir_Recordset(RProduccion, "Select ProductoConforme From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
            '                End If
            '
            '               If RProduccion.RecordCount > 0 Then
            '                  If IsNull(RProduccion(0)) Then
            '                        VPCN = 0
            '                  Else
            '                        VPCN = RProduccion(0)
            '                  End If
            '               Else
            '                        VPCN = 0
            '               End If
            '
            ''PRODUCTO NO CONFORME
            '            Set RProduccion = New ADODB.Recordset
            '                If GOrigenDeDatos = "AmaproAccess" Then
            '                    Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
            '                Else
            '                    Call Abrir_Recordset(RProduccion, "Select ProductoNoConforme From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
            '                End If
            '
            '                If RProduccion.RecordCount > 0 Then
            '                    If IsNull(RProduccion(0)) Then
            '                        VPNCN = 0
            '                    Else
            '                        VPNCN = RProduccion(0)
            '                    End If
            '                Else
            '                        VPNCN = 0
            '                End If
            '
            ''DESPERDICIO
            '            Set RProduccion = New ADODB.Recordset
            '                If GOrigenDeDatos = "AmaproAccess" Then
            '                    Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
            '                Else
            '                    Call Abrir_Recordset(RProduccion, "Select Desperdicio From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
            '                End If
            '
            '                If RProduccion.RecordCount > 0 Then
            '                    If IsNull(RProduccion(0)) Then
            '                        VPDN = 0
            '                    Else
            '                        VPDN = RProduccion(0)
            '                    End If
            '                Else
            '                        VPDN = 0
            '                End If
            '
                 '           Set RProduccion = New ADODB.Recordset
                 '           If GOrigenDeDatos = "AmaproAccess" Then
                 '               Call Abrir_Recordset(RProduccion, "Select ProductoConforme, ProductoNoConforme, Desperdicio From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                 '           Else
                 '               Call Abrir_Recordset(RProduccion, "Select ProductoConforme, ProductoNoConforme, Desperdicio From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                 '           End If
                 '
                 '          If RProduccion.RecordCount > 0 Then
                 '             If IsNull(RProduccion(0)) Then
                 '                   VPCN = 0
                 '             Else
                 '                   VPCN = RProduccion(0)
                 '             End If
                 '             If IsNull(RProduccion(1)) Then
                 '                   VPNCN = 0
                 '             Else
                 '                   VPNCN = RProduccion(1)
                 '             End If
                 '             If IsNull(RProduccion(2)) Then
                 '                   VPDN = 0
                 '             Else
                 '                   VPDN = RProduccion(2)
                 '             End If
                 '          Else
                 '                   VPCN = 0
                 '                   VPNCN = 0
                 '                   VPDN = 0
                 '          End If
'________________________________________________________________________________________________________________________
'________________________________________________________________________________________________________________________
                                                                        
                        'EL TOTAL DE LA PRODUCCION ES LA SUMA DEL PRODUCTO CONFORME Y NO CONFORME NO INCLUYE EL DESPERDICIO
                        'TOTAL PRODUCCION
                        VTotalProduccion = VPCD + VPNCD + VPCN + VPNCN
                        'TOTAL PRODUCCION DE DIA
                        VTotalProduccionD = VPCD + VPNCD
                        'TOTAL PRODUCCION DE NOCHE
                        VTotalProduccionN = VPCN + VPNCN
                        
                        'SELECCIONA LA VELOCIDAD TEORICA DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaNoche
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaDia
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = ((VVelocidadTeoricaDia + VVelocidadTeoricaNoche) / 2)
                        ElseIf (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = 0
                        End If
                        
                        'SELECCIONA LA VELOCIDAD REAL DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadRealDia <= 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = VVelocidadRealNoche
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = VVelocidadRealDia
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = ((VVelocidadRealDia + VVelocidadRealNoche) / 2)
                        ElseIf (VVelocidadRealDia <= 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = 0
                        End If
                        
                                                                                                                                                                        
                                                                            
  'CONVIERTE LAS VARIABLES DE HORAS A MINUTOS PARA CALCULOS ________________________________________________________
                            VTiempoProgramadoD = VTiempoProgramadoD * 60
                            VParosSD = VParosSD * 60
                            VParosND = VParosND * 60
                        'NOCHE
                            VTiempoProgramadoN = VTiempoProgramadoN * 60
                            VParosSN = VParosSN * 60
                            VParosNN = VParosNN * 60
                        
                                        
'____________________________________________________________ FACTORES __________________________________________________
                                       
                'DIA
                'FACTOR 1______________________________________________________________________
                            VFactor1D = VTiempoProgramadoD - VParosND
                            VFactor1D = VFactor1D - VParosSD
                            VFactor1D = VFactor1D * VVelocidadRealDia
                            If VFactor1D = 0 Then
                            Else
                               VFactor1D = VTotalProduccionD / VFactor1D
                            End If
                '
                            'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                            'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                            If VFactor1D = 0 Then
                                VFactor1D = 1
                            End If
                '
                '
                ''FACTOR 2______________________________________________________________________
                '            If VPNCD > VTotalProduccionD Then
                '                VFactor2D = 0
                '            Else
                '                    VFactor2D = VTotalProduccionD - VPNCD
                '                    If VTotalProduccionD = 0 Then
                '                    Else
                '                       VFactor2D = VFactor2D / VTotalProduccionD
                '                    End If
                '            End If
                '
                ''FACTOR 3______________________________________________________________________
                '            If VPDD > VTotalProduccionD Then
                '                VFactor3D = 0
                '            Else
                '                    VFactor3D = VTotalProduccionD - VPDD
                '                    If VTotalProduccionD = 0 Then
                '                    Else
                '                       VFactor3D = VFactor3D / VTotalProduccionD
                '                    End If
                '            End If
                '
                ''FACTOR 4______________________________________________________________________
                '            VFactor4D = VTiempoProgramadoD - VParosND
                '            VFactor4D = VFactor4D - VParosSD
                '            If (VTiempoProgramadoD - VParosND) = 0 Then
                '            Else
                '                VFactor4D = (VFactor4D / (VTiempoProgramadoD - VParosND))
                '            End If
                '
                '
                ''FACTOR 5______________________________________________________________________
                            If VVelocidadTeoricaDia = 0 Then
                                VFactor5D = 0
                            Else
                                VFactor5D = VVelocidadRealDia / VVelocidadTeoricaDia
                            End If
                
                            'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                                    If VFactor5D = 0 Then
                                        VFactor5D = 1
                                    End If
                
               
                'EFICIENCIA REAL DEL TURNO DE DIA
                             'VEficienciaRealD = VFactor1D * VFactor2D * VFactor3D * VFactor4D * VFactor5D * 100
                             
                             'Set RBuscaEficiencia = New ADODB.Recordset
                        '    if GOrigenDeDatos = "AmaproAccess" Then
                        '        Call Abrir_Recordset(RBuscaEficiencia, "Select Eficiencia From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                        '    Else
                        '        Call Abrir_Recordset(RBuscaEficiencia, "Select Eficiencia From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                        '    End If
                        '
                        '   If RBuscaEficiencia.RecordCount > 0 Then
                        '        If IsNull(RBuscaEficiencia(0)) Then
                        '            VEficienciaRealD = 0
                        '        Else
                        '            VEficienciaRealD = RBuscaEficiencia(0)
                        '        End If
                        '   Else
                        '        VEficienciaRealD = 0
                        '   End If
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
                                        
                'NOCHE
                'SI EL TIEMPO PROGRAMADO ES CER0 NO HACE NADA
                If VTiempoProgramadoN > 0 Then
                                         
                            'EFICIENCIA DEL TURNO DE NOCHE
                                                            
                '            'FACTOR 1___________________________________________________________
                                    VFactor1N = VTiempoProgramadoN - VParosNN
                                    VFactor1N = VFactor1N - VParosSN
                                    VFactor1N = VFactor1N * VVelocidadRealNoche
                                    If VFactor1N = 0 Then
                                    Else
                                       VFactor1N = VTotalProduccionN / VFactor1N
                                    End If
                '
                '                    'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                '                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                '                    If VFactor1N = 0 Then
                '                        VFactor1N = 1
                '                    End If
               '
               '
               '             'FACTOR 2__________________________________________________________
               '                     If VPNCN > VTotalProduccionN Then
               '                         VFactor2N = 0
               '                     Else
               '                         VFactor2N = VTotalProduccionN - VPNCN
               '                         If VTotalProduccionN = 0 Then
               '                         Else
               '                             VFactor2N = (VFactor2N / VTotalProduccionN)
               '                         End If
               '                     End If
               '
              ''              'FACTOR 3___________________________________________________________
               '                     If VPDN > VTotalProduccionN Then
               '                         VFactor3N = 0
               '                     Else
               '                         VFactor3N = VTotalProduccionN - VPDN
              '                          If VTotalProduccionN = 0 Then
              '                          Else
              '                              VFactor3N = VFactor3N / VTotalProduccionN
              '                          End If
              '                      End If
              '
                            'FACTOR 4___________________________________________________________
              '                      VFactor4N = VTiempoProgramadoN - VParosNN
              '                      VFactor4N = VFactor4N - VParosSN
              '                      If (VTiempoProgramadoN - VParosNN) = 0 Then
              '                      Else
              '                          VFactor4N = (VFactor4N / (VTiempoProgramadoN - VParosNN))
              '                      End If
              '
              '              'FACTOR 5___________________________________________________________
                                    If VVelocidadTeoricaNoche = 0 Then
                                        VFactor5N = 0
                                    Else
                                        VFactor5N = VVelocidadRealNoche / VVelocidadTeoricaNoche
                                    End If
                                   
                                   'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
              '                      If VFactor5N = 0 Then
              '                          VFactor5N = 1
              '                      End If
              
              '          Set RBuscaEficiencia = New ADODB.Recordset
              '              If GOrigenDeDatos = "AmaproAccess" Then
              '                  Call Abrir_Recordset(RBuscaEficiencia, "Select Eficiencia From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
              '              Else
              '                  Call Abrir_Recordset(RBuscaEficiencia, "Select Eficiencia From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
              '              End If
              '
              '             If RBuscaEficiencia.RecordCount > 0 Then
              '                  If IsNull(RBuscaEficiencia(0)) Then
              '                      VEficienciaRealN = 0
              '                  Else
              '                      VEficienciaRealN = RBuscaEficiencia(0)
              '                  End If
              '             Else
              '                  VEficienciaRealN = 0
              '             End If
                Else
                        VFactor1N = 0
                        VFactor2N = 0
                        VFactor3N = 0
                        VFactor4N = 0
                        VFactor5N = 0
                End If
                           
                           'EFICIENCIA REAL DEL TURNO DE NOCHE
                            'VEficienciaRealN = VFactor1N * VFactor2N * VFactor3N * VFactor4N * VFactor5N * 100
                            
                                                            
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
                        
                        
                '% DE LINEA___________________________________________________________________________________________
                '_____________________________________________________________________________________________________
                                        
                            'FACTOR 1_________________________________________________________________________________
                                    VFactor1DN = (VTiempoProgramadoD - (VParosND + VParosSD)) + (VTiempoProgramadoN - (VParosNN + VParosSN))
                                    VFactor1DN = VFactor1DN * VVelocidadRealLinea
                                    If VFactor1DN = 0 Then
                                    Else
                                        VFactor1DN = VTotalProduccion / VFactor1DN
                                    End If
                                    
                                    'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                                    If VFactor1DN = 0 Then
                                        VFactor1DN = 1
                                    End If
                                                        
                            'FACTOR 2_________________________________________________________________________________
                                    If (VPNCD + VPNCN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor2DN = VTotalProduccion - (VPNCD + VPNCN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor2DN = VFactor2DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 3_________________________________________________________________________________
                                    If (VPDD + VPDN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor3DN = VTotalProduccion - (VPDD + VPDN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor3DN = VFactor3DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 4_________________________________________________________________________________
                                    VFactor4DN = (((VTiempoProgramadoD - VParosND) - VParosSD) + ((VTiempoProgramadoN - VParosNN) - VParosSN))
                                    If (VTiempoProgramadoD + VTiempoProgramadoN) = 0 Then
                                    ElseIf ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)) = 0 Then
                                    Else
                                        VFactor4DN = (VFactor4DN / ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)))
                                    End If
                                    
                            'FACTOR 5___________________________________________________________
                                    If VVelocidadTeoricaLinea = 0 Then
                                        VFactor5DN = 0
                                    Else
                                        VFactor5DN = VVelocidadRealLinea / VVelocidadTeoricaLinea
                                    End If
                                    
                                    'SI EL FACTOR 5 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 5
                                    If VFactor5DN = 0 Then
                                        VFactor5DN = 1
                                    End If
                                                                                
                            'EFICIENCIA DE LINEA______________________________________________________________________
                                    VPorcentajeLinea = VFactor1DN * VFactor2DN * VFactor3DN * VFactor4DN * VFactor5DN * 100
                                    
                                   
                '% DE RECHAZO DIA__________________________________________________________________________________________
                '______________________________________________________________________________________________________
                                    VPorcentajeRechazoD = VPNCD
                                    If VTotalProduccionD = 0 Then
                                    Else
                                        VPorcentajeRechazoD = VPorcentajeRechazoD / VTotalProduccionD
                                        VPorcentajeRechazoD = VPorcentajeRechazoD * 100
                                    End If
                '% DE RECHAZO NOCHE__________________________________________________________________________________________
                '______________________________________________________________________________________________________
                                    VPorcentajeRechazoN = VPNCN
                                    If VTotalProduccionN = 0 Then
                                    Else
                                        VPorcentajeRechazoN = VPorcentajeRechazoN / VTotalProduccionN
                                        VPorcentajeRechazoN = VPorcentajeRechazoN * 100
                                    End If
                
                        
                        
                '% DE DESPERDICIO DIA________________________________________________________________________________
                '______________________________________________________________________________________________________
                
                                    VPorcentajeDesperdicioD = VPDD
                                    If VTotalProduccionD = 0 Then
                                    Else
                                        VPorcentajeDesperdicioD = VPorcentajeDesperdicioD / VTotalProduccionD
                                        VPorcentajeDesperdicioD = VPorcentajeDesperdicioD * 100
                                    End If
                
                '% DE DESPERDICIO NOCHE______________________________________________________________________________
                '______________________________________________________________________________________________________
                
                                    
                                    VPorcentajeDesperdicioN = VPDN
                                    If VTotalProduccionN = 0 Then
                                    Else
                                        VPorcentajeDesperdicioN = VPorcentajeDesperdicioN / VTotalProduccionN
                                        VPorcentajeDesperdicioN = VPorcentajeDesperdicioN * 100
                                    End If
                                    
                                    
'_______________________________________________________________________________________________________________________
'************************************************* HORAS A MINUTOS *****************************************************
'_______________________________________________________________________________________________________________________

                                                
                'CONVIERTE LAS VARIABLES A HORAS PARA IMPRIMIR LOS DATOS
                                        
                            'DIA______________________________________________________________________
                                    If VTiempoProgramadoD = 0 Then
                                        VTiempoProgramadoD = 0
                                    Else
                                            VTiempoProgramadoD = VTiempoProgramadoD / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE DIA
                                    If VParosSD = 0 Then
                                        VParosSD = 0
                                    Else
                                        VParosSD = VParosSD / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE DIA
                                    If VParosND = 0 Then
                                        VParosND = 0
                                    Else
                                        VParosND = VParosND / 60
                                    End If
                                    
                            'NOCHE______________________________________________________________________
                                    If VTiempoProgramadoN = 0 Then
                                        VTiempoProgramadoN = 0
                                    Else
                                            VTiempoProgramadoN = VTiempoProgramadoN / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE NOCHE
                                    If VParosSN = 0 Then
                                        VParosSN = 0
                                    Else
                                        VParosSN = VParosSN / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE NOCHE
                                    If VParosNN = 0 Then
                                        VParosNN = 0
                                    Else
                                        VParosNN = VParosNN / 60
                                    End If
                                    
                                        
                                    If Err > 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                        Err.Clear
                                    End If
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________

                        
                        'BUSCA LA DESCRIPCION DE LA LINEA
                        Set RBuscaDescripcionLinea = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip From Lineas Where Linea = '" & VLinea & "'")
                            Else
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(VLinea) & "'")
                            End If
                        
                        'INICIALIZA UN RECORDSET PARA AGREGAR DATOS A LA BASE DE DATOS
                        'Set RReporteEficiencia = Db.OpenRecordset("Select * From ReporteEficienciaGrupos")
                                       
                                       If VTiempoProgramadoD > 0 Then
                                            If GOrigenDeDatos = "AmaproAccess" Then
                                                Conexion.Execute "Insert Into ReporteEficienciaGrupos (Linea, Grupo, Turno, Fecha, TiempoProgramado, TiempoRealProducidoD, ParosAfectanD, TiempoRealProducido, ParosNoAfectanD, ProductoConformeD, ProductoNoConformeD, DesperdicioD, TotalProduccion, EficienciaD, PorcentajeRechazo, PorcentajeDesperdicio, Factor1, Factor5, VelocidadTeoricaDia, VelocidadRealDia) VALUES('" & RBuscaDescripcionLinea(0) & "', '" & VGrupoDia & "', '1', #" & Format(VFechaInicial, "mm/dd/yyyy") & "#, " & VTiempoProgramadoD & ", " & VProduccionD & ", " & VParosSD & ", " & VProduccionD & ", " & VParosND & ", " & VPCD & ", " & VPNCD & ", " & VPDD & ", " & VTotalProduccionD & ", " & VEficienciaRealD & ", " & VPorcentajeRechazoD & ", " & VPorcentajeDesperdicioD & ", " & VFactor1D & ", " & VFactor5D & ", " & VVelocidadTeoricaDia & ", " & VVelocidadRealDia & ")"
                                            Else
                                                Conexion.Execute "Insert Into ReporteEficienciaGrupos (Linea, Grupo, Turno, Fecha, TiempoProgramado, TiempoRealProducidoD, ParosAfectanD, TiempoRealProducido, ParosNoAfectanD, ProductoConformeD, ProductoNoConformeD, DesperdicioD, TotalProduccion, EficienciaD, PorcentajeRechazo, PorcentajeDesperdicio, Factor1, Factor5, VelocidadTeoricaDia, VelocidadRealDia) VALUES('" & RBuscaDescripcionLinea(0) & "', '" & VGrupoDia & "', '1', To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & ", " & VTiempoProgramadoD & ", " & VProduccionD & ", " & VParosSD & ", " & VProduccionD & ", " & VParosND & ", " & VPCD & ", " & VPNCD & ", " & VPDD & ", " & VTotalProduccionD & ", " & VEficienciaRealD & ", " & VPorcentajeRechazoD & ", " & VPorcentajeDesperdicioD & ", " & VFactor1D & ", " & VFactor5D & ", " & VVelocidadTeoricaDia & ", " & VVelocidadRealDia & ")"
                                            End If
                                                                                       
                                                If Err <> 0 Then
                                                    MsgBox Err.Number & Err.Description
                                                    Err.Clear
                                                End If
                                    End If
                                                
                                    If VTiempoProgramadoN > 0 Then
                                                'AGREGA DATOS TURNO NOCHE
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Conexion.Execute "Insert Into ReporteEficienciaGrupos (Linea, Grupo, Turno, Fecha, TiempoProgramado, TiempoRealProducidoD, ParosAfectanD, TiempoRealProducido, ParosNoAfectanD, ProductoConformeD, ProductoNoConformeD, DesperdicioD, TotalProduccion, EficienciaD, PorcentajeRechazo, PorcentajeDesperdicio, Factor1, Factor5, VelocidadTeoricaDia, VelocidadRealDia) VALUES('" & RBuscaDescripcionLinea(0) & "', '" & VGrupoNoche & "', '2', #" & Format(VFechaInicial, "mm/dd/yyyy") & "#, " & VTiempoProgramadoN & ", " & VProduccionN & ", " & VParosSN & ", " & VProduccionN & ", " & VParosNN & ", " & VPCN & ", " & VPNCN & ", " & VPDN & ", " & VTotalProduccionN & ", " & VEficienciaRealN & ", " & VPorcentajeRechazoN & ", " & VPorcentajeDesperdicioN & ", " & VFactor1N & ", " & VFactor5N & ", " & VVelocidadTeoricaNoche & ", " & VVelocidadRealNoche & ")"
                                                Else
                                                    Conexion.Execute "Insert Into ReporteEficienciaGrupos (Linea, Grupo, Turno, Fecha, TiempoProgramado, TiempoRealProducidoD, ParosAfectanD, TiempoRealProducido, ParosNoAfectanD, ProductoConformeD, ProductoNoConformeD, DesperdicioD, TotalProduccion, EficienciaD, PorcentajeRechazo, PorcentajeDesperdicio, Factor1, Factor5, VelocidadTeoricaDia, VelocidadRealDia) VALUES('" & RBuscaDescripcionLinea(0) & "', '" & VGrupoNoche & "', '2', To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & ", " & VTiempoProgramadoN & ", " & VProduccionN & ", " & VParosSN & ", " & VProduccionN & ", " & VParosNN & ", " & VPCN & ", " & VPNCN & ", " & VPDN & ", " & VTotalProduccionN & ", " & VEficienciaRealN & ", " & VPorcentajeRechazoN & ", " & VPorcentajeDesperdicioN & ", " & VFactor1N & ", " & VFactor5N & ", " & VVelocidadTeoricaNoche & ", " & VVelocidadRealNoche & ")"
                                                End If
                                                
                                       
                                   End If
                                    
                                    If Err > 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                            Err.Clear
                                    End If
                                    
                            VFechaInicial = VFechaInicial + 1
                                    
                    Else ' IF DE CONDICION DE LA QUE NO ENCUENTRA DATOS EN RANGO DE FECHAS
                            
                            'A LA FECHA ACTUAL LE SUMA 1 PARA INCREMENTAR LA FECHA
                            'LA EFICIENCIA SE SACA POR DIA
                            VFechaInicial = VFechaInicial + 1
                    End If
                
                        
                Loop
                
                
                'SIGUE AL SIGUIENTE REGISTRO DE LINEA
                RSeleccionaLineas.MoveNext
  Loop
                'SI PIDE EL REPORTE DE LA EFICIENCIA
                If (OptEfiGru.Value = True Or OptEfiGruRes.Value = True Or OptEfiGruResEmp = True) Then
                    GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "                        Reporte del " & DTPFecEfi.Value & " Al " & DTPFecEfiFin.Value
                End If
            
                
                If OptEfiGru.Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "EficienciaEnGrupos.rpt"
                                Else
                                     GNombreReporte = "EficienciaEnGruposO.rpt"
                                End If
                    
                ElseIf OptEfiGruRes.Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "EficienciaEnGruposResumen.rpt"
                                Else
                                     GNombreReporte = "EficienciaEnGruposResumenO.rpt"
                                End If
                ElseIf OptEfiGruResEmp.Value = True Then
                                If OptEfi.Item(0).Value = True Then
                                    GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} >= #" & Format(DTPFecEfi.Value, "mm/dd/yyyy") & "# And {EncabezadoCapturaParos.Fecha} <= #" & Format(DTPFecEfiFin.Value) & "# And {EncabezadoCapturaParos.Linea} = '" & TxtLinea.Text & "'"
                                ElseIf OptEfi.Item(1).Value = True Then
                                    GCriteriaReporte = "{EncabezadoCapturaParos.Fecha} >= #" & Format(DTPFecEfi.Value, "mm/dd/yyyy") & "# And {EncabezadoCapturaParos.Fecha} <= #" & Format(DTPFecEfiFin.Value) & "# And {EncabezadoCapturaParos.Linea} = {Lineas.Linea} And {Lineas.Grupo} = '" & TxtLinea.Text & "'"
                                End If
                        
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "EficienciaEnEquiposEmpleados.rpt"
                                Else
                                     GNombreReporte = "EficienciaEnEquiposEmpleadosO.rpt"
                                End If
                                
                End If
                


End Sub


Public Sub ReporteEjecutivo()
On Error Resume Next
                            
                            
                            
                                                        

            
            'SELECCIONA LAS LINEAS DE ACUERDO A LA OPCION
            If OptEfi.Item(0).Value = True = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Linea = '" & TxtLinea.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                            End If
            ElseIf OptEfi.Item(1).Value = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Grupo = '" & TxtLinea.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Grupo) = '" & UCase(TxtLinea.Text) & "'")
                            End If
            End If
             
             If RSeleccionaLineas.RecordCount > 0 Then
             Else
                    MsgBox "Linea O Lineas No Existen ", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
             End If
  
  
  
  'CREA UN CICLO CON LAS LINEAS POSIBLES DE ACUERDO A LA OPCION ELEGIDA
  Do Until RSeleccionaLineas.EOF
                        
                    'ASIGNA LA LINEA QUE ES SELECCIONADA
                    VLinea = RSeleccionaLineas!Linea
                            
                    'FECHA DE INICIO DEL RANGO
                    VFechaInicial = Format(DTPFecEfi.Value, "dd/mm/yyyy")
                    'FECHA DEL FINAL DEL RANGO
                    VFechaFinal = Format(DTPFecEfiFin.Value, "dd/mm/yyyy")
                        
                
                Do Until VFechaInicial > VFechaFinal
                
                            VToneladasPC = 0
                            VToneladasPNC = 0
                            VToneladasDes = 0
                            VProduccionPC = 0
                            VProduccionPNC = 0
                            VProduccionDes = 0
                            VUnidadesPC = 0
                            VUnidadesPNC = 0
                            VUnidadesDes = 0
                        
                        
'VERIFICA SI HAY DATOS EN LA PRESENTE FECHA Y SI NO HAY PASA A LA SIGUIENTE FECHA
'ESTO NOS SIRVE PARA CUANDO SAQUEMOS EL REPORTE DE EFICIENCIA NO TOME EN CUENTA LOS DIAS QUE
'NO SE TRABAJO PORQUE AFECTA LA EFICIENCIA DE LINEA Y PLANTA
                
                Set RCapturaParos = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Linea = '" & VLinea & "'")
                    Else
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                    End If
                    If RCapturaParos.RecordCount > 0 Then
                        'NO HACE NADA SI HAY DATOS ESTA BIEN
                    
                                
                  
                  '*******  TIEMPO PROGRAMADO Y VELOCIDAD **************************************************
  
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE DIA
                         Set RTiempoProgramadoD = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoD.RecordCount > 0 Then
                                            VTiempoProgramadoD = RTiempoProgramadoD!HorasProgramadas
                                            VVelocidadTeoricaDia = RTiempoProgramadoD!VelocidadTeorica
                                            VVelocidadRealDia = RTiempoProgramadoD!VelocidadReal
                                            VGrupoDia = RTiempoProgramadoD!Grupo
                                            VParosND = RTiempoProgramadoD!ParoN / 60
                                            VParosSD = RTiempoProgramadoD!Paros / 60
                                            VProduccionD = RTiempoProgramadoD!ParoP / 60
                                            VPCD = RTiempoProgramadoD!ProductoConforme
                                            VPNCD = RTiempoProgramadoD!ProductoNoConforme
                                            VPDD = RTiempoProgramadoD!Desperdicio
                                            VEficienciaRealD = RTiempoProgramadoD!Eficiencia
                                            VParosCFD = RTiempoProgramadoD!ParoCF
                                            VParosMPD = RTiempoProgramadoD!ParoMP
                                Else
                                            VTiempoProgramadoD = 0
                                            VVelocidadTeoricaDia = 0
                                            VVelocidadRealDia = 0
                                            VGrupoDia = ""
                                            VParosND = 0
                                            VParosSD = 0
                                            VProduccionD = 0
                                            VPCD = 0
                                            VPNCD = 0
                                            VPDD = 0
                                            VEficienciaRealD = 0
                                            VParosCFD = 0
                                            VParosMPD = 0
                             End If
                                               
                        
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE NOCHE
                         Set RTiempoProgramadoN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoN.RecordCount > 0 Then
                                            VTiempoProgramadoN = RTiempoProgramadoN!HorasProgramadas
                                            VVelocidadTeoricaNoche = RTiempoProgramadoN!VelocidadTeorica
                                            VVelocidadRealNoche = RTiempoProgramadoN!VelocidadReal
                                            VGrupoNoche = RTiempoProgramadoN!Grupo
                                            VParosNN = RTiempoProgramadoN!ParoN / 60
                                            VParosSN = RTiempoProgramadoN!Paros / 60
                                            VProduccionN = RTiempoProgramadoN!ParoP / 60
                                            VPCN = RTiempoProgramadoN!ProductoConforme
                                            VPNCN = RTiempoProgramadoN!ProductoNoConforme
                                            VPDN = RTiempoProgramadoN!Desperdicio
                                            VEficienciaRealN = RTiempoProgramadoN!Eficiencia
                                            VParosCFN = RTiempoProgramadoN!ParoCF
                                            VParosMPN = RTiempoProgramadoN!ParoMP
                             Else
                                            VTiempoProgramadoN = 0
                                            VVelocidadTeoricaNoche = 0
                                            VVelocidadRealNoche = 0
                                            VGrupoNoche = ""
                                            VParosNN = 0
                                            VParosSN = 0
                                            VProduccionN = 0
                                            VPCN = 0
                                            VPNCN = 0
                                            VPDN = 0
                                            VEficienciaRealN = 0
                                            VParosCFN = RTiempoProgramadoN!ParoCF
                                            VParosMPN = RTiempoProgramadoN!ParoMP
                             End If
                                                
                                                                                                
                 
                        
                        
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE DIA
                        VTiempoRealProducidoD = VTiempoProgramadoD - VParosND
                                                    
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE NOCHE
                        If VTiempoProgramadoN = 0 Then
                            VTiempoRealProducidoN = 0
                        Else
                            VTiempoRealProducidoN = VTiempoProgramadoN - VParosNN
                        End If
                                                    
                        'HORAS PRODUCIDAS POR LOS 2 TURNOS
                            VHorasProducidasDN = Format(VTiempoRealProducidoD + VTiempoRealProducidoN, "#,###,##0.00")
                                                    
                        'TOTAL DE PAROS S "NO AFECTAN"
                            VParosDN = Format(VParosND + VParosNN, "#,###,##0.00")
                        
    
'________________________________________________________________________________________________________________________
         'PARA EL REPORTE EJECUTIVO
         'BUSCAMOS EL DETALLE DE LA PRODUCCION EN EL RANGO DE FECHAS INDICADO
                            Set RBuscaDetalleProduccion = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaDetalleProduccion, "Select DP.Orden, DP.ProductoConforme, DP.ProductoNoConforme, DP.Desperdicio From DetalleProduccionPorOrden DP, EncabezadoCapturaParos EP Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Linea = '" & VLinea & "' And EP.Documento = DP.Documento")
                                Else
                                    Call Abrir_Recordset(RBuscaDetalleProduccion, "Select DP.Orden, DP.ProductoConforme, DP.ProductoNoConforme, DP.Desperdicio From DetalleProduccionPorOrden DP, EncabezadoCapturaParos EP Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = '" & UCase(VLinea) & "' And EP.Documento = DP.Documento")
                                End If
                                If RBuscaDetalleProduccion.RecordCount > 0 Then
                                    Do Until RBuscaDetalleProduccion.EOF
                                                'ASIGNAMOS VARIABLES
                                                VOrdenDetalle = RBuscaDetalleProduccion(0)
                                                VProduccionPC = RBuscaDetalleProduccion(1)
                                                VProduccionPNC = RBuscaDetalleProduccion(2)
                                                VProduccionDes = RBuscaDetalleProduccion(3)
                                                
                                                'AHORA BUSCAMOS LA ORDEN
                                                Set RBuscaOrden = New ADODB.Recordset
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Call Abrir_Recordset(RBuscaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & VOrdenDetalle & "'")
                                                    Else
                                                        Call Abrir_Recordset(RBuscaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(VOrdenDetalle) & "'")
                                                    End If
                                                    'ASIGNAMOS LA FICHA TECNICA QUE USA LA ORDEN
                                                    If RBuscaOrden.RecordCount > 0 Then
                                                        VFichaTecnicaOrden = RBuscaOrden!FichaTecnica
                                                            'BUSCAMOS LA FICHA TECNICA PARA OBTENER EL PESO
                                                            Set RBuscaFichaTecnica = New ADODB.Recordset
                                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select PesoxUnidad From FichaTecnica Where Esp_Tec = '" & VFichaTecnicaOrden & "'")
                                                                Else
                                                                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select PesoxUnidad From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(VFichaTecnicaOrden) & "'")
                                                                End If
                                                                'ASIGNAMOS EL PESO
                                                                If RBuscaFichaTecnica.RecordCount > 0 Then
                                                                    VPesoFichaTecnica = RBuscaFichaTecnica!PesoxUnidad
                                                                Else
                                                                    VPesoFichaTecnica = 0
                                                                End If
                                                    Else
                                                        VFichaTecnicaOrden = ""
                                                    End If
                                                    'ASIGNAMOS EL PESO A LAS VARIABLES
                                                    VToneladasCalculoPC = ((VProduccionPC * VPesoFichaTecnica) / 1000)
                                                    VToneladasPC = VToneladasPC + VToneladasCalculoPC
                                                    VToneladasCalculoPNC = ((VProduccionPNC * VPesoFichaTecnica) / 1000)
                                                    VToneladasPNC = VToneladasPNC + VToneladasCalculoPNC
                                                    VToneladasCalculoDes = ((VProduccionDes * VPesoFichaTecnica) / 1000)
                                                    VToneladasDes = VToneladasDes + VToneladasCalculoDes
                                                    
                                                    'UNIDADES
                                                    VUnidadesPC = VUnidadesPC + VProduccionPC
                                                    VUnidadesPNC = VUnidadesPNC + VProduccionPNC
                                                    VUnidadesDes = VUnidadesDes + VProduccionDes
                                        
                                        RBuscaDetalleProduccion.MoveNext
                                    Loop
                                Else
                                End If
                                                            
'________________________________________________________________________________________________________________________
'________________________________________________________________________________________________________________________
                        
                        
                        
                        'EL TOTAL DE LA PRODUCCION ES LA SUMA DEL PRODUCTO CONFORME Y NO CONFORME NO INCLUYE EL DESPERDICIO
                        'TOTAL PRODUCCION
                        VTotalProduccion = VPCD + VPNCD + VPCN + VPNCN
                        'TOTAL PRODUCCION DE DIA
                        VTotalProduccionD = VPCD + VPNCD
                        'TOTAL PRODUCCION DE NOCHE
                        VTotalProduccionN = VPCN + VPNCN
                        
                        'SELECCIONA LA VELOCIDAD TEORICA DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaNoche
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaDia
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = ((VVelocidadTeoricaDia + VVelocidadTeoricaNoche) / 2)
                        ElseIf (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = 0
                        End If
                        
                        'SELECCIONA LA VELOCIDAD REAL DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadRealDia <= 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = VVelocidadRealNoche
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = VVelocidadRealDia
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = ((VVelocidadRealDia + VVelocidadRealNoche) / 2)
                        ElseIf (VVelocidadRealDia <= 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = 0
                        End If
                        
                                                                                                                                                                        
                                                                            
  'CONVIERTE LAS VARIABLES DE HORAS A MINUTOS PARA CALCULOS ________________________________________________________
                        
                            VTiempoProgramadoD = VTiempoProgramadoD * 60
                            VParosSD = VParosSD * 60
                            VParosND = VParosND * 60
                            
                        'NOCHE
                            VTiempoProgramadoN = VTiempoProgramadoN * 60
                            VParosSN = VParosSN * 60
                            VParosNN = VParosNN * 60
                                        
'_______________________________________________________________________________________________________________________
                                        
                'NOCHE
                'SI EL TIEMPO PROGRAMADO ES CER0 NO HACE NADA
                If VTiempoProgramadoN > 0 Then
                                         
                Else
                        VFactor1N = 0
                        VFactor2N = 0
                        VFactor3N = 0
                        VFactor4N = 0
                        VFactor5N = 0
                End If
'_______________________________________________________________________________________________________________________
                        
                        
                '% DE LINEA___________________________________________________________________________________________
                '_____________________________________________________________________________________________________
                                        
                            'FACTOR 1_________________________________________________________________________________
                                    VFactor1DN = (VTiempoProgramadoD - (VParosND + VParosSD)) + (VTiempoProgramadoN - (VParosNN + VParosSN))
                                    VFactor1DN = VFactor1DN * VVelocidadRealLinea
                                    If VFactor1DN = 0 Then
                                    Else
                                        VFactor1DN = VTotalProduccion / VFactor1DN
                                    End If
                                    
                                    'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                                    If VFactor1DN = 0 Then
                                        VFactor1DN = 1
                                    End If
                                                        
                            'FACTOR 2_________________________________________________________________________________
                                    If (VPNCD + VPNCN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor2DN = VTotalProduccion - (VPNCD + VPNCN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor2DN = VFactor2DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 3_________________________________________________________________________________
                                    If (VPDD + VPDN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor3DN = VTotalProduccion - (VPDD + VPDN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor3DN = VFactor3DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 4_________________________________________________________________________________
                                    VFactor4DN = (((VTiempoProgramadoD - VParosND) - VParosSD) + ((VTiempoProgramadoN - VParosNN) - VParosSN))
                                    If (VTiempoProgramadoD + VTiempoProgramadoN) = 0 Then
                                    ElseIf ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)) = 0 Then
                                    Else
                                        VFactor4DN = (VFactor4DN / ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)))
                                    End If
                                    
                            'FACTOR 5___________________________________________________________
                                    If VVelocidadTeoricaLinea = 0 Then
                                        VFactor5DN = 0
                                    Else
                                        VFactor5DN = VVelocidadRealLinea / VVelocidadTeoricaLinea
                                    End If
                                    
                                    'SI EL FACTOR 5 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 5
                                    If VFactor5DN = 0 Then
                                        VFactor5DN = 1
                                    End If
                                                                                
                            'EFICIENCIA DE LINEA______________________________________________________________________
                                    VPorcentajeLinea = VFactor1DN * VFactor2DN * VFactor3DN * VFactor4DN * VFactor5DN * 100
                                   
                '% DE RECHAZO__________________________________________________________________________________________
                '______________________________________________________________________________________________________
                                    VPorcentajeRechazo = VPNCD + VPNCN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeRechazo = VPorcentajeRechazo / VTotalProduccion
                                        VPorcentajeRechazo = VPorcentajeRechazo * 100
                                    End If
                        
                        
                '% DE DESPERDICIO______________________________________________________________________________________
                '______________________________________________________________________________________________________
                
                                    VPorcentajeDesperdicio = VPDD + VPDN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio / VTotalProduccion
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio * 100
                                    End If
                                    
'_______________________________________________________________________________________________________________________
'************************************************* HORAS A MINUTOS *****************************************************
'_______________________________________________________________________________________________________________________

                                                
                'CONVIERTE LAS VARIABLES A HORAS PARA IMPRIMIR LOS DATOS
                                        
                            'DIA______________________________________________________________________
                                    If VTiempoProgramadoD = 0 Then
                                        VTiempoProgramadoD = 0
                                    Else
                                        VTiempoProgramadoD = VTiempoProgramadoD / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE DIA
                                    If VParosSD = 0 Then
                                        VParosSD = 0
                                    Else
                                        VParosSD = VParosSD / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE DIA
                                    If VParosND = 0 Then
                                        VParosND = 0
                                    Else
                                        VParosND = VParosND / 60
                                    End If
                                    
                            'NOCHE______________________________________________________________________
                                    If VTiempoProgramadoN = 0 Then
                                        VTiempoProgramadoN = 0
                                    Else
                                        VTiempoProgramadoN = VTiempoProgramadoN / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE NOCHE
                                    If VParosSN = 0 Then
                                        VParosSN = 0
                                    Else
                                        VParosSN = VParosSN / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE NOCHE
                                    If VParosNN = 0 Then
                                        VParosNN = 0
                                    Else
                                        VParosNN = VParosNN / 60
                                    End If
                                    
                                        
                                    If Err > 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                        Err.Clear
                                    End If
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________

                        
                        'BUSCA LA DESCRIPCION DE LA LINEA
                        Set RBuscaDescripcionLinea = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip, UnidadMedida, TipoProducto From Lineas Where Linea = '" & VLinea & "'")
                            Else
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip, UnidadMedida, TipoProducto From Lineas Where UPPER(Linea) = '" & UCase(VLinea) & "'")
                            End If
                        
                           
                                                VTexto = "'" & RBuscaDescripcionLinea(0) & "', "
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "#" & Format(VFechaInicial, "mm/dd/yyyy") & "#, " 'FECHA
                                                Else 'ORACLE
                                                     VTexto = VTexto & "To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & ", " 'FECHA
                                                End If
                                                VTexto = VTexto & VTiempoProgramadoD + VTiempoProgramadoN & ", "
                                                VTexto = VTexto & VProduccionD & ", "
                                                VTexto = VTexto & VProduccionN & ", "
                                                VTexto = VTexto & VParosSD & ", "
                                                VTexto = VTexto & VParosSN & ", "
                                                VTexto = VTexto & VProduccionD + VProduccionN & ", " 'VTiempoRealProducidoD + VTiempoRealProducidoN
                                                VTexto = VTexto & VParosND & ", "
                                                VTexto = VTexto & VParosNN & ", "
                                                VTexto = VTexto & VPCD & ", "
                                                VTexto = VTexto & VPNCD & ", "
                                                VTexto = VTexto & VPDD & ", "
                                                VTexto = VTexto & VPCN & ", "
                                                VTexto = VTexto & VPNCN & ", "
                                                VTexto = VTexto & VPDN & ", "
                                                VTexto = VTexto & VTotalProduccion & ", "
                                                VTexto = VTexto & VEficienciaRealD & ", "
                                                VTexto = VTexto & VEficienciaRealN & ", "
                                                VTexto = VTexto & VPorcentajeLinea & ", "
                                                VTexto = VTexto & VPorcentajeRechazo & ", "
                                                VTexto = VTexto & Format(VPorcentajeDesperdicio, "#,###,##0.00") & ", "
                                                VTexto = VTexto & VFactor1DN & ", "
                                                VTexto = VTexto & VFactor5DN & ", "
                                                VTexto = VTexto & VVelocidadTeoricaDia & ", "
                                                VTexto = VTexto & VVelocidadRealDia & ", "
                                                VTexto = VTexto & VVelocidadTeoricaNoche & ", "
                                                VTexto = VTexto & VVelocidadRealNoche & ", '"
                                                VTexto = VTexto & RBuscaDescripcionLinea(1) & "', "
                                                VTexto = VTexto & VToneladasPC & ", "
                                                VTexto = VTexto & VToneladasPNC & ", "
                                                VTexto = VTexto & Format(VToneladasDes, "#,###,##0.00") & ", "
                                                VTexto = VTexto & VUnidadesPC & ", "
                                                VTexto = VTexto & VUnidadesPNC & ", "
                                                VTexto = VTexto & VUnidadesDes & ", '"
                                                VTexto = VTexto & RBuscaDescripcionLinea(2) & "', "
                                                VTexto = VTexto & (VParosCFD + VParosCFN) & ", "
                                                VTexto = VTexto & (VParosMPD + VParosMPN)
                           
                                                Conexion.Execute "Insert Into ReporteEjecutivoDia Values(" & VTexto & ")"
                            
                                    If Err > 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                            Err.Clear
                                    End If
                                    
                                    
                                    
                            VFechaInicial = VFechaInicial + 1
                                    
                    Else ' IF DE CONDICION DE LA QUE NO ENCUENTRA DATOS EN RANGO DE FECHAS
                            
                            'A LA FECHA ACTUAL LE SUMA 1 PARA INCREMENTAR LA FECHA
                            'LA EFICIENCIA SE SACA POR DIA
                            VFechaInicial = VFechaInicial + 1
                    End If
                
                        
                Loop
                
                
                'SIGUE AL SIGUIENTE REGISTRO DE LINEA
                RSeleccionaLineas.MoveNext
  Loop
  
  
  
                
'EMPIEZA TODO EL PROCESO OTRA VEZ PARA SACAR EL ACUMULADO
'_____________________________________________________________________________________________
'_____________________________________________________________________________________________


            'SELECCIONA LAS LINEAS DE ACUERDO A LA OPCION
            If OptEfi.Item(0).Value = True = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Linea = '" & TxtLinea.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
                            End If
            ElseIf OptEfi.Item(1).Value = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Grupo = '" & TxtLinea.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Grupo) = '" & UCase(TxtLinea.Text) & "'")
                            End If
            End If
             
             If RSeleccionaLineas.RecordCount > 0 Then
             Else
                    MsgBox "Linea O Lineas No Existen ", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
             End If
  
  
  
  'CREA UN CICLO CON LAS LINEAS POSIBLES DE ACUERDO A LA OPCION ELEGIDA
  Do Until RSeleccionaLineas.EOF
                        
                    'ASIGNA LA LINEA QUE ES SELECCIONADA
                    VLinea = RSeleccionaLineas!Linea
                            
                    'FECHA DE INICIO DEL RANGO
                    VFechaInicial = DTPFecEfiAcu.Value
                    'FECHA DEL FINAL DEL RANGO
                    VFechaFinal = DTPFecEfiFinAcu.Value
                        
                
                Do Until VFechaInicial > VFechaFinal
                            
                            VAToneladasPC = 0
                            VAToneladasPNC = 0
                            VAToneladasDes = 0
                            VAProduccionPC = 0
                            VAProduccionPNC = 0
                            VAProduccionDes = 0
                            VAUnidadesPC = 0
                            VAUnidadesPNC = 0
                            VAUnidadesDes = 0
                            
                        
                        
'VERIFICA SI HAY DATOS EN LA PRESENTE FECHA Y SI NO HAY PASA A LA SIGUIENTE FECHA
'ESTO NOS SIRVE PARA CUANDO SAQUEMOS EL REPORTE DE EFICIENCIA NO TOME EN CUENTA LOS DIAS QUE
'NO SE TRABAJO PORQUE AFECTA LA EFICIENCIA DE LINEA Y PLANTA
                Set RCapturaParos = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Linea = '" & VLinea & "'")
                    Else
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                    End If
                    If RCapturaParos.RecordCount > 0 Then
                        'NO HACE NADA SI HAY DATOS ESTA BIEN
                    
                                
                  
                  '*******  TIEMPO PROGRAMADO Y VELOCIDAD **************************************************
  
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE DIA
                         Set RTiempoProgramadoD = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoD.RecordCount > 0 Then
                                            VTiempoProgramadoD = RTiempoProgramadoD!HorasProgramadas
                                            VVelocidadTeoricaDia = RTiempoProgramadoD!VelocidadTeorica
                                            VVelocidadRealDia = RTiempoProgramadoD!VelocidadReal
                                            VGrupoDia = RTiempoProgramadoD!Grupo
                                            VParosND = RTiempoProgramadoD!ParoN / 60
                                            VParosSD = RTiempoProgramadoD!Paros / 60
                                            VProduccionD = RTiempoProgramadoD!ParoP / 60
                                            VPCD = RTiempoProgramadoD!ProductoConforme
                                            VPNCD = RTiempoProgramadoD!ProductoNoConforme
                                            VPDD = RTiempoProgramadoD!Desperdicio
                                            VEficienciaRealD = RTiempoProgramadoD!Eficiencia
                                            VParosCFD = RTiempoProgramadoD!ParoCF
                                            VParosMPD = RTiempoProgramadoD!ParoMP
                                Else
                                            VTiempoProgramadoD = 0
                                            VVelocidadTeoricaDia = 0
                                            VVelocidadRealDia = 0
                                            VGrupoDia = ""
                                            VParosND = 0
                                            VParosSD = 0
                                            VProduccionD = 0
                                            VPCD = 0
                                            VPNCD = 0
                                            VPDD = 0
                                            VEficienciaRealD = 0
                                            VParosCFD = 0
                                            VParosMPD = 0
                             End If
                                               
                        
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE NOCHE
                         Set RTiempoProgramadoN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoN.RecordCount > 0 Then
                                            VTiempoProgramadoN = RTiempoProgramadoN!HorasProgramadas
                                            VVelocidadTeoricaNoche = RTiempoProgramadoN!VelocidadTeorica
                                            VVelocidadRealNoche = RTiempoProgramadoN!VelocidadReal
                                            VGrupoNoche = RTiempoProgramadoN!Grupo
                                            VParosNN = RTiempoProgramadoN!ParoN / 60
                                            VParosSN = RTiempoProgramadoN!Paros / 60
                                            VProduccionN = RTiempoProgramadoN!ParoP / 60
                                            VPCN = RTiempoProgramadoN!ProductoConforme
                                            VPNCN = RTiempoProgramadoN!ProductoNoConforme
                                            VPDN = RTiempoProgramadoN!Desperdicio
                                            VEficienciaRealN = RTiempoProgramadoN!Eficiencia
                                            VParosCFN = RTiempoProgramadoN!ParoCF
                                            VParosMPN = RTiempoProgramadoN!ParoMP
                             Else
                                            VTiempoProgramadoN = 0
                                            VVelocidadTeoricaNoche = 0
                                            VVelocidadRealNoche = 0
                                            VGrupoNoche = ""
                                            VParosNN = 0
                                            VParosSN = 0
                                            VProduccionN = 0
                                            VPCN = 0
                                            VPNCN = 0
                                            VPDN = 0
                                            VEficienciaRealN = 0
                                            VParosCFN = 0
                                            VParosMPN = 0
                             End If
                                                
                                                                                                
                 
                        
                        
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE DIA
                       VTiempoRealProducidoD = VTiempoProgramadoD - VParosND
                                                    
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE NOCHE
                        If VTiempoProgramadoN = 0 Then
                            VTiempoRealProducidoN = 0
                        Else
                            VTiempoRealProducidoN = VTiempoProgramadoN - VParosNN
                        End If
                                                    
                        'HORAS PRODUCIDAS POR LOS 2 TURNOS
                            VHorasProducidasDN = Format(VTiempoRealProducidoD + VTiempoRealProducidoN, "#,###,##0.00")
                                                    
                        'TOTAL DE PAROS S "NO AFECTAN"
                            VParosDN = Format(VParosND + VParosNN, "#,###,##0.00")
                        
    
'________________________________________________________________________________________________________________________
         'PARA EL REPORTE EJECUTIVO ACUMULADO
         'BUSCAMOS EL DETALLE DE LA PRODUCCION EN EL RANGO DE FECHAS INDICADO
                            
                            'BUSCAMOS SOLO UN DIA
                            Set RABuscaDetalleProduccion = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RABuscaDetalleProduccion, "Select DP.Orden, DP.ProductoConforme, DP.ProductoNoConforme, DP.Desperdicio From DetalleProduccionPorOrden DP, EncabezadoCapturaParos EP Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Linea = '" & VLinea & "' And EP.Documento = DP.Documento")
                                Else
                                    Call Abrir_Recordset(RABuscaDetalleProduccion, "Select DP.Orden, DP.ProductoConforme, DP.ProductoNoConforme, DP.Desperdicio From DetalleProduccionPorOrden DP, EncabezadoCapturaParos EP Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = '" & UCase(VLinea) & "' And EP.Documento = DP.Documento")
                                End If
                                    
                                If RABuscaDetalleProduccion.RecordCount > 0 Then
                                    Do Until RABuscaDetalleProduccion.EOF
                                                'ASIGNAMOS VARIABLES
                                                VAOrdenDetalle = RABuscaDetalleProduccion(0)
                                                VAProduccionPC = RABuscaDetalleProduccion(1)
                                                VAProduccionPNC = RABuscaDetalleProduccion(2)
                                                VAProduccionDes = RABuscaDetalleProduccion(3)
                                                
                                                'AHORA BUSCAMOS LA ORDEN
                                                Set RABuscaOrden = New ADODB.Recordset
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Call Abrir_Recordset(RABuscaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & VAOrdenDetalle & "'")
                                                    Else
                                                        Call Abrir_Recordset(RABuscaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(VAOrdenDetalle) & "'")
                                                    End If
                                                    'ASIGNAMOS LA FICHA TECNICA QUE USA LA ORDEN
                                                    If RABuscaOrden.RecordCount > 0 Then
                                                        VAFichaTecnicaOrden = RABuscaOrden!FichaTecnica
                                                            'BUSCAMOS LA FICHA TECNICA PARA OBTENER EL PESO
                                                            Set RABuscaFichaTecnica = New ADODB.Recordset
                                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                                    Call Abrir_Recordset(RABuscaFichaTecnica, "Select PesoxUnidad From FichaTecnica Where Esp_Tec = '" & VAFichaTecnicaOrden & "'")
                                                                Else
                                                                    Call Abrir_Recordset(RABuscaFichaTecnica, "Select PesoxUnidad From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(VAFichaTecnicaOrden) & "'")
                                                                End If
                                                                'ASIGNAMOS EL PESO
                                                                If RABuscaFichaTecnica.RecordCount > 0 Then
                                                                    VAPesoFichaTecnica = RABuscaFichaTecnica!PesoxUnidad
                                                                Else
                                                                    VAPesoFichaTecnica = 0
                                                                End If
                                                    Else
                                                        VAFichaTecnicaOrden = ""
                                                    End If
                                                    'ASIGNAMOS EL PESO A LAS VARIABLES
                                                    VAToneladasCalculoPC = ((VAProduccionPC * VAPesoFichaTecnica) / 1000)
                                                    VAToneladasPC = VAToneladasPC + VAToneladasCalculoPC
                                                    VAToneladasCalculoPNC = ((VAProduccionPNC * VAPesoFichaTecnica) / 1000)
                                                    VAToneladasPNC = VAToneladasPNC + VAToneladasCalculoPNC
                                                    VAToneladasCalculoDes = ((VAProduccionDes * VAPesoFichaTecnica) / 1000)
                                                    VAToneladasDes = VAToneladasDes + VAToneladasCalculoDes
                                                    'UNIDADES
                                                    VAUnidadesPC = VAUnidadesPC + VAProduccionPC
                                                    VAUnidadesPNC = VAUnidadesPNC + VAProduccionPNC
                                                    VAUnidadesDes = VAUnidadesDes + VAProduccionDes
                                                
                                        RABuscaDetalleProduccion.MoveNext
                                    Loop
                                Else
                                End If
                        
                        
                        
                        'EL TOTAL DE LA PRODUCCION ES LA SUMA DEL PRODUCTO CONFORME Y NO CONFORME NO INCLUYE EL DESPERDICIO
                        'TOTAL PRODUCCION
                        VTotalProduccion = VPCD + VPNCD + VPCN + VPNCN
                        'TOTAL PRODUCCION DE DIA
                        VTotalProduccionD = VPCD + VPNCD
                        'TOTAL PRODUCCION DE NOCHE
                        VTotalProduccionN = VPCN + VPNCN
                        
                        'SELECCIONA LA VELOCIDAD TEORICA DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaNoche
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaDia
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = ((VVelocidadTeoricaDia + VVelocidadTeoricaNoche) / 2)
                        ElseIf (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = 0
                        End If
                        
                        'SELECCIONA LA VELOCIDAD REAL DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadRealDia <= 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = VVelocidadRealNoche
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = VVelocidadRealDia
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = ((VVelocidadRealDia + VVelocidadRealNoche) / 2)
                        ElseIf (VVelocidadRealDia <= 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = 0
                        End If
                        
                                                                                                                                                                        
                                                                            
  'CONVIERTE LAS VARIABLES DE HORAS A MINUTOS PARA CALCULOS ________________________________________________________
                        'DIA
                            VTiempoProgramadoD = VTiempoProgramadoD * 60
                            VParosSD = VParosSD * 60
                            VParosND = VParosND * 60
                            
                        'NOCHE
                            VTiempoProgramadoN = VTiempoProgramadoN * 60
                            VParosSN = VParosSN * 60
                            VParosNN = VParosNN * 60
                                        
                                        
                'NOCHE
                'SI EL TIEMPO PROGRAMADO ES CER0 NO HACE NADA
                If VTiempoProgramadoN > 0 Then
                                         
                Else
                        VFactor1N = 0
                        VFactor2N = 0
                        VFactor3N = 0
                        VFactor4N = 0
                        VFactor5N = 0
                End If
'_______________________________________________________________________________________________________________________
                        
                        
                '% DE LINEA___________________________________________________________________________________________
                '_____________________________________________________________________________________________________
                                        
                            'FACTOR 1_________________________________________________________________________________
                                    VFactor1DN = (VTiempoProgramadoD - (VParosND + VParosSD)) + (VTiempoProgramadoN - (VParosNN + VParosSN))
                                    VFactor1DN = VFactor1DN * VVelocidadRealLinea
                                    If VFactor1DN = 0 Then
                                    Else
                                        VFactor1DN = VTotalProduccion / VFactor1DN
                                    End If
                                    
                                    'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                                    If VFactor1DN = 0 Then
                                        VFactor1DN = 1
                                    End If
                                                        
                            'FACTOR 2_________________________________________________________________________________
                                    If (VPNCD + VPNCN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor2DN = VTotalProduccion - (VPNCD + VPNCN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor2DN = VFactor2DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 3_________________________________________________________________________________
                                    If (VPDD + VPDN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor3DN = VTotalProduccion - (VPDD + VPDN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor3DN = VFactor3DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 4_________________________________________________________________________________
                                    VFactor4DN = (((VTiempoProgramadoD - VParosND) - VParosSD) + ((VTiempoProgramadoN - VParosNN) - VParosSN))
                                    If (VTiempoProgramadoD + VTiempoProgramadoN) = 0 Then
                                    ElseIf ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)) = 0 Then
                                    Else
                                        VFactor4DN = (VFactor4DN / ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)))
                                    End If
                                    
                            'FACTOR 5___________________________________________________________
                                    If VVelocidadTeoricaLinea = 0 Then
                                        VFactor5DN = 0
                                    Else
                                        VFactor5DN = VVelocidadRealLinea / VVelocidadTeoricaLinea
                                    End If
                                    
                                    'SI EL FACTOR 5 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 5
                                    If VFactor5DN = 0 Then
                                        VFactor5DN = 1
                                    End If
                                                                                
                            'EFICIENCIA DE LINEA______________________________________________________________________
                                    VPorcentajeLinea = VFactor1DN * VFactor2DN * VFactor3DN * VFactor4DN * VFactor5DN * 100
                                   
                '% DE RECHAZO__________________________________________________________________________________________
                '______________________________________________________________________________________________________
                                    VPorcentajeRechazo = VPNCD + VPNCN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeRechazo = VPorcentajeRechazo / VTotalProduccion
                                        VPorcentajeRechazo = VPorcentajeRechazo * 100
                                    End If
                        
                        
                '% DE DESPERDICIO______________________________________________________________________________________
                '______________________________________________________________________________________________________
                
                                    VPorcentajeDesperdicio = VPDD + VPDN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio / VTotalProduccion
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio * 100
                                    End If
                                    
'_______________________________________________________________________________________________________________________
'************************************************* HORAS A MINUTOS *****************************************************
'_______________________________________________________________________________________________________________________

                                                
                'CONVIERTE LAS VARIABLES A HORAS PARA IMPRIMIR LOS DATOS
                                        
                            'DIA______________________________________________________________________
                                    If VTiempoProgramadoD = 0 Then
                                        VTiempoProgramadoD = 0
                                    Else
                                        VTiempoProgramadoD = VTiempoProgramadoD / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE DIA
                                    If VParosSD = 0 Then
                                        VParosSD = 0
                                    Else
                                        VParosSD = VParosSD / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE DIA
                                    If VParosND = 0 Then
                                        VParosND = 0
                                    Else
                                        VParosND = VParosND / 60
                                    End If
                                    
                            'NOCHE______________________________________________________________________
                                    If VTiempoProgramadoN = 0 Then
                                        VTiempoProgramadoN = 0
                                    Else
                                        VTiempoProgramadoN = VTiempoProgramadoN / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE NOCHE
                                    If VParosSN = 0 Then
                                        VParosSN = 0
                                    Else
                                        VParosSN = VParosSN / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE NOCHE
                                    If VParosNN = 0 Then
                                        VParosNN = 0
                                    Else
                                        VParosNN = VParosNN / 60
                                    End If
                                    
                                        
                                    If Err > 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                        Err.Clear
                                    End If
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________

                        
                        'BUSCA LA DESCRIPCION DE LA LINEA
                        Set RBuscaDescripcionLinea = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip, UnidadMedida, TipoProducto From Lineas Where Linea = '" & VLinea & "'")
                            Else
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip, UnidadMedida, TipoProducto From Lineas Where UPPER(Linea) = '" & UCase(VLinea) & "'")
                            End If
                                
                                                VTexto = "'" & RBuscaDescripcionLinea(0) & "', "
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "#" & Format(VFechaInicial, "mm/dd/yyyy") & "#, " 'FECHA
                                                Else 'ORACLE
                                                     VTexto = VTexto & "To_Date('" & VFechaInicial & "', 'dd/mm/yyyy'), " 'FECHA
                                                End If
                                                VTexto = VTexto & VTiempoProgramadoD + VTiempoProgramadoN & ", "
                                                VTexto = VTexto & VProduccionD & ", "
                                                VTexto = VTexto & VProduccionN & ", "
                                                VTexto = VTexto & VParosSD & ", "
                                                VTexto = VTexto & VParosSN & ", "
                                                VTexto = VTexto & VProduccionD + VProduccionN & ", " 'VTiempoRealProducidoD + VTiempoRealProducidoN
                                                VTexto = VTexto & VParosND & ", "
                                                VTexto = VTexto & VParosNN & ", "
                                                VTexto = VTexto & VPCD & ", "
                                                VTexto = VTexto & VPNCD & ", "
                                                VTexto = VTexto & VPDD & ", "
                                                VTexto = VTexto & VPCN & ", "
                                                VTexto = VTexto & VPNCN & ", "
                                                VTexto = VTexto & VPDN & ", "
                                                VTexto = VTexto & VTotalProduccion & ", "
                                                VTexto = VTexto & VEficienciaRealD & ", "
                                                VTexto = VTexto & VEficienciaRealN & ", "
                                                VTexto = VTexto & VPorcentajeLinea & ", "
                                                VTexto = VTexto & VPorcentajeRechazo & ", "
                                                VTexto = VTexto & Format(VPorcentajeDesperdicio, "#,###,##0.00") & ", "
                                                VTexto = VTexto & VFactor1DN & ", "
                                                VTexto = VTexto & VFactor5DN & ", "
                                                VTexto = VTexto & VVelocidadTeoricaDia & ", "
                                                VTexto = VTexto & VVelocidadRealDia & ", "
                                                VTexto = VTexto & VVelocidadTeoricaNoche & ", "
                                                VTexto = VTexto & VVelocidadRealNoche & ", '"
                                                VTexto = VTexto & RBuscaDescripcionLinea(1) & "', "
                                                VTexto = VTexto & VAToneladasPC & ", "
                                                VTexto = VTexto & VAToneladasPNC & ", "
                                                VTexto = VTexto & Format(VAToneladasDes, "#,###,##0.00") & ", "
                                                VTexto = VTexto & VAUnidadesPC & ", "
                                                VTexto = VTexto & VAUnidadesPNC & ", "
                                                VTexto = VTexto & VAUnidadesDes & ", '"
                                                VTexto = VTexto & RBuscaDescripcionLinea(2) & "', "
                                                VTexto = VTexto & (VParosCFD + VParosCFN) & ", "
                                                VTexto = VTexto & (VParosMPD + VParosMPN)
                            
                                                Conexion.Execute "Insert Into ReporteEjecutivoAcumulado Values(" & VTexto & ")"
                           
                                    
                                    If Err > 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                            Err.Clear
                                            
                                    End If
                                    
                            VFechaInicial = VFechaInicial + 1
                                    
                    Else ' IF DE CONDICION DE LA QUE NO ENCUENTRA DATOS EN RANGO DE FECHAS
                            
                            'A LA FECHA ACTUAL LE SUMA 1 PARA INCREMENTAR LA FECHA
                            'LA EFICIENCIA SE SACA POR DIA
                            VFechaInicial = VFechaInicial + 1
                    End If
                
                        
                Loop
                                
                
                'SIGUE AL SIGUIENTE REGISTRO DE LINEA
                RSeleccionaLineas.MoveNext
  Loop
  
                
  
  
                'SACA EL OTRO REPORTE EJECUTIVO
                ReporteEjecutivo2
                
                
                
                
                'BUSCAMOS LAS ------------------------------ INVENTARIO ----------------------------------------
                Set RInventario = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RInventario, "Select * from Inventario Where Fecha = #" & Format(DTPInventario.Value, "mm/dd/yyyy") & "#")
                    Else
                        Call Abrir_Recordset(RInventario, "Select * from Inventario Where Fecha = To_Date('" & DTPInventario.Value & "', 'dd/mm/yyyy')")
                    End If
                    If RInventario.RecordCount > 0 Then
                            Do Until RInventario.EOF
                                    
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        VTexto = "#" & Format(RInventario(0), "mm/dd/yyyy") & "#, '" 'FECHA
                                    Else 'ORACLE
                                        VTexto = "To_Date('" & Format(RInventario(0), "dd/mm/yyyy") & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                                    End If
                                    VTexto = VTexto & RInventario(1) & "', '" 'FICHA
                                    VTexto = VTexto & RInventario(2) & "', " 'BODEGA
                                    VTexto = VTexto & RInventario(3) & ", '" 'CANTIDAD
                                    VTexto = VTexto & GUsuario & "'" 'USUARIO
                                    
                                    Conexion.Execute "Insert Into ReporteEjecutivoInventario Values(" & VTexto & ")"
                                    
                                    
                                    If Err <> 0 Then
                                            MsgBox Err.Number & " " & Err.Description
                                            Err.Clear
                                            
                                    End If
                                RInventario.MoveNext
                            Loop
                    End If
                    
                    
                    
                'BUSCAMOS LAS ------------------------------ VENTAS ----------------------------------------
                
                VTexto = ""
                
                Set RClientes = New ADODB.Recordset
                    Call Abrir_Recordset(RClientes, "Select CodigoCliente From Clientes Where Estado = 'ACTIVO'")
                    
                    'CONTAMOS CUANTOS DIAS SON DE VENTAS
                    Set RBuscaDiasVenta = New ADODB.Recordset
                        Call Abrir_Recordset(RBuscaDiasVenta, "Select Fecha from Ventas Where Fecha >= #" & Month(DTPVentas.Value) & "/" & "01/" & Year(DTPVentas.Value) & "# And Fecha <= #" & Format(DTPVentas.Value, "mm/dd/yyyy") & "# group by fecha")
                                If RBuscaDiasVenta.RecordCount > 0 Then
                                    VDiasVenta = RBuscaDiasVenta.RecordCount 'LOS DIAS DE VENTA VENDIDOS
                                Else
                                    VDiasVenta = "0"
                                End If
                                
                    'LA META A BUSCAR
                    VFechaMeta = CDate("01/" & Month(DTPVentas.Value) & "/" & Year(DTPVentas.Value))
                                
                    
                    Do Until RClientes.EOF
                    
                            Set RFichaTecnica = New ADODB.Recordset
                                Call Abrir_Recordset(RFichaTecnica, "Select Esp_Tec, TipoVenta From FichaTecnica Where Activa = -1")
                                                                                                   
                                    Do Until RFichaTecnica.EOF
                
                                                    Set RVentas = New ADODB.Recordset
                                                         Call Abrir_Recordset(RVentas, "Select * from Ventas Where Fecha = #" & Format(DTPVentas.Value, "mm/dd/yyyy") & "# And FichaTecnica = '" & RFichaTecnica!Esp_Tec & "' and Cliente = '" & RClientes!CodigoCliente & "'")
                                                                                                                
                                                        If RVentas.RecordCount > 0 Then
                                                                
                                                                Do Until RVentas.EOF
                                                               
                                                                    Set RBuscaMetaMensual = New ADODB.Recordset
                                                                        Call Abrir_Recordset(RBuscaMetaMensual, "Select MetaDolares, MetaCantidad, MetaToneladas From VentasMetasMensuales where TipoFichaTecnica = '" & RFichaTecnica!TipoVenta & "' And Cliente = '" & RClientes!CodigoCliente & "' And Fecha = #" & Format(VFechaMeta, "mm/dd/yyyy") & "#")
                                                                              If RBuscaMetaMensual.RecordCount > 0 Then
                                                                                 VMetaDolares = RBuscaMetaMensual!MetaDolares
                                                                                 VMetaCantidad = RBuscaMetaMensual!MetaCantidad
                                                                                 VMetaToneladas = RBuscaMetaMensual!MetaToneladas
                                                                              Else
                                                                                 VMetaDolares = 0
                                                                                 VMetaCantidad = 0
                                                                                 VMetaToneladas = 0
                                                                              End If
                                                                                                                                                                   
                                                                                    
                                                                                    'VENTAS DISEÑO NUEVO PARA EL REPORTE EJECUTIVO
                                                                                    '______________________________________________________________________________________________
                                                                                    
                                                                                    VTexto = "'" & RVentas!FichaTecnica & "', '" 'FICHA
                                                                                    VTexto = VTexto & RVentas!Cliente & "', " 'CLIENTE
                                                                                    If RVentas!Bodega = "T21" Then
                                                                                        VTexto = VTexto & RVentas!Cantidad & ", " 'UNIDADES DIARIAS
                                                                                    Else
                                                                                        VTexto = VTexto & "0" & ", "
                                                                                    End If
                                                                                    If RVentas!Bodega = "T22" Then
                                                                                        VTexto = VTexto & RVentas!Cantidad & ", " 'PESOS DIARIOS
                                                                                    Else
                                                                                        VTexto = VTexto & "0" & ", "
                                                                                    End If
                                                                                    If RVentas!Bodega = "T23" Then
                                                                                        VTexto = VTexto & RVentas!Cantidad & ", " 'DOLARES DIARIOS
                                                                                    Else
                                                                                        VTexto = VTexto & "0" & ", "
                                                                                    End If
                                                                                    If RVentas!Bodega = "T24" Then
                                                                                        VTexto = VTexto & RVentas!Cantidad & ", " 'UNIDADES ACUMULADAS
                                                                                    Else
                                                                                        VTexto = VTexto & "0" & ", "
                                                                                    End If
                                                                                    
                                                                                    VTexto = VTexto & VMetaCantidad & ", "
                                                                                    VTexto = VTexto & VDiasVenta & ", "
                                                                                    
                                                                                    If RVentas!Bodega = "T25" Then
                                                                                        VTexto = VTexto & RVentas!Cantidad & ", " 'PESOS ACUMULADOS
                                                                                    Else
                                                                                        VTexto = VTexto & "0" & ", "
                                                                                    End If
                                                                                    
                                                                                    If RVentas!Bodega = "T26" Then
                                                                                        VTexto = VTexto & RVentas!Cantidad & ", " 'DOLARES ACUMULADOS
                                                                                    Else
                                                                                        VTexto = VTexto & "0" & ", "
                                                                                    End If
                                                                                    
                                                                                    VTexto = VTexto & VMetaDolares & ", "
                                                                                    
                                                                                    If RVentas!Bodega = "T30" Then
                                                                                        VTexto = VTexto & RVentas!Cantidad & ", " 'TONELADAS VENTAS
                                                                                    Else
                                                                                        VTexto = VTexto & "0" & ", "
                                                                                    End If
                                                                                    
                                                                                    If RVentas!Bodega = "T31" Then
                                                                                        VTexto = VTexto & RVentas!Cantidad & ", " 'TONELADAS ACUMULADAS
                                                                                    Else
                                                                                        VTexto = VTexto & "0" & ", "
                                                                                    End If
                                                                                    
                                                                                    'SE AGREGAN CAMBIOS PARA REP EJEC TEMA: METAS TONELADAS
                                                                                    VTexto = VTexto & VMetaToneladas & " "
                                                                                    'MetaToneladas
                                                                                    
                                                                                    Conexion.Execute "Insert Into ReporteEjecutivoVentasNUEVAS Values(" & VTexto & ")"
                                                                                    
                                                                                    If Err <> 0 Then
                                                                                            MsgBox "Error en agregar VENTAS NUEVAS " & Err.Number & " " & Err.Description
                                                                                            Err.Clear
                                                                                    End If
                                                                                    
                                                                        RVentas.MoveNext
                                                                    Loop
                                                        Else ' SI NO HAY VENTAS DEL CLIENTE Y FICHA TECNICA BUSCA LAS METAS
                                                        
                                                                   'VENTAS DISEÑO NUEVO PARA EL REPORTE EJECUTIVO
                                                                                    '______________________________________________________________________________________________
                                                                                    Set RBuscaMetaMensual = New ADODB.Recordset
                                                                                            Call Abrir_Recordset(RBuscaMetaMensual, "Select MetaDolares, MetaCantidad, MetaToneladas From VentasMetasMensuales where TipoFichaTecnica = '" & RFichaTecnica!TipoVenta & "' And Cliente = '" & RClientes!CodigoCliente & "' And Fecha = #" & Format(VFechaMeta, "mm/dd/yyyy") & "#")
                                                                                                 If RBuscaMetaMensual.RecordCount > 0 Then
                                                                                                    VMetaDolares = RBuscaMetaMensual!MetaDolares
                                                                                                    VMetaCantidad = RBuscaMetaMensual!MetaCantidad
                                                                                                    VMetaToneladas = RBuscaMetaMensual!MetaToneladas
                                                                                                 Else
                                                                                                    VMetaDolares = 0
                                                                                                    VMetaCantidad = 0
                                                                                                    VMetaToneladas = 0
                                                                                                 End If
                                                                              
                                                                                    
                                                                                    VTexto = "'" & RFichaTecnica!Esp_Tec & "', '" 'FICHA
                                                                                    VTexto = VTexto & RClientes!CodigoCliente & "', " 'CLIENTE
                                                                                    VTexto = VTexto & "0" & ", " 'UNIDADES DIARIAS
                                                                                    VTexto = VTexto & "0" & ", " 'PESOS DIARIOS
                                                                                    VTexto = VTexto & "0" & ", " 'DOLARES DIARIOS
                                                                                    VTexto = VTexto & "0" & ", " 'UNIDADES ACUMULADAS
                                                                                    VTexto = VTexto & VMetaCantidad & ", "
                                                                                    VTexto = VTexto & VDiasVenta & ", "
                                                                                    VTexto = VTexto & "0" & ", " 'PESOS ACUMULADOS
                                                                                    VTexto = VTexto & "0" & ", " 'DOLARES ACUMULADOS
                                                                                    VTexto = VTexto & VMetaDolares & ", "
                                                                                    VTexto = VTexto & "0" & ", " 'TONELADAS VENTAS
                                                                                    VTexto = VTexto & "0" & ", " 'TONELADAS ACUMULADAS
                                                                                    VTexto = VTexto & VMetaToneladas & "" 'META TONELADAS
                                                                                                                                                                                             
                                                                                    'SOLO SI LA META DOLARES O LA DE CANTIDAD ES MAYOR QUE CERO, AGREGA LOS DATOS DE METAS
                                                                                    If (VMetaDolares > 0 Or VMetaCantidad > 0) Then
                                                                                        Conexion.Execute "Insert Into ReporteEjecutivoVentasNUEVAS Values(" & VTexto & ")"
                                                                                    End If
                                                                                    
                                                                                    If Err <> 0 Then
                                                                                            MsgBox "Error en agregar VENTAS NUEVAS " & Err.Number & " " & Err.Description
                                                                                            Err.Clear
                                                                                    End If
                                                        End If
                                                        
                                           RFichaTecnica.MoveNext
                                   Loop 'FICHA TECNICAS
                                          
                        RClientes.MoveNext
                    Loop 'CLIENTES
                        
  ' 11-08-2011 incluir grafica de produccion por dia
                        Set RSeleccionaLineas = New ADODB.Recordset
                            Call Abrir_Recordset(RSeleccionaLineas, "Select C.Fecha, L.Planta, Sum(C.ProductoConforme) From Lineas L, EncabezadoCapturaParos C Where C.Fecha >= #" & Format(DTPFecEfiAcu.Value, "MM/dd/yyyy") & "# And C.Fecha <= #" & Format(DTPFecEfiFinAcu.Value, "MM/dd/yyyy") & "# And C.Linea = L.Linea And L.IncluyeEnGraficaReporteEjecutivo = 'SI' Group By C.Fecha, L.Planta")
                            If RSeleccionaLineas.RecordCount > 0 Then
                                Do Until RSeleccionaLineas.EOF
                                                VTexto = Day(RSeleccionaLineas!fecha) & ", '"
                                                VTexto = VTexto & RSeleccionaLineas!Planta & "', "
                                                VTexto = VTexto & RSeleccionaLineas(2)
                                                
                                                Conexion.Execute "Insert Into ReporteEjecutivoProduccionPorDia Values(" & VTexto & ")"
                                                
                                                 If Err <> 0 Then
                                                    MsgBox "Error en agregar GRAFICA PRODUCCION POR DIA " & Err.Number & " " & Err.Description
                                                    Err.Clear
                                                End If
                                    RSeleccionaLineas.MoveNext
                                Loop
                            End If
                            
'30-08-2011 incluir grafica de produccion por dia y productos seleccionados, esto se hace con un campo en la ficha tecnica que es llama numerografica
                        Set RSeleccionaLineas = New ADODB.Recordset
                            Call Abrir_Recordset(RSeleccionaLineas, "Select E.Fecha, L.Planta, Sum(D.ProductoConforme) From Lineas L, EncabezadoCapturaParos E, DetalleProduccionPorOrden D, FichaTecnica F, EncabezadoOrdenProduccion O Where E.Fecha >= #" & Format(DTPFecEfiAcu.Value, "MM/dd/yyyy") & "# And E.Fecha <= #" & Format(DTPFecEfiFinAcu.Value, "MM/dd/yyyy") & "# And E.Linea = L.Linea And L.IncluyeEnGraficaReporteEjecutivo = 'SI' And E.Documento = D.Documento And D.Orden = O.Documento And O.FichaTecnica = F.Esp_Tec And F.NumeroGrafica = '1' Group By E.Fecha, L.Planta")
                            If RSeleccionaLineas.RecordCount > 0 Then
                                Do Until RSeleccionaLineas.EOF
                                                VTexto = Day(RSeleccionaLineas!fecha) & ", '"
                                                VTexto = VTexto & RSeleccionaLineas!Planta & "', "
                                                VTexto = VTexto & RSeleccionaLineas(2)
                                                
                                                Conexion.Execute "Insert Into ReporteEjecutivoProduccionPorDiaGrafica1 Values(" & VTexto & ")"
                                                
                                                 If Err <> 0 Then
                                                    MsgBox "Error en agregar GRAFICA PRODUCCION POR DIA 1 " & Err.Number & " " & Err.Description
                                                    Err.Clear
                                                End If
                                    RSeleccionaLineas.MoveNext
                                Loop
                            End If
                            
                        Set RSeleccionaLineas = New ADODB.Recordset
                            Call Abrir_Recordset(RSeleccionaLineas, "Select E.Fecha, L.Planta, Sum(D.ProductoConforme) From Lineas L, EncabezadoCapturaParos E, DetalleProduccionPorOrden D, FichaTecnica F, EncabezadoOrdenProduccion O Where E.Fecha >= #" & Format(DTPFecEfiAcu.Value, "MM/dd/yyyy") & "# And E.Fecha <= #" & Format(DTPFecEfiFinAcu.Value, "MM/dd/yyyy") & "# And E.Linea = L.Linea And L.IncluyeEnGraficaReporteEjecutivo = 'SI' And E.Documento = D.Documento And D.Orden = O.Documento And O.FichaTecnica = F.Esp_Tec And F.NumeroGrafica = '2' Group By E.Fecha, L.Planta")
                            If RSeleccionaLineas.RecordCount > 0 Then
                                Do Until RSeleccionaLineas.EOF
                                                VTexto = Day(RSeleccionaLineas!fecha) & ", '"
                                                VTexto = VTexto & RSeleccionaLineas!Planta & "', "
                                                VTexto = VTexto & RSeleccionaLineas(2)
                                                
                                                Conexion.Execute "Insert Into ReporteEjecutivoProduccionPorDiaGrafica2 Values(" & VTexto & ")"
                                                
                                                 If Err <> 0 Then
                                                    MsgBox "Error en agregar GRAFICA PRODUCCION POR DIA 2 " & Err.Number & " " & Err.Description
                                                    Err.Clear
                                                End If
                                    RSeleccionaLineas.MoveNext
                                Loop
                            End If
                            
                        Set RSeleccionaLineas = New ADODB.Recordset
                            Call Abrir_Recordset(RSeleccionaLineas, "Select E.Fecha, L.Planta, Sum(D.ProductoConforme) From Lineas L, EncabezadoCapturaParos E, DetalleProduccionPorOrden D, FichaTecnica F, EncabezadoOrdenProduccion O Where E.Fecha >= #" & Format(DTPFecEfiAcu.Value, "MM/dd/yyyy") & "# And E.Fecha <= #" & Format(DTPFecEfiFinAcu.Value, "MM/dd/yyyy") & "# And E.Linea = L.Linea And L.IncluyeEnGraficaReporteEjecutivo = 'SI' And E.Documento = D.Documento And D.Orden = O.Documento And O.FichaTecnica = F.Esp_Tec And F.NumeroGrafica = '3' Group By E.Fecha, L.Planta")
                            If RSeleccionaLineas.RecordCount > 0 Then
                                Do Until RSeleccionaLineas.EOF
                                                VTexto = Day(RSeleccionaLineas!fecha) & ", '"
                                                VTexto = VTexto & RSeleccionaLineas!Planta & "', "
                                                VTexto = VTexto & RSeleccionaLineas(2)
                                                
                                                Conexion.Execute "Insert Into ReporteEjecutivoProduccionPorDiaGrafica3 Values(" & VTexto & ")"
                                                
                                                 If Err <> 0 Then
                                                    MsgBox "Error en agregar GRAFICA PRODUCCION POR DIA 3 " & Err.Number & " " & Err.Description
                                                    Err.Clear
                                                End If
                                    RSeleccionaLineas.MoveNext
                                Loop
                            End If
                            
                            Set RSeleccionaLineas = New ADODB.Recordset
                            Call Abrir_Recordset(RSeleccionaLineas, "Select E.Fecha, L.Planta, Sum(D.ProductoConforme) From Lineas L, EncabezadoCapturaParos E, DetalleProduccionPorOrden D, FichaTecnica F, EncabezadoOrdenProduccion O Where E.Fecha >= #" & Format(DTPFecEfiAcu.Value, "MM/dd/yyyy") & "# And E.Fecha <= #" & Format(DTPFecEfiFinAcu.Value, "MM/dd/yyyy") & "# And E.Linea = L.Linea And L.IncluyeEnGraficaReporteEjecutivo = 'SI' And E.Documento = D.Documento And D.Orden = O.Documento And O.FichaTecnica = F.Esp_Tec And F.NumeroGrafica = '4' Group By E.Fecha, L.Planta")
                            If RSeleccionaLineas.RecordCount > 0 Then
                                Do Until RSeleccionaLineas.EOF
                                                VTexto = Day(RSeleccionaLineas!fecha) & ", '"
                                                VTexto = VTexto & RSeleccionaLineas!Planta & "', "
                                                VTexto = VTexto & RSeleccionaLineas(2)
                                                
                                                Conexion.Execute "Insert Into ReporteEjecutivoProduccionPorDiaGrafica4 Values(" & VTexto & ")"
                                                
                                                 If Err <> 0 Then
                                                    MsgBox "Error en agregar GRAFICA PRODUCCION POR DIA 4 " & Err.Number & " " & Err.Description
                                                    Err.Clear
                                                End If
                                    RSeleccionaLineas.MoveNext
                                Loop
                            End If
                            
                            Set RSeleccionaLineas = New ADODB.Recordset
                            Call Abrir_Recordset(RSeleccionaLineas, "Select E.Fecha, L.Planta, Sum(D.ProductoConforme) From Lineas L, EncabezadoCapturaParos E, DetalleProduccionPorOrden D, FichaTecnica F, EncabezadoOrdenProduccion O Where E.Fecha >= #" & Format(DTPFecEfiAcu.Value, "MM/dd/yyyy") & "# And E.Fecha <= #" & Format(DTPFecEfiFinAcu.Value, "MM/dd/yyyy") & "# And E.Linea = L.Linea And L.IncluyeEnGraficaReporteEjecutivo = 'SI' And E.Documento = D.Documento And D.Orden = O.Documento And O.FichaTecnica = F.Esp_Tec And F.NumeroGrafica = '5' Group By E.Fecha, L.Planta")
                            If RSeleccionaLineas.RecordCount > 0 Then
                                Do Until RSeleccionaLineas.EOF
                                                VTexto = Day(RSeleccionaLineas!fecha) & ", '"
                                                VTexto = VTexto & RSeleccionaLineas!Planta & "', "
                                                VTexto = VTexto & RSeleccionaLineas(2)
                                                
                                                Conexion.Execute "Insert Into ReporteEjecutivoProduccionPorDiaGrafica5 Values(" & VTexto & ")"
                                                
                                                 If Err <> 0 Then
                                                    MsgBox "Error en agregar GRAFICA PRODUCCION POR DIA 5 " & Err.Number & " " & Err.Description
                                                    Err.Clear
                                                End If
                                    RSeleccionaLineas.MoveNext
                                Loop
                            End If
  
                'TIPO DE REPORTE
                If OptRepEje.Value = True Then
                        
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     GNombreReporte = "ReporteE.rpt"
                                Else
                                     GNombreReporte = "ReporteEO.rpt"
                                End If
                        GTituloReporte = "Produccion, Inventarios y Ventas Al: " & DTPFecEfiFinAcu.Value
                End If

                

End Sub

Public Sub ReporteEjecutivo2()
On Error Resume Next
                            
                                                        

            
            'SELECCIONA LAS LINEAS DE ACUERDO A LA OPCION
            If OptEfi.Item(0).Value = True = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Linea = '" & TxtEfiGru.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Linea) = '" & UCase(TxtEfiGru.Text) & "'")
                            End If
            ElseIf OptEfi.Item(1).Value = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Grupo = '" & TxtEfiGru.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Grupo) = '" & UCase(TxtEfiGru.Text) & "'")
                            End If
            End If
             
             If RSeleccionaLineas.RecordCount > 0 Then
             Else
                    'MsgBox "Linea O Lineas No Existen ", vbOKOnly + vbInformation, "Informacion"
                    'Exit Sub
             End If
  
  
  
  'CREA UN CICLO CON LAS LINEAS POSIBLES DE ACUERDO A LA OPCION ELEGIDA
  Do Until RSeleccionaLineas.EOF
                        
                    'ASIGNA LA LINEA QUE ES SELECCIONADA
                    VLinea = RSeleccionaLineas!Linea
                            
                    'FECHA DE INICIO DEL RANGO
                    VFechaInicial = Format(DTPFecEfi.Value, "dd/mm/yyyy")
                    'FECHA DEL FINAL DEL RANGO
                    VFechaFinal = Format(DTPFecEfiFin.Value, "dd/mm/yyyy")
                        
                
                Do Until VFechaInicial > VFechaFinal
                
                            VToneladasPC = 0
                            VToneladasPNC = 0
                            VToneladasDes = 0
                            VProduccionPC = 0
                            VProduccionPNC = 0
                            VProduccionDes = 0
                            VUnidadesPC = 0
                            VUnidadesPNC = 0
                            VUnidadesDes = 0
                        
                        
'VERIFICA SI HAY DATOS EN LA PRESENTE FECHA Y SI NO HAY PASA A LA SIGUIENTE FECHA
'ESTO NOS SIRVE PARA CUANDO SAQUEMOS EL REPORTE DE EFICIENCIA NO TOME EN CUENTA LOS DIAS QUE
'NO SE TRABAJO PORQUE AFECTA LA EFICIENCIA DE LINEA Y PLANTA
                Set RCapturaParos = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Linea = '" & VLinea & "'")
                    Else
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                    End If
                    If RCapturaParos.RecordCount > 0 Then
                        'NO HACE NADA SI HAY DATOS ESTA BIEN
                    
                                
                  
                  '*******  TIEMPO PROGRAMADO Y VELOCIDAD **************************************************
  
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE DIA
                         Set RTiempoProgramadoD = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoD.RecordCount > 0 Then
                                            VTiempoProgramadoD = RTiempoProgramadoD!HorasProgramadas
                                            VVelocidadTeoricaDia = RTiempoProgramadoD!VelocidadTeorica
                                            VVelocidadRealDia = RTiempoProgramadoD!VelocidadReal
                                            VGrupoDia = RTiempoProgramadoD!Grupo
                                            VParosND = RTiempoProgramadoD!ParoN / 60
                                            VParosSD = RTiempoProgramadoD!Paros / 60
                                            VProduccionD = RTiempoProgramadoD!ParoP / 60
                                            VPCD = RTiempoProgramadoD!ProductoConforme
                                            VPNCD = RTiempoProgramadoD!ProductoNoConforme
                                            VPDD = RTiempoProgramadoD!Desperdicio
                                            VEficienciaRealD = RTiempoProgramadoD!Eficiencia
                                            VParosCFD = RTiempoProgramadoD!ParoCF
                                            VParosMPD = RTiempoProgramadoD!ParoMP
                                Else
                                            VTiempoProgramadoD = 0
                                            VVelocidadTeoricaDia = 0
                                            VVelocidadRealDia = 0
                                            VGrupoDia = ""
                                            VParosND = 0
                                            VParosSD = 0
                                            VProduccionD = 0
                                            VPCD = 0
                                            VPNCD = 0
                                            VPDD = 0
                                            VEficienciaRealD = 0
                                            VParosCFD = 0
                                            VParosMPD = 0
                                            VParosCFD = 0
                                            VParosMPD = 0
                             End If
                                               
                        
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE NOCHE
                         Set RTiempoProgramadoN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoN.RecordCount > 0 Then
                                            VTiempoProgramadoN = RTiempoProgramadoN!HorasProgramadas
                                            VVelocidadTeoricaNoche = RTiempoProgramadoN!VelocidadTeorica
                                            VVelocidadRealNoche = RTiempoProgramadoN!VelocidadReal
                                            VGrupoNoche = RTiempoProgramadoN!Grupo
                                            VParosNN = RTiempoProgramadoN!ParoN / 60
                                            VParosSN = RTiempoProgramadoN!Paros / 60
                                            VProduccionN = RTiempoProgramadoN!ParoP / 60
                                            VPCN = RTiempoProgramadoN!ProductoConforme
                                            VPNCN = RTiempoProgramadoN!ProductoNoConforme
                                            VPDN = RTiempoProgramadoN!Desperdicio
                                            VEficienciaRealN = RTiempoProgramadoN!Eficiencia
                             Else
                                            VTiempoProgramadoN = 0
                                            VVelocidadTeoricaNoche = 0
                                            VVelocidadRealNoche = 0
                                            VGrupoNoche = ""
                                            VParosNN = 0
                                            VParosSN = 0
                                            VProduccionN = 0
                                            VPCN = 0
                                            VPNCN = 0
                                            VPDN = 0
                                            VEficienciaRealN = 0
                             End If
                                                
                        
                        
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE DIA
                        VTiempoRealProducidoD = VTiempoProgramadoD - VParosND
                                                    
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE NOCHE
                        If VTiempoProgramadoN = 0 Then
                            VTiempoRealProducidoN = 0
                        Else
                            VTiempoRealProducidoN = VTiempoProgramadoN - VParosNN
                        End If
                                                    
                        'HORAS PRODUCIDAS POR LOS 2 TURNOS
                            VHorasProducidasDN = Format(VTiempoRealProducidoD + VTiempoRealProducidoN, "#,###,##0.00")
                                                    
                        'TOTAL DE PAROS S "NO AFECTAN"
                            VParosDN = Format(VParosND + VParosNN, "#,###,##0.00")
                        
'________________________________________________________________________________________________________________________
         'PARA EL REPORTE EJECUTIVO
         'BUSCAMOS EL DETALLE DE LA PRODUCCION EN EL RANGO DE FECHAS INDICADO
                            Set RBuscaDetalleProduccion = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaDetalleProduccion, "Select DP.Orden, DP.ProductoConforme, DP.ProductoNoConforme, DP.Desperdicio From DetalleProduccionPorOrden DP, EncabezadoCapturaParos EP Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Linea = '" & VLinea & "' And EP.Documento = DP.Documento")
                                Else
                                    Call Abrir_Recordset(RBuscaDetalleProduccion, "Select DP.Orden, DP.ProductoConforme, DP.ProductoNoConforme, DP.Desperdicio From DetalleProduccionPorOrden DP, EncabezadoCapturaParos EP Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = '" & UCase(VLinea) & "' And EP.Documento = DP.Documento")
                                End If
                                If RBuscaDetalleProduccion.RecordCount > 0 Then
                                    Do Until RBuscaDetalleProduccion.EOF
                                                'ASIGNAMOS VARIABLES
                                                VOrdenDetalle = RBuscaDetalleProduccion(0)
                                                VProduccionPC = RBuscaDetalleProduccion(1)
                                                VProduccionPNC = RBuscaDetalleProduccion(2)
                                                VProduccionDes = RBuscaDetalleProduccion(3)
                                                
                                                'AHORA BUSCAMOS LA ORDEN
                                                Set RBuscaOrden = New ADODB.Recordset
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Call Abrir_Recordset(RBuscaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & VOrdenDetalle & "'")
                                                        
                                                    Else
                                                        Call Abrir_Recordset(RBuscaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(VOrdenDetalle) & "'")
                                                    End If
                                                    'ASIGNAMOS LA FICHA TECNICA QUE USA LA ORDEN
                                                    If RBuscaOrden.RecordCount > 0 Then
                                                        VFichaTecnicaOrden = RBuscaOrden!FichaTecnica
                                                            'BUSCAMOS LA FICHA TECNICA PARA OBTENER EL PESO
                                                            Set RBuscaFichaTecnica = New ADODB.Recordset
                                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select PesoxUnidad From FichaTecnica Where Esp_Tec = '" & VFichaTecnicaOrden & "'")
                                                                Else
                                                                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select PesoxUnidad From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(VFichaTecnicaOrden) & "'")
                                                                End If
                                                                'ASIGNAMOS EL PESO
                                                                If RBuscaFichaTecnica.RecordCount > 0 Then
                                                                    VPesoFichaTecnica = RBuscaFichaTecnica!PesoxUnidad
                                                                Else
                                                                    VPesoFichaTecnica = 0
                                                                End If
                                                    Else
                                                        VFichaTecnicaOrden = ""
                                                    End If
                                                    'ASIGNAMOS EL PESO A LAS VARIABLES
                                                    VToneladasCalculoPC = ((VProduccionPC * VPesoFichaTecnica) / 1000)
                                                    VToneladasPC = VToneladasPC + VToneladasCalculoPC
                                                    VToneladasCalculoPNC = ((VProduccionPNC * VPesoFichaTecnica) / 1000)
                                                    VToneladasPNC = VToneladasPNC + VToneladasCalculoPNC
                                                    VToneladasCalculoDes = ((VProduccionDes * VPesoFichaTecnica) / 1000)
                                                    VToneladasDes = VToneladasDes + VToneladasCalculoDes
                                                    
                                                    'UNIDADES
                                                    VUnidadesPC = VUnidadesPC + VProduccionPC
                                                    VUnidadesPNC = VUnidadesPNC + VProduccionPNC
                                                    VUnidadesDes = VUnidadesDes + VProduccionDes
                                        
                                        RBuscaDetalleProduccion.MoveNext
                                    Loop
                                Else
                                End If
                                                            
'________________________________________________________________________________________________________________________
'________________________________________________________________________________________________________________________
                        
                        
                        
                        'EL TOTAL DE LA PRODUCCION ES LA SUMA DEL PRODUCTO CONFORME Y NO CONFORME NO INCLUYE EL DESPERDICIO
                        'TOTAL PRODUCCION
                        VTotalProduccion = VPCD + VPNCD + VPCN + VPNCN
                        'TOTAL PRODUCCION DE DIA
                        VTotalProduccionD = VPCD + VPNCD
                        'TOTAL PRODUCCION DE NOCHE
                        VTotalProduccionN = VPCN + VPNCN
                                                
                        'SELECCIONA LA VELOCIDAD TEORICA DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaNoche
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaDia
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = ((VVelocidadTeoricaDia + VVelocidadTeoricaNoche) / 2)
                        ElseIf (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = 0
                        End If
                        
                        'SELECCIONA LA VELOCIDAD REAL DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadRealDia <= 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = VVelocidadRealNoche
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = VVelocidadRealDia
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = ((VVelocidadRealDia + VVelocidadRealNoche) / 2)
                        ElseIf (VVelocidadRealDia <= 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = 0
                        End If
                        
                                                                                                                                                                        
                                                                            
  'CONVIERTE LAS VARIABLES DE HORAS A MINUTOS PARA CALCULOS ________________________________________________________
                        'DIA
                            VTiempoProgramadoD = VTiempoProgramadoD * 60
                            VParosSD = VParosSD * 60
                            VParosND = VParosND * 60
                            
                        'NOCHE
                            VTiempoProgramadoN = VTiempoProgramadoN * 60
                            VParosSN = VParosSN * 60
                            VParosNN = VParosNN * 60
                                        
                                        
                'NOCHE
                'SI EL TIEMPO PROGRAMADO ES CER0 NO HACE NADA
                If VTiempoProgramadoN > 0 Then
                                         
                Else
                        VFactor1N = 0
                        VFactor2N = 0
                        VFactor3N = 0
                        VFactor4N = 0
                        VFactor5N = 0
                End If
'_______________________________________________________________________________________________________________________
                        
                        
                '% DE LINEA___________________________________________________________________________________________
                '_____________________________________________________________________________________________________
                                        
                            'FACTOR 1_________________________________________________________________________________
                                    VFactor1DN = (VTiempoProgramadoD - (VParosND + VParosSD)) + (VTiempoProgramadoN - (VParosNN + VParosSN))
                                    VFactor1DN = VFactor1DN * VVelocidadRealLinea
                                    If VFactor1DN = 0 Then
                                    Else
                                        VFactor1DN = VTotalProduccion / VFactor1DN
                                    End If
                                    
                                    'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                                    If VFactor1DN = 0 Then
                                        VFactor1DN = 1
                                    End If
                                                        
                            'FACTOR 2_________________________________________________________________________________
                                    If (VPNCD + VPNCN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor2DN = VTotalProduccion - (VPNCD + VPNCN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor2DN = VFactor2DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 3_________________________________________________________________________________
                                    If (VPDD + VPDN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor3DN = VTotalProduccion - (VPDD + VPDN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor3DN = VFactor3DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 4_________________________________________________________________________________
                                    VFactor4DN = (((VTiempoProgramadoD - VParosND) - VParosSD) + ((VTiempoProgramadoN - VParosNN) - VParosSN))
                                    If (VTiempoProgramadoD + VTiempoProgramadoN) = 0 Then
                                    ElseIf ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)) = 0 Then
                                    Else
                                        VFactor4DN = (VFactor4DN / ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)))
                                    End If
                                    
                            'FACTOR 5___________________________________________________________
                                    If VVelocidadTeoricaLinea = 0 Then
                                        VFactor5DN = 0
                                    Else
                                        VFactor5DN = VVelocidadRealLinea / VVelocidadTeoricaLinea
                                    End If
                                    
                                    'SI EL FACTOR 5 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 5
                                    If VFactor5DN = 0 Then
                                        VFactor5DN = 1
                                    End If
                                                                                
                            'EFICIENCIA DE LINEA______________________________________________________________________
                                    VPorcentajeLinea = VFactor1DN * VFactor2DN * VFactor3DN * VFactor4DN * VFactor5DN * 100
                                   
                '% DE RECHAZO__________________________________________________________________________________________
                '______________________________________________________________________________________________________
                                    VPorcentajeRechazo = VPNCD + VPNCN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeRechazo = VPorcentajeRechazo / VTotalProduccion
                                        VPorcentajeRechazo = VPorcentajeRechazo * 100
                                    End If
                        
                        
                '% DE DESPERDICIO______________________________________________________________________________________
                '______________________________________________________________________________________________________
                
                                    VPorcentajeDesperdicio = VPDD + VPDN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio / VTotalProduccion
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio * 100
                                    End If
                                    
'_______________________________________________________________________________________________________________________
'************************************************* HORAS A MINUTOS *****************************************************
'_______________________________________________________________________________________________________________________

                                                
                'CONVIERTE LAS VARIABLES A HORAS PARA IMPRIMIR LOS DATOS
                                        
                            'DIA______________________________________________________________________
                                    If VTiempoProgramadoD = 0 Then
                                        VTiempoProgramadoD = 0
                                    Else
                                            VTiempoProgramadoD = VTiempoProgramadoD / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE DIA
                                    If VParosSD = 0 Then
                                        VParosSD = 0
                                    Else
                                        VParosSD = VParosSD / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE DIA
                                    If VParosND = 0 Then
                                        VParosND = 0
                                    Else
                                        VParosND = VParosND / 60
                                    End If
                                    
                            'NOCHE______________________________________________________________________
                                    If VTiempoProgramadoN = 0 Then
                                        VTiempoProgramadoN = 0
                                    Else
                                            VTiempoProgramadoN = VTiempoProgramadoN / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE NOCHE
                                    If VParosSN = 0 Then
                                        VParosSN = 0
                                    Else
                                        VParosSN = VParosSN / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE NOCHE
                                    If VParosNN = 0 Then
                                        VParosNN = 0
                                    Else
                                        VParosNN = VParosNN / 60
                                    End If
                                    
                                        
                                    If Err > 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                        Err.Clear
                                    End If
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________

                        
                        'BUSCA LA DESCRIPCION DE LA LINEA
                        Set RBuscaDescripcionLinea = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip, UnidadMedida, TipoProducto From Lineas Where Linea = '" & VLinea & "'")
                            Else
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip, UnidadMedida, TipoProducto From Lineas Where UPPER(Linea) = '" & UCase(VLinea) & "'")
                            End If
                        
                        'INICIALIZA UN RECORDSET PARA AGREGAR DATOS A LA BASE DE DATOS
                                                VTexto = "'" & RBuscaDescripcionLinea(0) & "', "
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "#" & Format(VFechaInicial, "mm/dd/yyyy") & "#, " 'FECHA
                                                Else 'ORACLE
                                                     VTexto = VTexto & "To_Date('" & VFechaInicial & "', 'dd/mm/yyyy'), " 'FECHA
                                                End If
                                                VTexto = VTexto & VTiempoProgramadoD + VTiempoProgramadoN & ", "
                                                VTexto = VTexto & VProduccionD & ", "
                                                VTexto = VTexto & VProduccionN & ", "
                                                VTexto = VTexto & VParosSD & ", "
                                                VTexto = VTexto & VParosSN & ", "
                                                VTexto = VTexto & VProduccionD + VProduccionN & ", " 'VTiempoRealProducidoD + VTiempoRealProducidoN
                                                VTexto = VTexto & VParosND & ", "
                                                VTexto = VTexto & VParosNN & ", "
                                                VTexto = VTexto & VPCD & ", "
                                                VTexto = VTexto & VPNCD & ", "
                                                VTexto = VTexto & VPDD & ", "
                                                VTexto = VTexto & VPCN & ", "
                                                VTexto = VTexto & VPNCN & ", "
                                                VTexto = VTexto & VPDN & ", "
                                                VTexto = VTexto & VTotalProduccion & ", "
                                                VTexto = VTexto & VEficienciaRealD & ", "
                                                VTexto = VTexto & VEficienciaRealN & ", "
                                                VTexto = VTexto & VPorcentajeLinea & ", "
                                                VTexto = VTexto & VPorcentajeRechazo & ", "
                                                VTexto = VTexto & Format(VPorcentajeDesperdicio, "#,###,##0.00") & ", "
                                                VTexto = VTexto & VFactor1DN & ", "
                                                VTexto = VTexto & VFactor5DN & ", "
                                                VTexto = VTexto & VVelocidadTeoricaDia & ", "
                                                VTexto = VTexto & VVelocidadRealDia & ", "
                                                VTexto = VTexto & VVelocidadTeoricaNoche & ", "
                                                VTexto = VTexto & VVelocidadRealNoche & ", '"
                                                VTexto = VTexto & RBuscaDescripcionLinea(1) & "', "
                                                VTexto = VTexto & VToneladasPC & ", "
                                                VTexto = VTexto & VToneladasPNC & ", "
                                                VTexto = VTexto & Format(VToneladasDes, "#,###,##0.00") & ", "
                                                VTexto = VTexto & VUnidadesPC & ", "
                                                VTexto = VTexto & VUnidadesPNC & ", "
                                                VTexto = VTexto & VUnidadesDes & ", '"
                                                VTexto = VTexto & RBuscaDescripcionLinea(2) & "', "
                                                VTexto = VTexto & (VParosCFD + VParosCFN) & ", "
                                                VTexto = VTexto & (VParosMPD + VParosMPN)
                            
                                                Conexion.Execute "Insert Into ReporteEjecutivoDia2 Values(" & VTexto & ")"
                                      
                                    If Err > 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                            Err.Clear
                                            
                                    End If
                                    
                            VFechaInicial = VFechaInicial + 1
                                    
                    Else ' IF DE CONDICION DE LA QUE NO ENCUENTRA DATOS EN RANGO DE FECHAS
                            
                            'A LA FECHA ACTUAL LE SUMA 1 PARA INCREMENTAR LA FECHA
                            'LA EFICIENCIA SE SACA POR DIA
                            VFechaInicial = VFechaInicial + 1
                    End If
                
                        
                Loop
                
                
                'SIGUE AL SIGUIENTE REGISTRO DE LINEA
                RSeleccionaLineas.MoveNext
  Loop
  
  
                
'EMPIEZA TODO EL PROCESO OTRA VEZ PARA SACAR EL ACUMULADO
'_____________________________________________________________________________________________


            'SELECCIONA LAS LINEAS DE ACUERDO A LA OPCION
            If OptEfi.Item(0).Value = True = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Linea = '" & TxtEfiGru.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Linea) = '" & UCase(TxtEfiGru.Text) & "'")
                            End If
            ElseIf OptEfi.Item(1).Value = True Then
                        Set RSeleccionaLineas = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where Grupo = '" & TxtEfiGru.Text & "'")
                            Else
                                Call Abrir_Recordset(RSeleccionaLineas, "Select Linea From Lineas Where UPPER(Grupo) = '" & UCase(TxtEfiGru.Text) & "'")
                            End If
            End If
             
             If RSeleccionaLineas.RecordCount > 0 Then
             Else
                    'MsgBox "Linea O Lineas No Existen ", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
             End If
  
  
  
  'CREA UN CICLO CON LAS LINEAS POSIBLES DE ACUERDO A LA OPCION ELEGIDA
  Do Until RSeleccionaLineas.EOF
                        
                    'ASIGNA LA LINEA QUE ES SELECCIONADA
                    VLinea = RSeleccionaLineas!Linea
                            
                    'FECHA DE INICIO DEL RANGO
                    VFechaInicial = DTPFecEfiAcu.Value
                    'FECHA DEL FINAL DEL RANGO
                    VFechaFinal = DTPFecEfiFinAcu.Value
                        
                
                Do Until VFechaInicial > VFechaFinal
                            
                            VAToneladasPC = 0
                            VAToneladasPNC = 0
                            VAToneladasDes = 0
                            VAProduccionPC = 0
                            VAProduccionPNC = 0
                            VAProduccionDes = 0
                            VAUnidadesPC = 0
                            VAUnidadesPNC = 0
                            VAUnidadesDes = 0
                            
                        
                        
'VERIFICA SI HAY DATOS EN LA PRESENTE FECHA Y SI NO HAY PASA A LA SIGUIENTE FECHA
'ESTO NOS SIRVE PARA CUANDO SAQUEMOS EL REPORTE DE EFICIENCIA NO TOME EN CUENTA LOS DIAS QUE
'NO SE TRABAJO PORQUE AFECTA LA EFICIENCIA DE LINEA Y PLANTA
                Set RCapturaParos = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Linea = '" & VLinea & "'")
                    Else
                        Call Abrir_Recordset(RCapturaParos, "Select * From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                    End If
                    If RCapturaParos.RecordCount > 0 Then
                        'NO HACE NADA SI HAY DATOS ESTA BIEN
                    
                                
                  
                  '*******  TIEMPO PROGRAMADO Y VELOCIDAD **************************************************
  
               '*******  TIEMPO PROGRAMADO Y VELOCIDAD **************************************************
  
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE DIA
                         Set RTiempoProgramadoD = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '1' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoD, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '1' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoD.RecordCount > 0 Then
                                            VTiempoProgramadoD = RTiempoProgramadoD!HorasProgramadas
                                            VVelocidadTeoricaDia = RTiempoProgramadoD!VelocidadTeorica
                                            VVelocidadRealDia = RTiempoProgramadoD!VelocidadReal
                                            VGrupoDia = RTiempoProgramadoD!Grupo
                                            VParosND = RTiempoProgramadoD!ParoN / 60
                                            VParosSD = RTiempoProgramadoD!Paros / 60
                                            VProduccionD = RTiempoProgramadoD!ParoP / 60
                                            VPCD = RTiempoProgramadoD!ProductoConforme
                                            VPNCD = RTiempoProgramadoD!ProductoNoConforme
                                            VPDD = RTiempoProgramadoD!Desperdicio
                                            VEficienciaRealD = RTiempoProgramadoD!Eficiencia
                                            VParosCFD = RTiempoProgramadoD!ParoCF
                                            VParosMPD = RTiempoProgramadoD!ParoMP
                                Else
                                            VTiempoProgramadoD = 0
                                            VVelocidadTeoricaDia = 0
                                            VVelocidadRealDia = 0
                                            VGrupoDia = ""
                                            VParosND = 0
                                            VParosSD = 0
                                            VProduccionD = 0
                                            VPCD = 0
                                            VPNCD = 0
                                            VPDD = 0
                                            VEficienciaRealD = 0
                                            VParosCFD = 0
                                            VParosMPD = 0
                             End If
                                               
                        
                        'TIEMPO PROGRAMADO Y VELOCIDAD REAL DE LA MAQUINA TURNO DE NOCHE
                         Set RTiempoProgramadoN = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And Turno = '2' And Linea = '" & VLinea & "'")
                                Else
                                    Call Abrir_Recordset(RTiempoProgramadoN, "Select HorasProgramadas, VelocidadTeorica, VelocidadReal, Grupo, ParoS, ParoN, ParoP, ProductoConforme, ProductoNoConforme, Desperdicio, Eficiencia, ParoCF, ParoMP From EncabezadoCapturaParos Where Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(Turno) = '2' And UPPER(Linea) = '" & UCase(VLinea) & "'")
                                End If
                             If RTiempoProgramadoN.RecordCount > 0 Then
                                            VTiempoProgramadoN = RTiempoProgramadoN!HorasProgramadas
                                            VVelocidadTeoricaNoche = RTiempoProgramadoN!VelocidadTeorica
                                            VVelocidadRealNoche = RTiempoProgramadoN!VelocidadReal
                                            VGrupoNoche = RTiempoProgramadoN!Grupo
                                            VParosNN = RTiempoProgramadoN!ParoN / 60
                                            VParosSN = RTiempoProgramadoN!Paros / 60
                                            VProduccionN = RTiempoProgramadoN!ParoP / 60
                                            VPCN = RTiempoProgramadoN!ProductoConforme
                                            VPNCN = RTiempoProgramadoN!ProductoNoConforme
                                            VPDN = RTiempoProgramadoN!Desperdicio
                                            VEficienciaRealN = RTiempoProgramadoN!Eficiencia
                                            VParosCFN = RTiempoProgramadoN!ParoCF
                                            VParosMPN = RTiempoProgramadoN!ParoMP
                             Else
                                            VTiempoProgramadoN = 0
                                            VVelocidadTeoricaNoche = 0
                                            VVelocidadRealNoche = 0
                                            VGrupoNoche = ""
                                            VParosNN = 0
                                            VParosSN = 0
                                            VProduccionN = 0
                                            VPCN = 0
                                            VPNCN = 0
                                            VPDN = 0
                                            VEficienciaRealN = 0
                                            VParosCFN = 0
                                            VParosMPN = 0
                             End If
                        
                        
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE DIA
                        VTiempoRealProducidoD = VTiempoProgramadoD - VParosND
                                                    
                        'TIEMPO REAL PRODUCIDO DEL TURNO DE NOCHE
                        If VTiempoProgramadoN = 0 Then
                            VTiempoRealProducidoN = 0
                        Else
                            VTiempoRealProducidoN = VTiempoProgramadoN - VParosNN
                        End If
                                                    
                        'HORAS PRODUCIDAS POR LOS 2 TURNOS
                            VHorasProducidasDN = Format(VTiempoRealProducidoD + VTiempoRealProducidoN, "#,###,##0.00")
                                                    
                        'TOTAL DE PAROS S "NO AFECTAN"
                            VParosDN = Format(VParosND + VParosNN, "#,###,##0.00")
                        
'________________________________________________________________________________________________________________________
         'PARA EL REPORTE EJECUTIVO ACUMULADO
         'BUSCAMOS EL DETALLE DE LA PRODUCCION EN EL RANGO DE FECHAS INDICADO
                            
                            'BUSCAMOS SOLO UN DIA
                            Set RABuscaDetalleProduccion = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RABuscaDetalleProduccion, "Select DP.Orden, DP.ProductoConforme, DP.ProductoNoConforme, DP.Desperdicio From DetalleProduccionPorOrden DP, EncabezadoCapturaParos EP Where EP.Fecha = #" & Format(VFechaInicial, "mm/dd/yyyy") & "# And EP.Linea = '" & VLinea & "' And EP.Documento = DP.Documento")
                                Else
                                    Call Abrir_Recordset(RABuscaDetalleProduccion, "Select DP.Orden, DP.ProductoConforme, DP.ProductoNoConforme, DP.Desperdicio From DetalleProduccionPorOrden DP, EncabezadoCapturaParos EP Where EP.Fecha = To_Date('" & VFechaInicial & "', 'dd/mm/yyyy')" & " And UPPER(EP.Linea) = '" & UCase(VLinea) & "' And EP.Documento = DP.Documento")
                                End If
                                
                                If RABuscaDetalleProduccion.RecordCount > 0 Then
                                    Do Until RABuscaDetalleProduccion.EOF
                                                'ASIGNAMOS VARIABLES
                                                VAOrdenDetalle = RABuscaDetalleProduccion(0)
                                                VAProduccionPC = RABuscaDetalleProduccion(1)
                                                VAProduccionPNC = RABuscaDetalleProduccion(2)
                                                VAProduccionDes = RABuscaDetalleProduccion(3)
                                                
                                                'AHORA BUSCAMOS LA ORDEN
                                                Set RABuscaOrden = New ADODB.Recordset
                                                    If GOrigenDeDatos = "AmaproAccess" Then
                                                        Call Abrir_Recordset(RABuscaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & VAOrdenDetalle & "'")
                                                    Else
                                                        Call Abrir_Recordset(RABuscaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(VAOrdenDetalle) & "'")
                                                    End If
                                                    'ASIGNAMOS LA FICHA TECNICA QUE USA LA ORDEN
                                                    If RABuscaOrden.RecordCount > 0 Then
                                                        VAFichaTecnicaOrden = RABuscaOrden!FichaTecnica
                                                            'BUSCAMOS LA FICHA TECNICA PARA OBTENER EL PESO
                                                            Set RABuscaFichaTecnica = New ADODB.Recordset
                                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                                    Call Abrir_Recordset(RABuscaFichaTecnica, "Select PesoxUnidad From FichaTecnica Where Esp_Tec = '" & VAFichaTecnicaOrden & "'")
                                                                Else
                                                                    Call Abrir_Recordset(RABuscaFichaTecnica, "Select PesoxUnidad From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(VAFichaTecnicaOrden) & "'")
                                                                End If
                                                                'ASIGNAMOS EL PESO
                                                                If RABuscaFichaTecnica.RecordCount > 0 Then
                                                                    VAPesoFichaTecnica = RABuscaFichaTecnica!PesoxUnidad
                                                                Else
                                                                    VAPesoFichaTecnica = 0
                                                                End If
                                                    Else
                                                        VAFichaTecnicaOrden = ""
                                                    End If
                                                    'ASIGNAMOS EL PESO A LAS VARIABLES
                                                    VAToneladasCalculoPC = ((VAProduccionPC * VAPesoFichaTecnica) / 1000)
                                                    VAToneladasPC = VAToneladasPC + VAToneladasCalculoPC
                                                    VAToneladasCalculoPNC = ((VAProduccionPNC * VAPesoFichaTecnica) / 1000)
                                                    VAToneladasPNC = VAToneladasPNC + VAToneladasCalculoPNC
                                                    VAToneladasCalculoDes = ((VAProduccionDes * VAPesoFichaTecnica) / 1000)
                                                    VAToneladasDes = VAToneladasDes + VAToneladasCalculoDes
                                                    'UNIDADES
                                                    VAUnidadesPC = VAUnidadesPC + VAProduccionPC
                                                    VAUnidadesPNC = VAUnidadesPNC + VAProduccionPNC
                                                    VAUnidadesDes = VAUnidadesDes + VAProduccionDes
                                                
                                        RABuscaDetalleProduccion.MoveNext
                                    Loop
                                Else
                                End If
                        
                        
                        
                        'EL TOTAL DE LA PRODUCCION ES LA SUMA DEL PRODUCTO CONFORME Y NO CONFORME NO INCLUYE EL DESPERDICIO
                        'TOTAL PRODUCCION
                        VTotalProduccion = VPCD + VPNCD + VPCN + VPNCN
                        'TOTAL PRODUCCION DE DIA
                        VTotalProduccionD = VPCD + VPNCD
                        'TOTAL PRODUCCION DE NOCHE
                        VTotalProduccionN = VPCN + VPNCN
                                                
                        
                        'SELECCIONA LA VELOCIDAD TEORICA DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaNoche
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = VVelocidadTeoricaDia
                        ElseIf (VVelocidadTeoricaDia > 0 And VVelocidadTeoricaNoche > 0) Then
                            VVelocidadTeoricaLinea = ((VVelocidadTeoricaDia + VVelocidadTeoricaNoche) / 2)
                        ElseIf (VVelocidadTeoricaDia <= 0 And VVelocidadTeoricaNoche <= 0) Then
                            VVelocidadTeoricaLinea = 0
                        End If
                        
                        'SELECCIONA LA VELOCIDAD REAL DE LA LINEA EN BASE A LOS DOS TURNOS
                        If (VVelocidadRealDia <= 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = VVelocidadRealNoche
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = VVelocidadRealDia
                        ElseIf (VVelocidadRealDia > 0 And VVelocidadRealNoche > 0) Then
                            VVelocidadRealLinea = ((VVelocidadRealDia + VVelocidadRealNoche) / 2)
                        ElseIf (VVelocidadRealDia <= 0 And VVelocidadRealNoche <= 0) Then
                            VVelocidadRealLinea = 0
                        End If
                        
                                                                                                                                                                        
                                                                            
  'CONVIERTE LAS VARIABLES DE HORAS A MINUTOS PARA CALCULOS ________________________________________________________
                            VTiempoProgramadoD = VTiempoProgramadoD * 60
                            
                            VParosSD = VParosSD * 60
                            VParosND = VParosND * 60
                            
                        'NOCHE
                            VTiempoProgramadoN = VTiempoProgramadoN * 60
                            VParosSN = VParosSN * 60
                            VParosNN = VParosNN * 60
                                        
                'NOCHE
                'SI EL TIEMPO PROGRAMADO ES CER0 NO HACE NADA
                If VTiempoProgramadoN > 0 Then
                                         
                Else
                        VFactor1N = 0
                        VFactor2N = 0
                        VFactor3N = 0
                        VFactor4N = 0
                        VFactor5N = 0
                End If
                           
'_______________________________________________________________________________________________________________________
                        
                        
                '% DE LINEA___________________________________________________________________________________________
                '_____________________________________________________________________________________________________
                                        
                            'FACTOR 1_________________________________________________________________________________
                                    VFactor1DN = (VTiempoProgramadoD - (VParosND + VParosSD)) + (VTiempoProgramadoN - (VParosNN + VParosSN))
                                    VFactor1DN = VFactor1DN * VVelocidadRealLinea
                                    If VFactor1DN = 0 Then
                                    Else
                                        VFactor1DN = VTotalProduccion / VFactor1DN
                                    End If
                                    
                                    'SI EL FACTOR 1 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 1
                                    If VFactor1DN = 0 Then
                                        VFactor1DN = 1
                                    End If
                                                        
                            'FACTOR 2_________________________________________________________________________________
                                    If (VPNCD + VPNCN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor2DN = VTotalProduccion - (VPNCD + VPNCN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor2DN = VFactor2DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 3_________________________________________________________________________________
                                    If (VPDD + VPDN) > VTotalProduccion Then
                                        VFactor2DN = 0
                                    Else
                                        VFactor3DN = VTotalProduccion - (VPDD + VPDN)
                                        If VTotalProduccion = 0 Then
                                        Else
                                            VFactor3DN = VFactor3DN / VTotalProduccion
                                        End If
                                    End If
                                                        
                            'FACTOR 4_________________________________________________________________________________
                                    VFactor4DN = (((VTiempoProgramadoD - VParosND) - VParosSD) + ((VTiempoProgramadoN - VParosNN) - VParosSN))
                                    If (VTiempoProgramadoD + VTiempoProgramadoN) = 0 Then
                                    ElseIf ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)) = 0 Then
                                    Else
                                        VFactor4DN = (VFactor4DN / ((VTiempoProgramadoD - VParosND) + (VTiempoProgramadoN - VParosNN)))
                                    End If
                                    
                            'FACTOR 5___________________________________________________________
                                    If VVelocidadTeoricaLinea = 0 Then
                                        VFactor5DN = 0
                                    Else
                                        VFactor5DN = VVelocidadRealLinea / VVelocidadTeoricaLinea
                                    End If
                                    
                                    'SI EL FACTOR 5 SE QUEDA EN CERO SE CAMBIA A 1 PARA QUE NO AFECTE LA EFICIENCIA
                                    'ACUMULADA AL SUMAR Y SACAR EL PROMEDIO DE TODOS LOS FACTORES 5
                                    If VFactor5DN = 0 Then
                                        VFactor5DN = 1
                                    End If
                                                                                
                            'EFICIENCIA DE LINEA______________________________________________________________________
                                    VPorcentajeLinea = VFactor1DN * VFactor2DN * VFactor3DN * VFactor4DN * VFactor5DN * 100
                                   
                '% DE RECHAZO__________________________________________________________________________________________
                '______________________________________________________________________________________________________
                                    VPorcentajeRechazo = VPNCD + VPNCN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeRechazo = VPorcentajeRechazo / VTotalProduccion
                                        VPorcentajeRechazo = VPorcentajeRechazo * 100
                                    End If
                        
                        
                '% DE DESPERDICIO______________________________________________________________________________________
                '______________________________________________________________________________________________________
                
                                    VPorcentajeDesperdicio = VPDD + VPDN
                                    If VTotalProduccion = 0 Then
                                    Else
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio / VTotalProduccion
                                        VPorcentajeDesperdicio = VPorcentajeDesperdicio * 100
                                    End If
                                    
'_______________________________________________________________________________________________________________________
'************************************************* HORAS A MINUTOS *****************************************************
'_______________________________________________________________________________________________________________________

                                                
                'CONVIERTE LAS VARIABLES A HORAS PARA IMPRIMIR LOS DATOS
                                        
                            'DIA______________________________________________________________________
                                    If VTiempoProgramadoD = 0 Then
                                        VTiempoProgramadoD = 0
                                    Else
                                        VTiempoProgramadoD = VTiempoProgramadoD / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE DIA
                                    If VParosSD = 0 Then
                                        VParosSD = 0
                                    Else
                                        VParosSD = VParosSD / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE DIA
                                    If VParosND = 0 Then
                                        VParosND = 0
                                    Else
                                        VParosND = VParosND / 60
                                    End If
                                    
                            'NOCHE______________________________________________________________________
                                    If VTiempoProgramadoN = 0 Then
                                        VTiempoProgramadoN = 0
                                    Else
                                        VTiempoProgramadoN = VTiempoProgramadoN / 60
                                    End If
                                    
                                    'PAROS QUE SI AFECTAN DE NOCHE
                                    If VParosSN = 0 Then
                                        VParosSN = 0
                                    Else
                                        VParosSN = VParosSN / 60
                                    End If
                                    
                                    'PAROS QUE NO AFECTAN DE NOCHE
                                    If VParosNN = 0 Then
                                        VParosNN = 0
                                    Else
                                        VParosNN = VParosNN / 60
                                    End If
                                    
                                        
                                    If Err > 0 Then
                                        MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                        Err.Clear
                                    End If
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________
'_______________________________________________________________________________________________________________________

                        
                        'BUSCA LA DESCRIPCION DE LA LINEA
                        Set RBuscaDescripcionLinea = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip, UnidadMedida, TipoProducto From Lineas Where Linea = '" & VLinea & "'")
                            Else
                                Call Abrir_Recordset(RBuscaDescripcionLinea, "Select Descrip, UnidadMedida, TipoProducto From Lineas Where UPPER(Linea) = '" & UCase(VLinea) & "'")
                            End If
                                               
                                                VTexto = "'" & RBuscaDescripcionLinea(0) & "', "
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "#" & Format(VFechaInicial, "mm/dd/yyyy") & "#, " 'FECHA
                                                Else 'ORACLE
                                                     VTexto = VTexto & "To_Date('" & VFechaInicial & "', 'dd/mm/yyyy'), " 'FECHA
                                                End If
                                                VTexto = VTexto & VTiempoProgramadoD + VTiempoProgramadoN & ", "
                                                VTexto = VTexto & VProduccionD & ", "
                                                VTexto = VTexto & VProduccionN & ", "
                                                VTexto = VTexto & VParosSD & ", "
                                                VTexto = VTexto & VParosSN & ", "
                                                VTexto = VTexto & VProduccionD + VProduccionN & ", " 'VTiempoRealProducidoD + VTiempoRealProducidoN
                                                VTexto = VTexto & VParosND & ", "
                                                VTexto = VTexto & VParosNN & ", "
                                                VTexto = VTexto & VPCD & ", "
                                                VTexto = VTexto & VPNCD & ", "
                                                VTexto = VTexto & VPDD & ", "
                                                VTexto = VTexto & VPCN & ", "
                                                VTexto = VTexto & VPNCN & ", "
                                                VTexto = VTexto & VPDN & ", "
                                                VTexto = VTexto & VTotalProduccion & ", "
                                                VTexto = VTexto & VEficienciaRealD & ", "
                                                VTexto = VTexto & VEficienciaRealN & ", "
                                                VTexto = VTexto & VPorcentajeLinea & ", "
                                                VTexto = VTexto & VPorcentajeRechazo & ", "
                                                VTexto = VTexto & Format(VPorcentajeDesperdicio, "#,###,##0.00") & ", "
                                                VTexto = VTexto & VFactor1DN & ", "
                                                VTexto = VTexto & VFactor5DN & ", "
                                                VTexto = VTexto & VVelocidadTeoricaDia & ", "
                                                VTexto = VTexto & VVelocidadRealDia & ", "
                                                VTexto = VTexto & VVelocidadTeoricaNoche & ", "
                                                VTexto = VTexto & VVelocidadRealNoche & ", '"
                                                VTexto = VTexto & RBuscaDescripcionLinea(1) & "', "
                                                VTexto = VTexto & VAToneladasPC & ", "
                                                VTexto = VTexto & VAToneladasPNC & ", "
                                                VTexto = VTexto & Format(VAToneladasDes, "#,###,##0.00") & ", "
                                                VTexto = VTexto & VAUnidadesPC & ", "
                                                VTexto = VTexto & VAUnidadesPNC & ", "
                                                VTexto = VTexto & VAUnidadesDes & ", '"
                                                VTexto = VTexto & RBuscaDescripcionLinea(2) & "', "
                                                VTexto = VTexto & (VParosCFD + VParosCFN) & ", "
                                                VTexto = VTexto & (VParosMPD + VParosMPN)
                            
                                                 Conexion.Execute "Insert Into ReporteEjecutivoAcumulado2 Values(" & VTexto & ")"
                                      
                                               
                                    If Err > 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description & " " & Err.Source, vbOKOnly, "Informacion"
                                            Err.Clear
                                    End If
                                    
                            VFechaInicial = VFechaInicial + 1
                                    
                    Else ' IF DE CONDICION DE LA QUE NO ENCUENTRA DATOS EN RANGO DE FECHAS
                            
                            'A LA FECHA ACTUAL LE SUMA 1 PARA INCREMENTAR LA FECHA
                            'LA EFICIENCIA SE SACA POR DIA
                            VFechaInicial = VFechaInicial + 1
                    End If
                
                        
                Loop
                
                
                'SIGUE AL SIGUIENTE REGISTRO DE LINEA
                RSeleccionaLineas.MoveNext
  Loop
  
            
                  
                
End Sub

