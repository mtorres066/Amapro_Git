VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Reportes 
   BackColor       =   &H000000FF&
   Caption         =   "Reportes Generales De Produccion"
   ClientHeight    =   6840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12150
   Icon            =   "Reportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
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
      Height          =   6615
      Left            =   120
      TabIndex        =   69
      Top             =   120
      Visible         =   0   'False
      Width           =   11895
      Begin MSDataGridLib.DataGrid DbGridBuscar 
         Height          =   5295
         Left            =   120
         TabIndex        =   73
         Top             =   1080
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   9340
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
         Height          =   615
         Left            =   10560
         Picture         =   "Reportes.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   78
         ToolTipText     =   "Salida"
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   1
         Left            =   2040
         TabIndex        =   71
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2040
         TabIndex        =   72
         ToolTipText     =   "digite sus datos para buscar"
         Top             =   720
         Width           =   4455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   70
         Top             =   360
         Value           =   -1  'True
         Width           =   1695
      End
      Begin VB.Label LblBuscar 
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
         Left            =   360
         TabIndex        =   77
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdSalida 
      Cancel          =   -1  'True
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
      Height          =   2000
      Left            =   10320
      Picture         =   "Reportes.frx":24B4
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2760
      Width           =   1575
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
      Height          =   2000
      Left            =   10320
      Picture         =   "Reportes.frx":2DE6
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   720
      Width           =   1575
   End
   Begin TabDlg.SSTab TabReportes 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   11668
      _Version        =   393216
      Tabs            =   7
      TabsPerRow      =   4
      TabHeight       =   882
      BackColor       =   255
      TabCaption(0)   =   "Produccion"
      TabPicture(0)   =   "Reportes.frx":3530
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblProduccion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblDesPro"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblProduccion2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblDesPro2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Lbltur"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "OptProduccion(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "PFecProIni"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "PFecProFin"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TxtTextoProduccion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "OptProduccion(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "FrameOpciones"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "FrameCalidad"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "OptProduccion(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtTextoProduccion2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "OptProduccion(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "OptProduccion(5)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "OptProduccion(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "OptProduccion(2)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TxtTur"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "OptProduccion(3)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Frame2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Produccion Liberada"
      TabPicture(1)   =   "Reportes.frx":384A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LblLibEti"
      Tab(1).Control(1)=   "LblLibDes"
      Tab(1).Control(2)=   "LblLibFecIni"
      Tab(1).Control(3)=   "LblFecFin(0)"
      Tab(1).Control(4)=   "OptProduccionLiberado(0)"
      Tab(1).Control(5)=   "OptProduccionLiberado(1)"
      Tab(1).Control(6)=   "OptProduccionLiberado(2)"
      Tab(1).Control(7)=   "Frame1"
      Tab(1).Control(8)=   "TxtLib"
      Tab(1).Control(9)=   "DtpLibFecIni"
      Tab(1).Control(10)=   "DtpLibFecFin"
      Tab(1).ControlCount=   11
      TabCaption(2)   =   "Rutinas"
      TabPicture(2)   =   "Reportes.frx":3B64
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "LblDesRut"
      Tab(2).Control(1)=   "LblRutinas(2)"
      Tab(2).Control(2)=   "LblRutinas(1)"
      Tab(2).Control(3)=   "LblRutinas(0)"
      Tab(2).Control(4)=   "LblHas"
      Tab(2).Control(5)=   "MskFecFin"
      Tab(2).Control(6)=   "MskFecRut"
      Tab(2).Control(7)=   "TxtHorRut"
      Tab(2).Control(8)=   "FrameRutinas"
      Tab(2).Control(9)=   "OptRutFicTec"
      Tab(2).Control(10)=   "OptRutArr"
      Tab(2).Control(11)=   "OptRutDet"
      Tab(2).Control(12)=   "TxtLin"
      Tab(2).Control(13)=   "OptRutRep"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "Batch"
      TabPicture(3)   =   "Reportes.frx":443E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "OptBatNum"
      Tab(3).Control(1)=   "TxtBatch"
      Tab(3).Control(2)=   "FrameBatchTipoDeReporte"
      Tab(3).Control(3)=   "FrameBatchUnidadMedida"
      Tab(3).Control(4)=   "TxtBatLin"
      Tab(3).Control(5)=   "Label8(0)"
      Tab(3).Control(6)=   "Label8(1)"
      Tab(3).Control(7)=   "LblBatDesLin"
      Tab(3).ControlCount=   8
      TabCaption(4)   =   "Fichas Tecnicas"
      TabPicture(4)   =   "Reportes.frx":4D18
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "FrameFichaTecnica2"
      Tab(4).Control(1)=   "OptFichaTecnica(0)"
      Tab(4).Control(2)=   "TxtFichaTecnica"
      Tab(4).Control(3)=   "OptFichaTecnica(3)"
      Tab(4).Control(4)=   "FrameFichaTecnica"
      Tab(4).Control(5)=   "LblFichaTecnica"
      Tab(4).Control(6)=   "LblDesFicTec"
      Tab(4).ControlCount=   7
      TabCaption(5)   =   "Catalogos"
      TabPicture(5)   =   "Reportes.frx":55F2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "LblDesHoj"
      Tab(5).Control(1)=   "Label6"
      Tab(5).Control(2)=   "LblDesCat"
      Tab(5).Control(3)=   "TxtVar"
      Tab(5).Control(4)=   "OptVarCod"
      Tab(5).ControlCount=   5
      TabCaption(6)   =   "Defectos"
      TabPicture(6)   =   "Reportes.frx":5ECC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "LblDefectos"
      Tab(6).Control(1)=   "OptDefCod"
      Tab(6).Control(2)=   "OptDefTip"
      Tab(6).Control(3)=   "TxtDefectos"
      Tab(6).ControlCount=   4
      Begin VB.Frame Frame2 
         Caption         =   "Tipo Busqueda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   3720
         TabIndex        =   108
         Top             =   3120
         Width           =   2295
         Begin VB.OptionButton OptProTipBus 
            Caption         =   "Palabra Inicial"
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
            Left            =   120
            TabIndex        =   110
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton OptProTipBus 
            Caption         =   "Cualquier Palabra"
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
            Left            =   120
            TabIndex        =   109
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.OptionButton OptProduccion 
         Caption         =   "Por Fechas Y Descripcion"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   107
         Top             =   2400
         Width           =   2295
      End
      Begin VB.TextBox TxtTur 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   29
         Top             =   6120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.OptionButton OptProduccion 
         Caption         =   "Por Fechas Y Linea Y Turno"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   105
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Frame FrameFichaTecnica2 
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
         Height          =   1455
         Left            =   -69360
         TabIndex        =   102
         Top             =   1320
         Width           =   2415
         Begin VB.OptionButton OptFicTecTipRep 
            Caption         =   "Resumen x Grupo"
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
            Height          =   192
            Index           =   2
            Left            =   120
            TabIndex        =   112
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton OptFicTecTipRep 
            Caption         =   "Resumen x Catalogo"
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
            Height          =   192
            Index           =   1
            Left            =   120
            TabIndex        =   104
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton OptFicTecTipRep 
            Caption         =   "Detalle"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Index           =   0
            Left            =   120
            TabIndex        =   103
            Top             =   360
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.OptionButton OptProduccion 
         Caption         =   "Por Batch Y Linea"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   99
         Top             =   3840
         Width           =   1815
      End
      Begin VB.OptionButton OptProduccion 
         Caption         =   "Por Fechas Y Ficha Tecnica"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   3
         Top             =   2040
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DtpLibFecFin 
         Height          =   255
         Left            =   -71640
         TabIndex        =   36
         Top             =   3960
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61079555
         CurrentDate     =   37550
      End
      Begin MSComCtl2.DTPicker DtpLibFecIni 
         Height          =   255
         Left            =   -71640
         TabIndex        =   35
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61079555
         CurrentDate     =   37550
      End
      Begin VB.TextBox TxtLib 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71640
         TabIndex        =   37
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Frame Frame1 
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
         Height          =   1815
         Left            =   -69600
         TabIndex        =   30
         Top             =   1320
         Width           =   2895
         Begin VB.OptionButton OptProLibOpc 
            Caption         =   "Detalle"
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
            Left            =   120
            TabIndex        =   31
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OptProLibOpc 
            Caption         =   "Resumen"
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
            Index           =   1
            Left            =   120
            TabIndex        =   32
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton OptProLibOpc 
            Caption         =   "Detalle Con Tarimas"
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
            Left            =   120
            TabIndex        =   33
            Top             =   720
            Width           =   2535
         End
         Begin VB.OptionButton OptProLibOpc 
            Caption         =   "Detalle Con Defectos"
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
            TabIndex        =   34
            Top             =   1080
            Width           =   2295
         End
      End
      Begin VB.OptionButton OptProduccionLiberado 
         Caption         =   "Producto Liberado Por Fechas"
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   28
         Top             =   2160
         Width           =   2655
      End
      Begin VB.OptionButton OptProduccionLiberado 
         Caption         =   "Producto Liberado Por Fechas Y Linea"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   26
         Top             =   1800
         Width           =   3255
      End
      Begin VB.OptionButton OptProduccionLiberado 
         Caption         =   "Producto Liberado Por Fechas Y Grupo De Linea"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   24
         Top             =   1440
         Width           =   3975
      End
      Begin VB.OptionButton OptRutRep 
         Caption         =   "Registro de Inspeccion De Rutinas"
         Height          =   195
         Left            =   -74880
         TabIndex        =   38
         Top             =   1320
         Width           =   3015
      End
      Begin VB.TextBox TxtLin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70560
         MaxLength       =   2
         TabIndex        =   46
         Top             =   3720
         Width           =   1095
      End
      Begin VB.OptionButton OptRutDet 
         Caption         =   "Reporte De Inspeccion de Rutinas Detalle"
         Height          =   195
         Left            =   -74880
         TabIndex        =   40
         Top             =   2040
         Width           =   3855
      End
      Begin VB.OptionButton OptRutArr 
         Caption         =   "Registro de Inspeccion de Rutinas de Arranque"
         Height          =   195
         Left            =   -74880
         TabIndex        =   39
         Top             =   1680
         Width           =   3855
      End
      Begin VB.OptionButton OptRutFicTec 
         Caption         =   "Reporte De Ficha Tecnica De Rutinas"
         Height          =   195
         Left            =   -74880
         TabIndex        =   41
         Top             =   2400
         Width           =   3615
      End
      Begin VB.Frame FrameRutinas 
         Caption         =   "Reporte En Unidad De Medida "
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
         Left            =   -70080
         TabIndex        =   42
         Top             =   1200
         Width           =   3015
         Begin VB.OptionButton OptPulgadas 
            Caption         =   "Pulgadas"
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
            Height          =   195
            Left            =   240
            TabIndex        =   43
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptMilimetros 
            Caption         =   "Milimetros"
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
            Height          =   195
            Left            =   240
            TabIndex        =   44
            Top             =   720
            Value           =   -1  'True
            Width           =   1695
         End
      End
      Begin VB.OptionButton OptBatNum 
         Caption         =   "Numero"
         Height          =   195
         Left            =   -74640
         TabIndex        =   55
         Top             =   3120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtBatch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   56
         Top             =   3600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame FrameBatchTipoDeReporte 
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
         ForeColor       =   &H8000000D&
         Height          =   1455
         Left            =   -74760
         TabIndex        =   49
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
         Begin VB.OptionButton OptBatTipRep 
            Caption         =   "Cliente Con Datos"
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
            Left            =   120
            Picture         =   "Reportes.frx":6466
            TabIndex        =   101
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton OptBatTipRep 
            Caption         =   "Cliente Sin Datos"
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
            Left            =   120
            Picture         =   "Reportes.frx":6D30
            TabIndex        =   51
            Top             =   720
            Width           =   2055
         End
         Begin VB.OptionButton OptBatTipRep 
            Caption         =   "Empresa"
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
            Left            =   120
            Picture         =   "Reportes.frx":CFBA
            TabIndex        =   50
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame FrameBatchUnidadMedida 
         Caption         =   "Unidad De Medida"
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
         Height          =   1335
         Left            =   -72240
         TabIndex        =   52
         Top             =   1200
         Visible         =   0   'False
         Width           =   1935
         Begin VB.OptionButton OptBatMilimetros 
            Caption         =   "Milimetros"
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
            Left            =   240
            Picture         =   "Reportes.frx":D884
            TabIndex        =   53
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptBatPulgadas 
            Caption         =   "Pulgadas"
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
            Left            =   240
            Picture         =   "Reportes.frx":DB8E
            TabIndex        =   54
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.TextBox TxtBatLin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         MaxLength       =   2
         TabIndex        =   57
         Top             =   3960
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.OptionButton OptFichaTecnica 
         Caption         =   "Codigo De Ficha Tecnica"
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   58
         Top             =   1440
         Width           =   2280
      End
      Begin VB.TextBox TxtFichaTecnica 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         TabIndex        =   64
         Top             =   3480
         Width           =   2000
      End
      Begin VB.OptionButton OptFichaTecnica 
         Caption         =   "Catalogo"
         Height          =   195
         Index           =   3
         Left            =   -74640
         TabIndex        =   59
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Frame FrameFichaTecnica 
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
         ForeColor       =   &H00FF0000&
         Height          =   1455
         Left            =   -71760
         TabIndex        =   60
         Top             =   1320
         Width           =   2295
         Begin VB.OptionButton OptFichaTecnica2 
            Caption         =   "Igual a"
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
            Left            =   120
            TabIndex        =   63
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton OptFichaTecnica2 
            Caption         =   "Palabra Inicial"
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
            Left            =   120
            TabIndex        =   61
            Top             =   360
            Value           =   -1  'True
            Width           =   1695
         End
         Begin VB.OptionButton OptFichaTecnica2 
            Caption         =   "Cualquier Palabra"
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
            Left            =   120
            TabIndex        =   62
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.OptionButton OptVarCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74640
         TabIndex        =   65
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox TxtVar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71160
         TabIndex        =   83
         Top             =   3360
         Width           =   2000
      End
      Begin VB.TextBox TxtDefectos 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -70920
         TabIndex        =   68
         Top             =   3240
         Width           =   2000
      End
      Begin VB.OptionButton OptDefTip 
         Caption         =   "Tipo"
         Height          =   195
         Left            =   -74400
         TabIndex        =   67
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton OptDefCod 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   -74400
         TabIndex        =   66
         Top             =   1320
         Width           =   975
      End
      Begin VB.OptionButton OptProduccion 
         Caption         =   "Por Fechas Y Grupo De Linea"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   1
         Top             =   1320
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.TextBox TxtTextoProduccion2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   27
         Top             =   5760
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.OptionButton OptProduccion 
         Caption         =   "Por Fechas Y Numero De Orden"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Frame FrameCalidad 
         Caption         =   "Tipo De Calidad"
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
         Height          =   1815
         Left            =   3720
         TabIndex        =   16
         Top             =   1200
         Width           =   2295
         Begin VB.OptionButton OptCalI 
            Caption         =   "Calidad I"
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
            Left            =   240
            TabIndex        =   111
            Top             =   1440
            Width           =   1455
         End
         Begin VB.OptionButton OptCalR 
            Caption         =   "Calidad R"
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
            Left            =   240
            TabIndex        =   19
            Top             =   1080
            Width           =   1215
         End
         Begin VB.OptionButton OptCalAIC 
            Caption         =   "Calidad A - I - C"
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
            Left            =   240
            TabIndex        =   18
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton OptCalTot 
            Caption         =   "Calidad Total"
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
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1575
         End
      End
      Begin VB.Frame FrameOpciones 
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
         Height          =   3615
         Left            =   6120
         TabIndex        =   6
         Top             =   1200
         Width           =   3255
         Begin VB.OptionButton OptDefCua 
            Caption         =   "Defectos Cuadricula"
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
            Left            =   120
            TabIndex        =   15
            Top             =   3360
            Width           =   2055
         End
         Begin VB.OptionButton OptLinAño 
            Caption         =   "Linea Y Año"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   3000
            Width           =   1455
         End
         Begin VB.OptionButton OptLinMes 
            Caption         =   "Linea Y Mes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   2640
            Width           =   1935
         End
         Begin VB.OptionButton OptFicTecAño 
            Caption         =   "Ficha Tecnica Y Año Y Grafica"
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
            Left            =   120
            TabIndex        =   12
            Top             =   2280
            Width           =   3015
         End
         Begin VB.OptionButton OptFicTecMes 
            Caption         =   "Ficha Tecnica Y Mes y Grafica"
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
            Left            =   120
            TabIndex        =   11
            Top             =   1920
            Width           =   3015
         End
         Begin VB.OptionButton OptOpcDetDef 
            Caption         =   "Detalle Con Defectos"
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
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   2295
         End
         Begin VB.OptionButton OptOpcDetMatPri 
            Caption         =   "Detalle Con Materia Prima"
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
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   2535
         End
         Begin VB.OptionButton OptOpcRes 
            Caption         =   "Resumen"
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
            Left            =   120
            TabIndex        =   10
            Top             =   1560
            Width           =   1215
         End
         Begin VB.OptionButton OptOpcDet 
            Caption         =   "Detalle Solo Tarimas"
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
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.OptionButton OptProduccion 
         Caption         =   "Por Fechas Y Linea"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   4
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox TxtTextoProduccion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   25
         ToolTipText     =   "Doble Click O Signo ""+"" Para Ayuda"
         Top             =   5400
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker PFecProFin 
         Height          =   255
         Left            =   3720
         TabIndex        =   23
         Top             =   5040
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61079555
         CurrentDate     =   36926
      End
      Begin MSComCtl2.DTPicker PFecProIni 
         Height          =   255
         Left            =   3720
         TabIndex        =   22
         Top             =   4680
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   61079555
         CurrentDate     =   36926
      End
      Begin VB.OptionButton OptProduccion 
         Caption         =   "Por Fechas"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   2
         Top             =   1680
         Width           =   2652
      End
      Begin MSMask.MaskEdBox TxtHorRut 
         Height          =   285
         Left            =   -70560
         TabIndex        =   45
         Top             =   3360
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
      Begin MSMask.MaskEdBox MskFecRut 
         Height          =   285
         Left            =   -70560
         TabIndex        =   47
         Top             =   4080
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MskFecFin 
         Height          =   285
         Left            =   -68760
         TabIndex        =   48
         Top             =   4080
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Format          =   "dd/mm/yyyy"
         PromptChar      =   "_"
      End
      Begin VB.Label LblHas 
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
         Height          =   195
         Left            =   -69360
         TabIndex        =   113
         Top             =   4080
         Visible         =   0   'False
         Width           =   510
      End
      Begin VB.Label Lbltur 
         Alignment       =   1  'Right Justify
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
         Height          =   255
         Left            =   2880
         TabIndex        =   106
         Top             =   6120
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label LblDesPro2 
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
         Left            =   6000
         TabIndex        =   100
         Top             =   5760
         Width           =   3735
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
         Index           =   0
         Left            =   -73080
         TabIndex        =   98
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label LblLibFecIni 
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
         Left            =   -73080
         TabIndex        =   97
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label LblLibDes 
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
         Left            =   -69840
         TabIndex        =   96
         Top             =   4440
         Width           =   4815
      End
      Begin VB.Label LblLibEti 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo De Linea"
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
         TabIndex        =   95
         Top             =   4440
         Width           =   2790
      End
      Begin VB.Label LblRutinas 
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
         Index           =   0
         Left            =   -71400
         TabIndex        =   94
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label LblRutinas 
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
         Index           =   1
         Left            =   -71400
         TabIndex        =   93
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label LblRutinas 
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
         Index           =   2
         Left            =   -71400
         TabIndex        =   92
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label LblDesRut 
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
         Left            =   -69360
         TabIndex        =   91
         Top             =   3720
         Width           =   3735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Numero de Batch"
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
         Left            =   -72960
         TabIndex        =   90
         Top             =   3600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   -72960
         TabIndex        =   89
         Top             =   3960
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label LblBatDesLin 
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
         TabIndex        =   88
         Top             =   3960
         Width           =   5415
      End
      Begin VB.Label LblFichaTecnica 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo De Ficha Tecnica"
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
         TabIndex        =   87
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label LblDesFicTec 
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
         Left            =   -69600
         TabIndex        =   86
         Top             =   3480
         Width           =   4335
      End
      Begin VB.Label LblDesCat 
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
         Left            =   -69000
         TabIndex        =   85
         Top             =   3360
         Width           =   5175
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo de Catalogo"
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
         Left            =   -73320
         TabIndex        =   84
         Top             =   3360
         Width           =   2055
      End
      Begin VB.Label LblDefectos 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo De Defecto"
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
         Left            =   -73680
         TabIndex        =   82
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label LblProduccion2 
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
         Height          =   195
         Left            =   660
         TabIndex        =   81
         Top             =   5760
         Width           =   2955
      End
      Begin VB.Label LblDesHoj 
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
         Left            =   -68280
         TabIndex        =   80
         Top             =   3720
         Width           =   4935
      End
      Begin VB.Label LblDesPro 
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
         Left            =   6000
         TabIndex        =   79
         Top             =   5400
         Width           =   3735
      End
      Begin VB.Label Label5 
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
         Left            =   2520
         TabIndex        =   76
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label4 
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
         Left            =   2520
         TabIndex        =   75
         Top             =   4680
         Width           =   1215
      End
      Begin VB.Label LblProduccion 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo De Linea"
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
         Left            =   1080
         TabIndex        =   74
         Top             =   5400
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Reportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaRutinas As New ADODB.Recordset
Dim RReporteRutinas As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset
Dim RBuscaFondo As New ADODB.Recordset
Dim RBuscaHojalata As New ADODB.Recordset
Dim RBuscaVariable As New ADODB.Recordset
Dim RBuscaMinimoMaximo As New ADODB.Recordset
Dim RBuscaBatch As New ADODB.Recordset
Dim RBuscaBatch2 As New ADODB.Recordset
Dim RBuscaBatchDatos As New ADODB.Recordset
Dim RReporteBatch As New ADODB.Recordset
Dim RBuscaMateriaPrima As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VFichaTecnica As String
Dim VCodigoVariable As String
Dim Cont As Integer
Dim VNombreComercial As String
Dim VDescrip As String
Dim VTextoProduccion As String

Dim VDia As String
Dim VMes As String
Dim VAño  As String
Dim VDia2 As String
Dim VMes2 As String
Dim VAño2 As String

Dim RBuscaDefecto As New ADODB.Recordset
Dim VDefecto1 As String
Dim VDefecto2 As String
Dim VDefecto3 As String

'VARIABLES PARA SACAR LA AYUDA DE DATOS
Dim BFichaTecnica As Boolean
Dim BPlatina As Boolean
Dim BTapa As Boolean
Dim BFondo As Boolean

Dim VCabezal As Long
Dim VBatch As Double
Dim i As Integer

Dim VTexto As String


Private Sub CmdImprimir_Click()
On Error Resume Next
MousePointer = 11


'----------------------------------------------------------------------------------------------------------
'                                                  PRODUCCION
'----------------------------------------------------------------------------------------------------------

If TabReportes.Tab = 0 Then
                 VDia = Day(PFecProIni.Value)
                 VMes = Month(PFecProIni.Value)
                 VAño = Year(PFecProIni.Value)
                 VDia2 = Day(PFecProFin.Value)
                 VMes2 = Month(PFecProFin.Value)
                 VAño2 = Year(PFecProFin.Value)
                 
                 'CUALQUIER PALABRA
                 If OptProTipBus.Item(0).Value = True Then
                    VTextoProduccion = "Like '*" & TxtTextoProduccion.Text & "*'"
                 'PALABRA INICIAL
                 ElseIf OptProTipBus.Item(1).Value = True Then
                    VTextoProduccion = "Like '" & TxtTextoProduccion.Text & "*'"
                 End If
                    
                'PRODUCCION POR FECHAS
                If OptProduccion.Item(0).Value = True Then
                          
                        'CrReportes.WindowTitle = "Produccion desde " & PFecProIni.Value & " Hasta " & PFecProFin.Value
                        'CrReportes.BoundReportHeading = "Produccion desde "
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value
                        
                        If OptOpcDet.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalle.rpt"
                            Else
                                GNombreReporte = "ProduccionDetalleO.rpt"
                            End If
                        ElseIf OptOpcRes.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionResumen.rpt"
                            Else
                                GNombreReporte = "ProduccionResumenO.rpt"
                            End If
                        ElseIf OptOpcDetMatPri.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConMateriaPrima.rpt"
                            Else
                                GNombreReporte = "ProduccionDetalleConMateriaPrimaO.rpt"
                            End If
                        ElseIf OptOpcDetDef.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                            Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                            End If
                        ElseIf OptDefCua.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                            Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                            End If
                        
                        ElseIf OptFicTecMes.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaMes.rpt"
                            Else
                                GNombreReporte = "ProduccionPorFichaTecnicaMesO.rpt"
                            End If
                        ElseIf OptFicTecAño.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichatecnicaAño.rpt"
                            Else
                                GNombreReporte = "ProduccionPorFichaTecnicaAñoO.rpt"
                            End If
                        ElseIf OptLinMes.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaMes.rpt"
                            Else
                                GNombreReporte = "ProduccionPorLineaMesO.rpt"
                            End If
                        ElseIf OptLinAño.Value = True Then
                            GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaAño.rpt"
                            Else
                                GNombreReporte = "ProduccionPorLineaAñoO.rpt"
                            End If
                        End If
                'PRODUCCION POR FECHAS Y LINEA
                ElseIf OptProduccion.Item(1).Value = True Then
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value & " Linea " & TxtTextoProduccion.Text
                        
                        If OptOpcDet.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalle.rpt"
                            Else
                                GNombreReporte = "ProduccionDetalleO.rpt"
                            End If
                        ElseIf OptOpcRes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionResumen.rpt"
                            Else
                                GNombreReporte = "ProduccionResumenO.rpt"
                            End If
                        ElseIf OptOpcDetMatPri.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConMateriaPrima.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConMateriaPrimaO.rpt"
                              End If
                        ElseIf OptOpcDetDef.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                              End If
                        ElseIf OptDefCua.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                              End If
                        
                        ElseIf OptFicTecMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaMesO.rpt"
                              End If
                        ElseIf OptFicTecAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaAñoO.rpt"
                              End If
                        ElseIf OptLinMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaMesO.rpt"
                              End If
                        ElseIf OptLinAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaAño.rpt"
                              End If
                        End If
                'PRODUCCION POR FECHAS Y LINEA Y TURNO
                ElseIf OptProduccion.Item(2).Value = True Then
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value & " Linea " & TxtTextoProduccion.Text
                        
                        If OptOpcDet.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleO.rpt"
                              End If
                        ElseIf OptOpcRes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionResumenO.rpt"
                              End If
                        ElseIf OptOpcDetMatPri.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConMateriaPrima.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConMateriaPrimaO.rpt"
                              End If
                        ElseIf OptOpcDetDef.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                              End If
                        ElseIf OptDefCua.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectosCuadricula.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosCuadriculaO.rpt"
                              End If
                        
                        ElseIf OptFicTecMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaMesO.rpt"
                              End If
                        ElseIf OptFicTecAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaAñoO.rpt"
                              End If
                        ElseIf OptLinMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaMesO.rpt"
                              End If
                              
                        ElseIf OptLinAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} " & VTextoProduccion & " And {Produccion.Turno} = '" & TxtTur.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLienaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaAñoO.rpt"
                              End If
                        End If
                
                'PRODUCCION POR FECHAS Y GRUPO DE LINEA
                ElseIf OptProduccion.Item(6).Value = True Then
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value & " Grupo " & TxtTextoProduccion.Text
                        
                        If OptOpcDet.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleO.rpt"
                              End If
                        ElseIf OptOpcRes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionResumenO.rpt"
                              End If
                        ElseIf OptOpcDetMatPri.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConMateriaPrima.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConMateriaPrimaO.rpt"
                              End If
                        ElseIf OptOpcDetDef.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                              End If
                        ElseIf OptDefCua.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectosCuadricula.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosCuadriculaO.rpt"
                              End If
                        
                        ElseIf OptFicTecMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaMesO.rpt"
                              End If
                        ElseIf OptFicTecAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaAñoO.rpt"
                              End If
                        ElseIf OptLinMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaMesO.rpt"
                              End If
                        ElseIf OptLinAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Linea} = {Lineas.Linea} And {Lineas.Grupo} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaAñoO.rpt"
                              End If
                        End If
                
                'PRODUCCION POR NUMERO DE ORDEN
                ElseIf OptProduccion.Item(4).Value = True Then
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value & " Numero De Orden " & TxtTextoProduccion.Text
                
                        If OptOpcDet.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Orden} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleO.rpt"
                              End If
                        ElseIf OptOpcRes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Orden} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionResumenO.rpt"
                              End If
                        ElseIf OptOpcDetMatPri.Value = True Then
                              'GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.CodigoMateriaPrima} = '" & TxtTextoProduccion.Text & "' And {ProduccionConMateriaPrima.Bulto} = " & TxtTextoProduccion2.Text
                              'GNombreReporte =  "\ProduccionDetalleConMateriaPrima.rpt"
                        ElseIf OptOpcDetDef.Value = True Then
                              'GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.CodigoMateriaPrima} = '" & TxtTextoProduccion.Text & "' And {ProduccionConDefectos.Bulto} = " & TxtTextoProduccion2.Text
                              'GNombreReporte =  "\ProduccionDetalleConDefectos.rpt"
                        ElseIf OptDefCua.Value = True Then
                        
                        
                        ElseIf OptFicTecMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Orden} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaMesO.rpt"
                              End If
                        ElseIf OptFicTecAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Orden} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaAñoO.rpt"
                              End If
                        ElseIf OptLinMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Orden} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaMesO.rpt"
                              End If
                        ElseIf OptLinAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Orden} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaAñoO.rpt"
                              End If
                        End If
                'PRODUCCION POR FECHAS Y FICHA TECNICA
                ElseIf OptProduccion.Item(5).Value = True Then
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value & " Ficha Tecnica " & TxtTextoProduccion.Text
                        
                        If OptOpcDet.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleO.rpt"
                              End If
                        ElseIf OptOpcRes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionResumenO.rpt"
                              End If
                        ElseIf OptOpcDetMatPri.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConMateriaPrima.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConMateriaPrimaO.rpt"
                              End If
                        ElseIf OptOpcDetDef.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                              End If
                        ElseIf OptDefCua.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                              End If
                        
                        ElseIf OptFicTecMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichatecnicaMesO.rpt"
                              End If
                        ElseIf OptFicTecAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaAñoO.rpt"
                              End If
                              
                        ElseIf OptLinMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaMesO.rpt"
                              End If
                        ElseIf OptLinAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Esp_Tec} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaAñoO.rpt"
                              End If
                        End If
                        
                'PRODUCCION POR FECHAS Y DESCRIPCION
                ElseIf OptProduccion.Item(3).Value = True Then
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value & " Descripcion " & TxtTextoProduccion.Text
                        
                        If OptOpcDet.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleO.rpt"
                              End If
                        ElseIf OptOpcRes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionResumenO.rpt"
                              End If
                        ElseIf OptOpcDetMatPri.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConMateriaPrima.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConMateriaPrimaO.rpt"
                              End If
                        ElseIf OptOpcDetDef.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosO.rpt"
                              End If
                              
                        ElseIf OptDefCua.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalleConDefectosCuadricula.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleConDefectosCuadriculaO.rpt"
                              End If
                        
                        ElseIf OptFicTecMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaMesO.rpt"
                              End If
                        ElseIf OptFicTecAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaAñoO.rpt"
                              End If
                        ElseIf OptLinMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLienaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaMesO.rpt"
                              End If
                              
                        ElseIf OptLinAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {FichaTecnica.Descrip} " & VTextoProduccion
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaAñoO.rpt"
                              End If
                        End If
                        
                        
                'PRODUCCION POR BATCH
                ElseIf OptProduccion.Item(7).Value = True Then
                        If Not IsNumeric(TxtTextoProduccion.Text) Then
                            MsgBox "El Batch Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                            Exit Sub
                        End If
                        
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value & " Numero De Batch " & TxtTextoProduccion.Text & " Linea " & LblDesPro2.Caption
                
                        If OptOpcDet.Value = True Then
                              'GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Batch} = " & TxtTextoProduccion.Text & " And {Produccion.Linea} = '" & TxtTextoProduccion2.Text & "'"
                              GCriteriaReporte = "{Produccion.Batch} = " & TxtTextoProduccion.Text & " And {Produccion.Linea} = '" & TxtTextoProduccion2.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionDetalleO.rpt"
                              End If
                        ElseIf OptOpcRes.Value = True Then
                              'GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.Batch} = " & TxtTextoProduccion.Text & " And {Produccion.Linea} = '" & TxtTextoProduccion2.Text & "'"
                              GCriteriaReporte = "{Produccion.Batch} = " & TxtTextoProduccion.Text & " And {Produccion.Linea} = '" & TxtTextoProduccion2.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionResumenO.rpt"
                              End If
                        ElseIf OptOpcDetMatPri.Value = True Then
                              'GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.CodigoMateriaPrima} = '" & TxtTextoProduccion.Text & "' And {ProduccionConMateriaPrima.Bulto} = " & TxtTextoProduccion2.Text
                              'GNombreReporte =  "\ProduccionDetalleConMateriaPrima.rpt"
                        ElseIf OptOpcDetDef.Value = True Then
                              'GCriteriaReporte = "{Produccion.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {Produccion.CodigoMateriaPrima} = '" & TxtTextoProduccion.Text & "' And {ProduccionConDefectos.Bulto} = " & TxtTextoProduccion2.Text
                              'GNombreReporte =  "\ProduccionDetalleConDefectos.rpt"
                        ElseIf OptDefCua.Value = True Then
                        
                        
                        ElseIf OptFicTecMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Batch} = " & TxtTextoProduccion.Text & " And {Produccion.Linea} = '" & TxtTextoProduccion2.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaMesO.rpt"
                              End If
                        ElseIf OptFicTecAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Batch} = " & TxtTextoProduccion.Text & " And {Produccion.Linea} = '" & TxtTextoProduccion2.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorFichaTecnicaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorFichaTecnicaAñoO.rpt"
                              End If
                        ElseIf OptLinMes.Value = True Then
                              GCriteriaReporte = "{Produccion.Batch} = " & TxtTextoProduccion.Text & " And {Produccion.Linea} = '" & TxtTextoProduccion2.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaMes.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaMesO.rpt"
                              End If
                              
                        ElseIf OptLinAño.Value = True Then
                              GCriteriaReporte = "{Produccion.Batch} = " & TxtTextoProduccion.Text & " And {Produccion.Linea} = '" & TxtTextoProduccion2.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionPorLineaAño.rpt"
                              Else
                                GNombreReporte = "ProduccionPorLineaAñoO.rpt"
                              End If
                        End If
                
                End If
                
                    'SI ES PRODUCTO LIBERADO NO PIDE CALIDAD CON LAS DEMAS OPCIONES SI COMPARA LA CALIDAD
                    'If (OptProduccion.Item(1).Value = True Or OptProduccion.Item(2).Value = True Or OptProduccion.Item(4).Value = True Or OptProduccion.Item(6).Value = True) Then
                         'SELECCIONA LA CALIDAD ACEPTADA, INCOMPLETA O COMPLEMENTO
                         If OptCalAIC.Value = True Then
                                     GCriteriaReporte = GCriteriaReporte & " And ({Produccion.Calidad} = 'A' OR {Produccion.Calidad} = 'I' Or {Produccion.Calidad} = 'C')"
                                     GTituloReporte = GTituloReporte & " Calidad Aceptada '"
                         'SELECCOINA LA CALIDAD RECHAZADA
                         ElseIf OptCalR.Value = True Then
                                     GCriteriaReporte = GCriteriaReporte & " And {Produccion.Calidad} = 'R'"
                                     GTituloReporte = GTituloReporte & " Calidad Retenida '"
                         'SELECCOINA LA CALIDAD INCOMPLETA
                         ElseIf OptCalI.Value = True Then
                                     GCriteriaReporte = GCriteriaReporte & " And {Produccion.Calidad} = 'I'"
                                     GTituloReporte = GTituloReporte & " Calidad Inspeccion '"
                         Else
                                    GTituloReporte = GTituloReporte & " Calidad Todas '"
                                    
                         End If
                    'End If
                         
                'CrReportes.DataFiles(0) = App.Path & "\mibase.mdb"
                FrmReporte.Show
                
                'BorrarFormulas CrReportes, 1
End If

'----------------------------------------------------------------------------------------------------------
'                                                  PRODUCCION LIBERADA
'----------------------------------------------------------------------------------------------------------


'PRODUCCION LIBERADA
If TabReportes.Tab = 1 Then
                
                 VDia = Day(DtpLibFecIni.Value)
                 VMes = Month(DtpLibFecIni.Value)
                 VAño = Year(DtpLibFecIni.Value)
                 VDia2 = Day(DtpLibFecFin.Value)
                 VMes2 = Month(DtpLibFecFin.Value)
                 VAño2 = Year(DtpLibFecFin.Value)
                    
                'PRODUCCION POR FECHAS
                If OptProduccionLiberado.Item(2).Value = True Then
                        
                        'CrReportes.WindowTitle = "Produccion desde " & PFecProIni.Value & " Hasta " & PFecProFin.Value
                        'CrReportes.BoundReportHeading = "Produccion desde "
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & DtpLibFecIni.Value & " Al " & DtpLibFecFin.Value & "'"
                        
                        If OptProLibOpc.Item(0).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberada.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(1).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberada.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaResumenO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(2).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberadaConTarimas.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalleConTarimas.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleConTarimasO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(3).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberadaConDefectos.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleConDefectosO.rpt"
                              End If
                        End If
                'PRODUCCION POR FECHAS Y LINEA
                ElseIf OptProduccionLiberado.Item(1).Value = True Then
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & DtpLibFecIni.Value & " Al " & DtpLibFecFin.Value & " Linea " & TxtLib.Text & "'"
                        
                        If OptProLibOpc.Item(0).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberada.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProduccionLiberada.Linea} = '" & TxtLib.Text & "'"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(1).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberada.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProduccionLiberada.Linea} = '" & TxtLib.Text & "'"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaResumenO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(2).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberadaConTarimas.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProduccionLiberadaConTarimas.Linea} = '" & TxtLib.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalleConTarimas.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleConTarimasO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(3).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberadaConDefectos.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProduccionLiberadaConDefectos.Linea} = '" & TxtLib.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleConDefectosO.rpt"
                              End If
                        End If
                'PRODUCCION POR FECHAS Y GRUPO DE LINEA
                ElseIf OptProduccionLiberado.Item(0).Value = True Then
                        GTituloReporte = Format(Date, "dd/mm/yyyy") & " " & Format(Time, "hh:mm:ss") & "     Reporte del " & PFecProIni.Value & " Al " & PFecProFin.Value & " Linea " & TxtTextoProduccion.Text & "'"
                        
                        If OptProLibOpc.Item(0).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberada.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProduccionLiberada.Linea} = {Lineas.Linea} And {Lineas.Grupo} = '" & TxtLib.Text & "'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalle.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(1).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberada.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProduccionLiberada.Linea} = {Lineas.Linea} And {Lineas.Grupo} = '" & TxtLib.Text & "'"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaResumen.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaResumenO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(2).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberadaConTarimas.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProduccionLiberadaConTarimas.Linea} = {Lineas.Linea} And {Lineas.Grupo} = '" & TxtLib.Text & "'"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalleConTarimas.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleConTarimasO.rpt"
                              End If
                        ElseIf OptProLibOpc.Item(3).Value = True Then
                            GCriteriaReporte = "{ProduccionLiberadaConDefectos.Fec_Prd} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {ProduccionLiberadaConDefectos.Linea} = {Lineas.Linea} And {Lineas.Grupo} = '" & TxtLib.Text & "'"
                            If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "ProduccionLiberadaDetalleConDefectos.rpt"
                              Else
                                GNombreReporte = "ProduccionLiberadaDetalleConDefectosO.rpt"
                              End If
                        End If
                End If
                
                FrmReporte.Show
                
                'BorrarFormulas CrReportes, 1
End If

'------------------------------------------------------------------------------------------------------------
'                                                    RUTINAS
'------------------------------------------------------------------------------------------------------------
If TabReportes.Tab = 2 Then
                Rutinas
End If ' FIN DE TAB
'------------------------------------------------------------------------------------------------------------
'                                                         BATCH
'------------------------------------------------------------------------------------------------------------

If TabReportes.Tab = 3 Then
                Batch
End If 'FIN DE TAB 2

'____________________________________________________________________________________________________________________
'                                        FICHA TECNICA
'____________________________________________________________________________________________________________________
If TabReportes.Tab = 4 Then
        'CODIGO DE FICHA TECNICA
        If OptFichaTecnica.Item(0).Value = True Then
                    GCriteriaReporte = "UPPERCASE({FichaTecnica.Esp_Tec})"
        'CODIGO DE CATALOGO
        ElseIf OptFichaTecnica.Item(3).Value = True Then
                    GCriteriaReporte = "UPPERCASE({FichaTecnica.Variables})"
        End If
        
                'IGUAL A
                If OptFichaTecnica2.Item(0).Value = True Then
                    GCriteriaReporte = GCriteriaReporte & " = '" & UCase(TxtFichaTecnica.Text) & "'"
                'PALABRA INICIAL
                ElseIf OptFichaTecnica2.Item(1).Value = True Then
                    GCriteriaReporte = GCriteriaReporte & " Like '" & UCase(TxtFichaTecnica.Text) & "*'"
                'CUALQUIER PALABRA
                ElseIf OptFichaTecnica2.Item(2).Value = True Then
                    GCriteriaReporte = GCriteriaReporte & " Like '*" & UCase(TxtFichaTecnica.Text) & "*'"
                End If
                    
                If OptFicTecTipRep.Item(0).Value = True Then
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "FichaTecnicaDetalle.rpt"
                              Else
                                GNombreReporte = "FichaTecnicaDetalleO.rpt"
                              End If
                ElseIf OptFicTecTipRep.Item(1).Value = True Then
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "FichaTecnicaResumen.rpt"
                              Else
                                GNombreReporte = "FichaTecnicaResumenO.rpt"
                              End If
                ElseIf OptFicTecTipRep.Item(2).Value = True Then
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "FichaTecnicaResumenGrupos.rpt"
                              Else
                                GNombreReporte = "FichaTecnicaResumenGruposO.rpt"
                              End If
                End If
                    
                FrmReporte.Show
                
End If

'--------------------------------------------------------------------------------------------------------------------
'                                                     VARIABLES
'--------------------------------------------------------------------------------------------------------------------
If TabReportes.Tab = 5 Then
    If OptVarCod.Value = True Then
        GCriteriaReporte = "UPPERCASE({VariablesMedia.Codigo}) LIKE '" & UCase(TxtVar.Text) & "*'"
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "Variables.rpt"
                              Else
                                GNombreReporte = "VariablesO.rpt"
                              End If
        FrmReporte.Show
        
    End If
End If

'--------------------------------------------------------------------------------------------------------------------
'                                                     DEFECTOS
'--------------------------------------------------------------------------------------------------------------------
If TabReportes.Tab = 6 Then
    If OptDefCod.Value = True Then
        GCriteriaReporte = "UPPERCASE({Defectos.Defecto}) LIKE '" & UCase(TxtDefectos.Text) & "*'"
    ElseIf OptDefTip.Value = True Then
        GCriteriaReporte = "UPPERCASE({Defectos.Tipo}) LIKE '" & UCase(TxtDefectos.Text) & "*'"
    End If
                              If GOrigenDeDatos = "AmaproAccess" Then
                                GNombreReporte = "Defectos.rpt"
                              Else
                                GNombreReporte = "DefectosO.rpt"
                              End If
        
        FrmReporte.Show
        
End If

        If Err <> 0 Then
            MousePointer = 0
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
            Exit Sub
        End If

MousePointer = 0
End Sub

Private Sub CmdSale_Click()
    FrameBuscar.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGridBuscar_DblClick()
        'PRODUCCION
        If TabReportes.Tab = 0 Then
                        TxtTextoProduccion.Text = DbGridBuscar.Columns(0)
                        TxtTextoProduccion.SetFocus
        'FICHA TECNICA
        ElseIf TabReportes.Tab = 4 Then
                TxtFichaTecnica.Text = DbGridBuscar.Columns(0)
                TxtFichaTecnica.SetFocus
        End If
        
        FrameBuscar.Visible = False
End Sub

Private Sub DBGridBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 43 Then
            'PRODUCCION
            If TabReportes.Tab = 0 Then
                        TxtTextoProduccion.Text = DbGridBuscar.Columns(0)
                        TxtTextoProduccion.SetFocus
            'FICHA TECNICA
            ElseIf TabReportes.Tab = 4 Then
                TxtFichaTecnica.Text = DbGridBuscar.Columns(0)
                TxtFichaTecnica.SetFocus
            End If
        End If
        FrameBuscar.Visible = False
            
End Sub

Private Sub Form_Load()
        PFecProIni.Value = Date
        PFecProFin.Value = Date
        
End Sub

Private Sub MskFecRut_GotFocus()
    MskFecRut.SelStart = 0
    MskFecRut.SelLength = Len(MskFecRut.Text)
End Sub

Private Sub OptBatNum_Click()
    'xtBatch.SetFocus
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
    If OptBusqueda.Item(0).Value = True Then
        LblBuscar.Caption = "Descripcion"
    ElseIf OptBusqueda.Item(1).Value = True Then
        LblBuscar.Caption = "Codigo"
    End If
        TxtBuscar.SetFocus
End Sub


Private Sub OptDefCod_Click()
        LblDefectos.Caption = "Codigo De Defecto"
        TxtDefectos.SetFocus
End Sub

Private Sub OptDefTip_Click()
        LblDefectos.Caption = "Tipo De Defecto"
        TxtDefectos.SetFocus
End Sub

Private Sub OptFichaTecnica_Click(Index As Integer)
        If Index = 0 Then
            LblFichaTecnica.Caption = "Codigo Ficha Tecnica"
        ElseIf Index = 3 Then
            LblFichaTecnica.Caption = "Codigo Catalogo"
        End If
            TxtFichaTecnica.SetFocus
            
End Sub

Private Sub OptFichaTecnica2_Click(Index As Integer)
        TxtFichaTecnica.SetFocus
End Sub



Private Sub OptProduccion_Click(Index As Integer)
        If OptProduccion.Item(0).Value = True Then
            LblProduccion.Caption = ""
        ElseIf OptProduccion.Item(1).Value = True Then
            LblProduccion.Caption = "Linea"
        ElseIf OptProduccion.Item(3).Value = True Then
            LblProduccion.Caption = "Descripcion"
        ElseIf OptProduccion.Item(4).Value = True Then
            LblProduccion.Caption = "Orden"
        ElseIf OptProduccion.Item(5).Value = True Then
            LblProduccion.Caption = "Ficha Tecnica"
        ElseIf OptProduccion.Item(6).Value = True Then
            LblProduccion.Caption = "Grupo"
        ElseIf OptProduccion.Item(7).Value = True Then
            LblProduccion.Caption = "Batch"
        End If
            
            'SI ELIGE LA OPCION DE ORDEN
            If (OptProduccion.Item(4).Value = True Or OptProduccion.Item(7).Value = True) Then
                OptOpcDet.Visible = True
                OptOpcRes.Visible = True
                OptOpcDetMatPri.Visible = False
                OptOpcDetDef.Visible = False
                
            'CUALQUIER OTRA OPCION
            Else
                OptOpcDet.Visible = True
                OptOpcRes.Visible = True
                OptOpcDetMatPri.Visible = True
                OptOpcDetDef.Visible = True
                
            End If
            
            
                'SI ES OPCION DE ORDEN
                If (OptProduccion.Item(4).Value = True Or OptProduccion.Item(6).Value = True) Then
                    OptOpcDetMatPri.Visible = False
                    OptOpcDetDef.Visible = False
                    
                Else
                    OptOpcDetMatPri.Visible = True
                    OptOpcDetDef.Visible = True
                    
                End If
            
                        
            'POR FECHAS Y PRODUCTO TERMINADO LIBERADO
            If OptProduccion.Item(0).Value = True Then
                TxtTextoProduccion.Visible = False
            Else
                TxtTextoProduccion.Visible = True
                TxtTextoProduccion.SetFocus
            End If
            
            If OptProduccion.Item(7).Value = True Then
                TxtTextoProduccion2.Visible = True
                LblProduccion2.Caption = "Linea"
            Else
                TxtTextoProduccion2.Visible = False
                LblProduccion2.Caption = ""
            End If
            
            'OPCION DE FECHAS LINEA Y TURNO
            If OptProduccion.Item(2).Value = True Then
                LblProduccion.Caption = "Linea"
                Lbltur.Visible = True
                TxtTur.Visible = True
                TxtTur.SetFocus
            Else
                Lbltur.Visible = False
                TxtTur.Visible = False
            End If
            
            
     
                        
            
End Sub


Private Sub OptProduccionLiberado_Click(Index As Integer)
        If Index = 0 Then
            LblLibEti.Caption = "Grupo De Linea"
            TxtLib.Visible = True
            TxtLib.SetFocus
        ElseIf Index = 1 Then
            LblLibEti.Caption = "Codigo De Linea"
            TxtLib.Visible = True
            TxtLib.SetFocus
        ElseIf Index = 2 Then
            LblLibEti.Caption = ""
            TxtLib.Visible = False
        End If
        
End Sub

Private Sub OptRutArr_Click()
        TxtLin.Visible = True
        MskFecRut.Visible = True
        TxtHorRut.Visible = True
        LblRutinas.Item(0).Visible = True
        LblRutinas.Item(1).Visible = True
        LblRutinas.Item(2).Visible = True
        TxtHorRut.SetFocus
        FrameRutinas.Visible = True
        MskFecFin.Visible = False
        LblHas.Visible = False
        LblRutinas.Item(2).Visible = True
End Sub

Private Sub OptRutDet_Click()
        TxtLin.Visible = True
        MskFecRut.Visible = True
        TxtHorRut.Visible = False
        LblRutinas.Item(0).Visible = True
        LblRutinas.Item(1).Visible = True
        LblRutinas.Item(2).Visible = True
        TxtLin.SetFocus
        FrameRutinas.Visible = True
        MskFecRut.Text = Date
        MskFecFin.Visible = True
        LblHas.Visible = True
        MskFecFin.Text = Date
        LblRutinas.Item(2).Visible = False
End Sub

Private Sub OptRutFicTec_Click()
        TxtLin.Visible = False
        MskFecRut.Visible = False
        TxtHorRut.Visible = False
        LblRutinas.Item(0).Visible = False
        LblRutinas.Item(1).Visible = False
        LblRutinas.Item(2).Visible = False
        FrameRutinas.Visible = False
        MskFecFin.Visible = False
        LblHas.Visible = False
        LblRutinas.Item(2).Visible = False
End Sub

Private Sub OptRutRep_Click()
        TxtLin.Visible = True
        MskFecRut.Visible = True
        TxtHorRut.Visible = True
        LblRutinas.Item(0).Visible = True
        LblRutinas.Item(1).Visible = True
        LblRutinas.Item(2).Visible = True
        TxtHorRut.SetFocus
        FrameRutinas.Visible = True
        MskFecFin.Visible = False
        LblHas.Visible = False
        LblRutinas.Item(2).Visible = True
End Sub


Private Sub OptVarCod_Click()
        TxtVar.SetFocus
End Sub

Private Sub TabReportes_Click(PreviousTab As Integer)
    'PRODUCCION
    If TabReportes.Tab = 0 Then
        OptProduccion.Item(0).Value = True
        PFecProIni.Value = Date
        PFecProFin.Value = Date
    ElseIf TabReportes.Tab = 1 Then
        OptProduccionLiberado.Item(0).Value = True
        DtpLibFecIni.Value = Date
        DtpLibFecFin.Value = Date
    'RUTINAS
    ElseIf TabReportes.Tab = 2 Then
        OptRutRep.Value = True
    'BATCH
    ElseIf TabReportes.Tab = 3 Then
        OptBatNum.Value = True
    'FICHAS TECNICAS
    ElseIf TabReportes.Tab = 4 Then
        OptFichaTecnica.Item(0).Value = True
    'CATALOGO
    ElseIf TabReportes.Tab = 5 Then
        OptVarCod.Value = True
    'DEFECTOS
    ElseIf TabReportes.Tab = 6 Then
        OptDefCod.Value = True
    End If
    
End Sub

Private Sub TxtBatLin_Change()
    Set RBuscaLinea = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas where Linea = '" & TxtBatLin.Text & "'")
        Else
            Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas where UPPER(Linea) = '" & UCase(TxtBatLin.Text) & "'")
        End If
        
            If RBuscaLinea.RecordCount > 0 Then
                    LblBatDesLin.Caption = RBuscaLinea!Descrip
            Else
                    LblBatDesLin.Caption = ""
            End If
End Sub

Private Sub Txtbuscar_Change()
            Set RBusqueda = New ADODB.Recordset
            'BUSCA LINEA
            If (OptProduccion.Item(1).Value = True) Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where Descrip Like '%" & TxtBuscar.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where Linea Like '%" & TxtBuscar.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas Where UPPER(Linea) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    End If
            End If

          
            'BUSCA FICHA TECNICA
            If (OptFichaTecnica.Item(0).Value = True Or OptProduccion.Item(5).Value = True) Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Descrip Like '%" & TxtBuscar.Text & "%' And Activa = -1")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' And Activa = -1")
                            End If
                    'OPCION DE CODIGO
                    Else
                            'OPCION DE CUALQUIER PALABRA
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Esp_Tec Like '%" & TxtBuscar.Text & "%' And Activa = -1")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscar.Text) & "%' And Activa = -1")
                            End If
                    End If
            End If
            
            
            'BUSCA VARIABLES
            If OptFichaTecnica.Item(3).Value = True Then
                    'OPCION POR DESCRIPCION
                    If OptBusqueda.Item(0).Value = True Then
                            'OPCION DE CUALQUIER PALABRA
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable from VariablesDescripcion Where DescripcionVariable Like '%" & TxtBuscar.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable from VariablesDescripcion Where UPPER(DescripcionVariable) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    'OPCION DE CODIGO
                    Else
                            'OPCION DE CUALQUIER PALABRA
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable from VariablesDescripcion Where CodigoVariable Like '%" & TxtBuscar.Text & "%'")
                            Else
                                Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable from VariablesDescripcion Where UPPER(CodigoVariable) Like '%" & UCase(TxtBuscar.Text) & "%'")
                            End If
                    End If
            End If
                        
            
            Set DbGridBuscar.DataSource = RBusqueda
            'ANCHO DE COLUMNAS
            If (TabReportes.Tab = 0 Or TabReportes.Tab = 4) Then
                DbGridBuscar.Columns(0).Width = "1500"
                DbGridBuscar.Columns(1).Width = "5000"
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

Private Sub TxtDefectos_GotFocus()
        TxtDefectos.SelStart = 0
        TxtDefectos.SelLength = Len(TxtDefectos.Text)
End Sub

Private Sub TxtFichaTecnica_Change()
        If OptFichaTecnica.Item(0).Value = True Then
                Set RBuscaFicha = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtFichaTecnica.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFichaTecnica.Text) & "'")
                    End If
                
                    If RBuscaFicha.RecordCount > 0 Then
                        LblDesFicTec.Caption = RBuscaFicha!Descrip
                    Else
                        LblDesFicTec.Caption = ""
                    End If
        'ElseIf OptFichaTecnica.Item(1).Value = True Then
        '        Set RBuscaHojalata = Db.OpenRecordset("Select Descrip From Platinas Where Platina = '" & TxtFichaTecnica.Text & "'")
                    'If RBuscaHojalata.RecordCount > 0 Then
                    '    LblDesFicTec.Caption = RBuscaHojalata!Descrip
                    'Else
                    '    LblDesFicTec.Caption = ""
                    'End If
        'ElseIf OptFichaTecnica.Item(2).Value = True Then
        '        Set RBuscaFondo = Db.OpenRecordset("Select Descrip From Fondos Where Fondo = '" & TxtFichaTecnica.Text & "'")
        '            If RBuscaFondo.RecordCount > 0 Then
        '                LblDesFicTec.Caption = RBuscaFondo!Descrip
        '            Else
        '                LblDesFicTec.Caption = ""
        '            End If
        ElseIf OptFichaTecnica.Item(3).Value = True Then
                Set RBuscaVariable = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaVariable, "Select DescripcionVariable From VariablesDescripcion Where CodigoVariable = '" & TxtFichaTecnica.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaVariable, "Select DescripcionVariable From VariablesDescripcion Where UPPER(CodigoVariable) = '" & UCase(TxtFichaTecnica.Text) & "'")
                    End If
                    
                    If RBuscaVariable.RecordCount > 0 Then
                        LblDesFicTec.Caption = RBuscaVariable!DescripcionVariable
                    Else
                        LblDesFicTec.Caption = ""
                    End If
        End If

End Sub

Private Sub TxtFichaTecnica_DblClick()
        
    'FICHA TECNICA
    If TabReportes.Tab = 4 Then
        
        FrameBuscar.Visible = True
        
        Set RBusqueda = New ADODB.Recordset
        'SI ES POR FICHA TECNICA
        If OptFichaTecnica.Item(0).Value = True Then
            'SELECCIONA LOS DATOS DE FICHAS TECNICAS
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
        'SI ES POR PLATINA
        'ElseIf OptFichaTecnica.Item(1).Value = True Then
        '    DataBuscar.RecordSource = "Select Platina, Descrip From Platinas"
        'SI ES POR FONDO
        'ElseIf OptFichaTecnica.Item(2).Value = True Then
        '    DataBuscar.RecordSource = "Select Fondo, Descrip From Fondos"
        'CATALOGO
        ElseIf OptFichaTecnica.Item(3).Value = True Then
            Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From VariablesDescripcion")
        End If
    End If 'FIN DE TAB
    
            
            Set DbGridBuscar.DataSource = RBusqueda
            TxtBuscar.SetFocus
            
            'ANCHO DE COLUMNAS
            If TabReportes.Tab = 4 Then
                DbGridBuscar.Columns(0).Width = "1500"
                DbGridBuscar.Columns(1).Width = "5000"
            End If

End Sub

Private Sub TxtFichaTecnica_KeyPress(KeyAscii As Integer)

    If KeyAscii = 43 Then
            'FICHA TECNICA
            If TabReportes.Tab = 4 Then
                
                FrameBuscar.Visible = True
                Set RBusqueda = New ADODB.Recordset
                'SI ES POR FICHA TECNICA
                If OptFichaTecnica.Item(0).Value = True Then
                    'SELECCIONA LOS DATOS DE FICHAS TECNICAS
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                'SI ES POR PLATINA
                'ElseIf OptFichaTecnica.Item(1).Value = True Then
                '    DataBuscar.RecordSource = "Select Platina, Descrip From Platinas"
                'SI ES POR FONDO
                'ElseIf OptFichaTecnica.Item(2).Value = True Then
                '    DataBuscar.RecordSource = "Select Fondo, Descrip From Fondos"
                'CATALOGO
                ElseIf OptFichaTecnica.Item(3).Value = True Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From VariablesDescripcion")
                End If
            End If 'FIN DE TAB
            
                    
                    Set DbGridBuscar.DataSource = RBusqueda
                    TxtBuscar.SetFocus
                    
                    'ANCHO DE COLUMNAS
                    If TabReportes.Tab = 4 Then
                        DbGridBuscar.Columns(0).Width = "1500"
                        DbGridBuscar.Columns(1).Width = "5000"
                    End If
    End If


End Sub


Private Sub TxtHorRut_GotFocus()
    TxtHorRut.SelStart = 0
    TxtHorRut.SelLength = Len(TxtHorRut.Text)
End Sub

Private Sub TxtLib_Change()
        'LINEA
        If OptProduccionLiberado.Item(1).Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLib.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLib.Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblLibDes.Caption = RBuscaLinea!Descrip
                Else
                    LblLibDes.Caption = ""
                End If
        End If
        
End Sub

Private Sub TxtLin_Change()
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas where Linea = '" & TxtLin.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
                    LblDesRut.Caption = RBuscaLinea!Descrip
            Else
                    LblDesRut.Caption = ""
            End If
End Sub

Private Sub TxtLin_GotFocus()
        TxtLin.SelStart = 0
        TxtLin.SelLength = Len(TxtLin.Text)
End Sub



Private Sub TxtTextoProduccion_Change()
        'LINEA
        If OptProduccion.Item(1).Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTextoProduccion.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTextoProduccion.Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblDesPro.Caption = RBuscaLinea!Descrip
                Else
                    LblDesPro.Caption = ""
                End If
        'FICHA TECNICA
        ElseIf OptProduccion.Item(5).Value = True Then
            Set RBuscaFicha = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTextoProduccion.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFicha, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTextoProduccion.Text) & "'")
                End If
                    If RBuscaFicha.RecordCount > 0 Then
                        LblDesPro.Caption = RBuscaFicha!Descrip
                    Else
                        LblDesPro.Caption = ""
                    End If
        Else
                LblDesPro.Caption = ""
        End If
End Sub

Private Sub TxtTextoProduccion_DblClick()
        
    'PRODUCCION
    If TabReportes.Tab = 0 Then
        Set RBusqueda = New ADODB.Recordset
        'SI ES POR LINEA
        If (OptProduccion.Item(1).Value = True Or OptProduccion.Item(6).Value = True) Then
            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas ")
            Set DbGridBuscar.DataSource = RBusqueda
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
        'SI ES POR FICHA TECNICA
        ElseIf OptProduccion.Item(5).Value = True Then
            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
            Set DbGridBuscar.DataSource = RBusqueda
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
        End If
            
            'ANCHO DE COLUMNAS
            If TabReportes.Tab = 0 Then
                DbGridBuscar.Columns(0).Width = "1600"
                DbGridBuscar.Columns(1).Width = "5000"
            End If
            
    End If 'FIN DE TAB
    
End Sub

Private Sub TxtTextoProduccion_GotFocus()
        TxtTextoProduccion.SelStart = 0
        TxtTextoProduccion.SelLength = Len(TxtTextoProduccion.Text)
End Sub

Private Sub TxtTextoProduccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 43 Then
            'PRODUCCION
            If TabReportes.Tab = 0 Then
                Set RBusqueda = New ADODB.Recordset
                'SI ES POR LINEA
                If (OptProduccion.Item(1).Value = True Or OptProduccion.Item(6).Value = True) Then
                    Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip, Grupo From Lineas")
                    Set DbGridBuscar.DataSource = RBusqueda
                    TxtBuscar.SetFocus
                    FrameBuscar.Visible = True
                
                'SI ES POR FICHA TECNICA
                ElseIf OptProduccion.Item(5).Value = True Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
                    Set DbGridBuscar.DataSource = RBusqueda
                    FrameBuscar.Visible = True
                    TxtBuscar.SetFocus
                End If
                    
                    'ANCHO DE COLUMNAS
                    If TabReportes.Tab = 0 Then
                        DbGridBuscar.Columns(0).Width = "1600"
                        DbGridBuscar.Columns(1).Width = "5000"
                    End If
                    
            End If 'FIN DE TAB
     End If

    
End Sub

Private Sub TxtTextoProduccion2_Change()
        'BATCH
        If OptProduccion.Item(7).Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtTextoProduccion2.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtTextoProduccion2.Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblDesPro2.Caption = RBuscaLinea!Descrip
                Else
                    LblDesPro2.Caption = ""
                End If
        End If
        
End Sub

Private Sub TxtTur_GotFocus()
                TxtTur.SelStart = 0
                TxtTur.SelLength = Len(TxtTur.Text)
End Sub

Private Sub TxtVar_Change()
        Set RBuscaVariable = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaVariable, "Select DescripcionVariable From VariablesDescripcion Where CodigoVariable = '" & TxtVar.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaVariable, "Select DescripcionVariable From VariablesDescripcion Where UPPER(CodigoVariable) = '" & UCase(TxtVar.Text) & "'")
            End If
                    If RBuscaVariable.RecordCount > 0 Then
                        LblDesCat.Caption = RBuscaVariable!DescripcionVariable
                    Else
                        LblDesCat.Caption = ""
                    End If
End Sub


Sub Rutinas()
                    
                    If OptRutFicTec.Value = True Then
                    
                    Else
                            If Not IsDate(MskFecRut.Text) Then
                                MsgBox "Fecha Invalida", vbOKOnly + vbInformation, "Informacion"
                                MousePointer = 0
                                Exit Sub
                            End If
                    End If
                       
                    Cont = 1
                    
                        'REPORTE DE DIMENSIONALES O ARRANQUE
                        If (OptRutRep.Value = True Or OptRutArr.Value) = True Then
                    
                                    VDia = Day(MskFecRut.Text)
                                    VMes = Month(MskFecRut.Text)
                                    VAño = Year(MskFecRut.Text)
                                    VDia2 = Day(MskFecRut.Text)
                                    VMes2 = Month(MskFecRut.Text)
                                    VAño2 = Year(MskFecRut.Text)
                                   
                                    GCriteriaReporte = "{CapturaRutinas.Fec_Rut} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {CapturaRutinas.Linea} = '" & TxtLin.Text & "' And {CapturaRutinas.Hor_Rut} = '" & TxtHorRut.Text & "'"
                                                                       
                                   'REGISTRO DE RUTINAS
                                    If OptRutRep.Value = True Then
                                            If OptPulgadas.Value = True Then
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    GNombreReporte = "\RutinasPulgadas.rpt"
                                                Else
                                                    GNombreReporte = "\RutinasPulgadasO.rpt"
                                                End If
                                            Else
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    GNombreReporte = "\RutinasMilimetros.rpt"
                                                Else
                                                    GNombreReporte = "\RutinasMilimetrosO.rpt"
                                                End If
                                            End If
                                    'REGISTRO DE ARRANQUE
                                    ElseIf OptRutArr.Value = True Then
                                            If OptPulgadas.Value = True Then
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    GNombreReporte = "\RutinasArranquePulgadas.rpt"
                                                Else
                                                    GNombreReporte = "\RutinasArranquePulgadasO.rpt"
                                                End If
                                                 
                                            Else
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    GNombreReporte = "\RutinasArranqueMilimetros.rpt"
                                                Else
                                                    GNombreReporte = "\RutinasArranqueMilimetrosO.rpt"
                                                End If
                                            End If
                                    End If
                                    
                                    
                        'RUTINAS DETALLE
                        ElseIf OptRutDet.Value = True Then
                                    VDia = Day(MskFecRut.Text)
                                    VMes = Month(MskFecRut.Text)
                                    VAño = Year(MskFecRut.Text)
                                    VDia2 = Day(MskFecFin.Text)
                                    VMes2 = Month(MskFecFin.Text)
                                    VAño2 = Year(MskFecFin.Text)
                                            GCriteriaReporte = "{CapturaRutinas.Fec_Rut} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And {CapturaRutinas.Linea} Like '" & TxtLin.Text & "*'"
                                                If OptPulgadas.Value = True Then
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            GNombreReporte = "\RutinasDetallePulgadas.rpt"
                                                        Else
                                                            GNombreReporte = "\RutinasDetallePulgadasO.rpt"
                                                        End If
                                                Else
                                                        If GOrigenDeDatos = "AmaproAccess" Then
                                                            GNombreReporte = "\RutinasDetalleMilimetros.rpt"
                                                        Else
                                                            GNombreReporte = "\RutinasDetalleMilimetrosO.rpt"
                                                        End If
                                                End If
                        'REPORTE DE FICHA TECNICA DE RUTINAS
                        ElseIf OptRutFicTec.Value = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        GNombreReporte = "\RutinasFichaTecnica.rpt"
                                    Else
                                        GNombreReporte = "\RutinasFichaTecnicaO.rpt"
                                    End If
                        End If ' OPCION DE RUTINAS DETALLE
                                    FrmReporte.Show
                    
                    
End Sub

Sub Batch()
On Error Resume Next

    GCriteriaReporte = ""
    
    If OptBatNum.Value = True Then
                                        
            If Not IsNumeric(TxtBatch.Text) Then
                MsgBox "Numero de Batch Incorrecto Tiene que ser Numerico", vbOKOnly + vbInformation, "Informacion"
                MousePointer = 0
                Exit Sub
            End If
            
            Conexion.Execute ("Delete from ReporteBatch")
            
                                                                
                        
                        'AGREGA ENCABEZADO DE PRODUCCION
                            VTexto = "' ', '" 'BATCH
                            VTexto = VTexto & " ', '" 'RUTINA
                            VTexto = VTexto & " ', '" 'LIM_PRO_IN
                            VTexto = VTexto & " ', '" 'LIM_PRO_SU
                            VTexto = VTexto & " ', '" 'CV
                            VTexto = VTexto & " ', '" 'LIM_ESP_IN
                            VTexto = VTexto & " ', '" 'LIM_ESP_SU
                            VTexto = VTexto & " ', '" 'CP
                            VTexto = VTexto & " ', '" 'DES_STD
                            VTexto = VTexto & " ', '" 'MEDIA
                            VTexto = VTexto & " ', '" 'DAT_MEN
                            VTexto = VTexto & " ', '" 'DAT_MAY
                            VTexto = VTexto & " ', '" 'FICHA TECNICA
                            VTexto = VTexto & " ', '" 'BATCH NUMERO
                            VTexto = VTexto & " '" 'CPK
                   
                            Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"
                            
                            If Err <> 0 Then
                                    MsgBox Err.Number & " " & Err.Description
                                    Err.Clear
                            End If '

                            VTexto = "'_____________', '" 'BATCH
                            VTexto = VTexto & "_____________', '" 'RUTINA
                            VTexto = VTexto & "_____________', '" 'LIM_PRO_IN
                            VTexto = VTexto & "_____________', '" 'LIM_PRO_SU
                            VTexto = VTexto & "_____________', '" 'CV
                            VTexto = VTexto & "_____________', '" 'LIM_ESP_IN
                            VTexto = VTexto & "_____________', '" 'LIM_ESP_SU
                            VTexto = VTexto & "_____________', '" 'CP
                            VTexto = VTexto & "_____________', '" 'DES_STD
                            VTexto = VTexto & "_____________', '" 'MEDIA
                            VTexto = VTexto & "_____________', '" 'DAT_MEN
                            VTexto = VTexto & "_____________', '" 'DAT_MAY
                            VTexto = VTexto & "_____________', '" 'FICHA TECNICA
                            VTexto = VTexto & "_____________', '" 'BATCH NUMERO
                            VTexto = VTexto & "_____________'" 'CPK
                    
                            Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"
                            
                            If Err <> 0 Then
                                    MsgBox Err.Number & " " & Err.Description
                                    Err.Clear
                            End If
                            
                            
                                                    
                'SELECCIONA LOS DATOS DE PRODUCCION
                    'TIPO DE REPORTE DE CLIENTE GUATEMALA
                    Set RBuscaBatch = New ADODB.Recordset
                    
                    If OptBatTipRep.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaBatch, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From Produccion Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtBatLin.Text & "' and Calidad = 'A' Order By Tarima")
                                Else
                                    Call Abrir_Recordset(RBuscaBatch, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From Produccion Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtBatLin.Text) & "' and UPPER(Calidad) = 'A' Order By Tarima")
                                End If
                    'CLIENTE MEXICO
                    ElseIf OptBatTipRep.Item(2).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaBatch, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From Produccion Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtBatLin.Text & "' and Calidad = 'A' Order By Tarima")
                                Else
                                    Call Abrir_Recordset(RBuscaBatch, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From Produccion Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtBatLin.Text) & "' and UPPER(Calidad) = 'A' Order By Tarima")
                                End If
                    'PARA LA EMPRESA
                    Else
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaBatch, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From Produccion Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtBatLin.Text & "' and Calidad <> 'C' Order By Tarima")
                                Else
                                    Call Abrir_Recordset(RBuscaBatch, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From Produccion Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtBatLin.Text) & "' and UPPER(Calidad) <> 'C' Order By Tarima")
                                End If
                    End If
                    
                    
                    
                    If RBuscaBatch.RecordCount > 0 Then
                                            
                        VFichaTecnica = RBuscaBatch!Esp_Tec
                        VBatch = RBuscaBatch!Batch
                                                    
                        'BUSCA EL NOMBRE COMERCIAL Y DESCRIPCION
                        Set RBuscaFicha = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaFicha, "Select Nombre_Comercial, Descrip, Imp_Defe From FichaTecnica Where Esp_tec = '" & VFichaTecnica & "'")
                            Else
                                Call Abrir_Recordset(RBuscaFicha, "Select Nombre_Comercial, Descrip, Imp_Defe From FichaTecnica Where UPPER(Esp_tec) = '" & UCase(VFichaTecnica) & "'")
                            End If
                            If RBuscaFicha.RecordCount > 0 Then
                                If IsNull(RBuscaFicha!Nombre_Comercial) Then
                                    VNombreComercial = ""
                                Else
                                    VNombreComercial = RBuscaFicha!Nombre_Comercial
                                End If
                                    VDescrip = RBuscaFicha!Descrip
                            Else
                                VNombreComercial = ""
                                VDescrip = ""
                            End If
                                            
                                                        
                        Do Until RBuscaBatch.EOF
                                                        
                                    
                                        VTexto = "'" & RBuscaBatch!Linea & "', '" 'BATCH
                                        VTexto = VTexto & Format(RBuscaBatch!fec_prd, "dd/mm/yyyy") & "', '" 'RUTINA
                                        VTexto = VTexto & RBuscaBatch!Calidad & "', '" 'LIM_PRO_IN
                                        VTexto = VTexto & "', '" 'LIM_PRO_SU
                                        VTexto = VTexto & "', '" 'CV
                                        
                                        If RBuscaFicha!Imp_Defe = -1 Then
                                                VTexto = VTexto & RBuscaBatch!Tarima & "', '" 'LIM_ESP_IN
                                                VTexto = VTexto & RBuscaBatch!Envases & "', '" 'LIM_ESP_SU
                                                VTexto = VTexto & "', '" 'CP
                                                VTexto = VTexto & "', '" 'DES_STD
                                                VTexto = VTexto & "', '" 'MEDIA
                                                VTexto = VTexto & "', '" 'DAT_MEN
                                                VTexto = VTexto & "', '" 'DAT_MAY
                                                VTexto = VTexto & "', '" 'FICHA TECNICA
                                                VTexto = VTexto & "', '" 'BATCH NUMERO
                                                VTexto = VTexto & "'" 'CPK
                                        Else
                                                VTexto = VTexto & RBuscaBatch!Tarima & "', '" 'LIM_ESP_IN
                                                VTexto = VTexto & RBuscaBatch!Envases & "', '" 'LIM_ESP_SU
                                                VTexto = VTexto & "', '" 'CP
                                                VTexto = VTexto & "', '" 'DES_STD
                                                VTexto = VTexto & "', '" 'MEDIA
                                                VTexto = VTexto & "', '" 'DAT_MEN
                                                VTexto = VTexto & "', '" 'DAT_MAY
                                                VTexto = VTexto & VFichaTecnica & "', '" 'FICHA TECNICA
                                                VTexto = VTexto & VBatch & "', '" 'BATCH NUMERO
                                                VTexto = VTexto & "'" 'CPK
                                        End If
                    
                    
                                        Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"
                                    
                                    
                                    If Err <> 0 Then
                                        MsgBox Err.Number & " " & Err.Description
                                        Err.Clear
                                    End If
                                    
                                RBuscaBatch.MoveNext
                        Loop
                    End If 'DATOS DE PRODUCCION
                                                                                
'__________________SELECCIONA LOS DATOS DE PRODUCCION DE LAS TARIMAS LIBERADAS_____________________________________
                        'TIPO DE REPORTE DE CLIENTE GUATEMALA
                        Set RBuscaBatch2 = New ADODB.Recordset
                        
                        If OptBatTipRep.Item(0).Value = True Then
                                
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaBatch2, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtBatLin.Text & "' and Calidad = 'A' Order By Tarima")
                                    Else
                                        Call Abrir_Recordset(RBuscaBatch2, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtBatLin.Text) & "' and UPPER(Calidad) = 'A' Order By Tarima")
                                    End If
                        'TIPO DE REPORTE DE CLIENTE MEXICO
                        ElseIf OptBatTipRep.Item(2).Value = True Then
                                
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaBatch2, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtBatLin.Text & "' and Calidad = 'A' Order By Tarima")
                                    Else
                                        Call Abrir_Recordset(RBuscaBatch2, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtBatLin.Text) & "' and UPPER(Calidad) = 'A' Order By Tarima")
                                    End If
                        'EMPRESA
                        Else
                                
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        Call Abrir_Recordset(RBuscaBatch2, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtBatLin.Text & "' Order By Tarima")
                                    Else
                                        Call Abrir_Recordset(RBuscaBatch2, "Select Linea, Fec_prd, Tarima, Envases, Esp_Tec, Batch, Calidad From ProduccionLiberada Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtBatLin.Text) & "' Order By Tarima")
                                    End If
                        End If
                        'DATOS DE PRODUCCION DE TARIMAS LIBERADAS
                        If RBuscaBatch2.RecordCount > 0 Then
                        
                            'TIPO DE REPORTE DE CLIENTE GUATEMALA Y CLIENTE DE MEXICO
                            If (OptBatTipRep.Item(0).Value = True Or OptBatTipRep.Item(2).Value = True) Then
                            Else
                                'AGREGA UNA LINEA
                                 
                                            VTexto = "'_____________', '" 'BATCH
                                            VTexto = VTexto & "_____________', '" 'RUTINA
                                            VTexto = VTexto & "_____________', '" 'LIM_PRO_IN
                                            VTexto = VTexto & "_____________', '" 'LIM_PRO_SU
                                            VTexto = VTexto & "_____________', '" 'CV
                                            VTexto = VTexto & "_____________', '" 'LIM_ESP_IN
                                            VTexto = VTexto & "_____________', '" 'LIM_ESP_SU
                                            VTexto = VTexto & "_____________', '" 'CP
                                            VTexto = VTexto & "Tarimas', '" 'DES_STD
                                            VTexto = VTexto & "Liberadas', '" 'MEDIA
                                            VTexto = VTexto & "_____________', '" 'DAT_MEN
                                            VTexto = VTexto & "_____________', '" 'DAT_MAY
                                            VTexto = VTexto & "_____________', '" 'FICHA TECNICA
                                            VTexto = VTexto & "_____________', '" 'BATCH NUMERO
                                            VTexto = VTexto & "_____________'" 'CPK

                                            Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"
                                 
                                            If Err <> 0 Then
                                                    MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                    Err.Clear
                                            End If
                                 
                            End If
                
                        
                                Do Until RBuscaBatch2.EOF
                                                    
                                    VFichaTecnica = RBuscaBatch2!Esp_Tec
                                    VBatch = RBuscaBatch2!Batch
                                                                    
                                    'BUSCA EL NOMBRE COMERCIAL Y DESCRIPCION
                                    Set RBuscaFicha = New ADODB.Recordset
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            Call Abrir_Recordset(RBuscaFicha, "Select Nombre_Comercial, Descrip From FichaTecnica Where Esp_tec = '" & VFichaTecnica & "'")
                                        Else
                                            Call Abrir_Recordset(RBuscaFicha, "Select Nombre_Comercial, Descrip From FichaTecnica Where UPPER(Esp_tec) = '" & UCase(VFichaTecnica) & "'")
                                        End If
                                        If RBuscaFicha.RecordCount > 0 Then
                                            If IsNull(RBuscaFicha!Nombre_Comercial) Then
                                                VNombreComercial = ""
                                            Else
                                                VNombreComercial = RBuscaFicha!Nombre_Comercial
                                            End If
                                                VDescrip = RBuscaFicha!Descrip
                                        Else
                                            VNombreComercial = ""
                                            VDescrip = ""
                                        End If
                                        
                                                VTexto = "'" & RBuscaBatch2!Linea & "', '" 'BATCH
                                                VTexto = VTexto & Format(RBuscaBatch2!fec_prd, "dd/mm/yyyy") & "', '" 'RUTINA
                                                VTexto = VTexto & RBuscaBatch2!Calidad & "', '" 'LIM_PRO_IN
                                                VTexto = VTexto & "', '"  'LIM_PRO_SU
                                                VTexto = VTexto & "', '" 'CV
                                                VTexto = VTexto & RBuscaBatch2!Tarima & "', '" 'LIM_ESP_IN
                                                VTexto = VTexto & RBuscaBatch2!Envases & "', '" 'LIM_ESP_SU
                                                VTexto = VTexto & "', '" 'CP
                                                VTexto = VTexto & "', '" 'DES_STD
                                                VTexto = VTexto & "', '" 'MEDIA
                                                VTexto = VTexto & "', '" 'DAT_MEN
                                                VTexto = VTexto & "', '" 'DAT_MAY
                                                VTexto = VTexto & VFichaTecnica & "', '" 'FICHA TECNICA
                                                VTexto = VTexto & VBatch & "', '" 'BATCH NUMERO
                                                VTexto = VTexto & "'" 'CPK
                                                
                                        Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"
                                        
                                        If Err <> 0 Then
                                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                Err.Clear
                                        End If
                                    
                                    
                                    RBuscaBatch2.MoveNext
                                Loop
                        End If
                            
                            VTexto = "'', '" 'BATCH
                            VTexto = VTexto & "', '" 'RUTINA
                            VTexto = VTexto & "', '" 'LIM_PRO_IN
                            VTexto = VTexto & "', '" 'LIM_PRO_SU
                            VTexto = VTexto & "', '" 'CV
                            VTexto = VTexto & "', '" 'LIM_ESP_IN
                            VTexto = VTexto & "', '" 'LIM_ESP_SU
                            VTexto = VTexto & "', '" 'CP
                            VTexto = VTexto & "', '" 'DES_STD
                            VTexto = VTexto & "', '" 'MEDIA
                            VTexto = VTexto & "', '" 'DAT_MEN
                            VTexto = VTexto & "', '" 'DAT_MAY
                            VTexto = VTexto & "', '" 'FICHA TECNICA
                            VTexto = VTexto & "', '" 'BATCH NUMERO
                            VTexto = VTexto & "'" 'CPK
                    
                            Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"
                            
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Err.Clear
                            End If
                                                                                   
                
                'TIPO DE REPORTE DE CLIENTE GUATEMALA
                    If OptBatTipRep.Item(0).Value = True Then
                            VTexto = "'Batch', '" 'BATCH
                            VTexto = VTexto & "Rutina', '" 'RUTINA
                            VTexto = VTexto & "LEI', '" 'LIM_PRO_IN
                            VTexto = VTexto & "LES', '" 'LIM_PRO_SU
                            VTexto = VTexto & "MEDIA', '" 'CV
                            VTexto = VTexto & "', '" 'LIM_ESP_IN
                            VTexto = VTexto & "', '" 'LIM_ESP_SU
                            VTexto = VTexto & "', '" 'CP
                            VTexto = VTexto & "', '" 'DES_STD
                            VTexto = VTexto & "', '" 'MEDIA
                            VTexto = VTexto & "', '" 'DAT_MEN
                            VTexto = VTexto & "', '" 'DAT_MAY
                            VTexto = VTexto & "', '" 'FICHA TECNICA
                            VTexto = VTexto & "', '" 'BATCH NUMERO
                            VTexto = VTexto & "'" 'CPK
                    'PARA LA EMPRESA Y PARA EL CLIENTE DE MEXICO
                    Else
                            VTexto = "'Batch', '" 'BATCH
                            VTexto = VTexto & "Rutina', '" 'RUTINA
                            VTexto = VTexto & "LPI', '" 'LIM_PRO_IN
                            VTexto = VTexto & "LPS', '" 'LIM_PRO_SU
                            VTexto = VTexto & "CV%', '" 'CV
                            VTexto = VTexto & "LEI', '" 'LIM_ESP_IN
                            VTexto = VTexto & "LES', '" 'LIM_ESP_SU
                            VTexto = VTexto & "CP', '" 'CP
                            VTexto = VTexto & "Desv. Std.', '" 'DES_STD
                            VTexto = VTexto & "Media', '" 'MEDIA
                            VTexto = VTexto & "Dato Menor', '" 'DAT_MEN
                            VTexto = VTexto & "Dato Mayor', '" 'DAT_MAY
                            VTexto = VTexto & "', '" 'FICHA TECNICA
                            VTexto = VTexto & "', '" 'BATCH NUMERO
                            VTexto = VTexto & "CPK '" 'CPK
                    End If
                            Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"
                    
                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Err.Clear
                            End If
                
                    'TIPO DE REPORTE DE CLIENTE GUATEMALA
                    If OptBatTipRep.Item(0).Value = True Then
                            VTexto = "'_____________', '" 'BATCH
                            VTexto = VTexto & "_____________', '" 'RUTINA
                            VTexto = VTexto & "_____________', '" 'LIM_PRO_IN
                            VTexto = VTexto & "_____________', '" 'LIM_PRO_SU
                            VTexto = VTexto & "_____________', '" 'CV
                            VTexto = VTexto & "', '" 'LIM_ESP_IN
                            VTexto = VTexto & "', '" 'LIM_ESP_SU
                            VTexto = VTexto & "', '" 'CP
                            VTexto = VTexto & "', '" 'DES_STD
                            VTexto = VTexto & "', '" 'MEDIA
                            VTexto = VTexto & "', '" 'DAT_MEN
                            VTexto = VTexto & "', '" 'DAT_MAY
                            VTexto = VTexto & "_____________', '" 'FICHA TECNICA
                            VTexto = VTexto & "_____________', '" 'BATCH NUMERO
                            VTexto = VTexto & "'" 'CPK
                    'PARA LA EMPRESA Y PARA EL CLIENTE DE MEXICO
                    Else
                            VTexto = "'_____________', '" 'BATCH
                            VTexto = VTexto & "_____________', '" 'RUTINA
                            VTexto = VTexto & "_____________', '" 'LIM_PRO_IN
                            VTexto = VTexto & "_____________', '" 'LIM_PRO_SU
                            VTexto = VTexto & "_____________', '" 'CV
                            VTexto = VTexto & "_____________', '" 'LIM_ESP_IN
                            VTexto = VTexto & "_____________', '" 'LIM_ESP_SU
                            VTexto = VTexto & "_____________', '" 'CP
                            VTexto = VTexto & "_____________', '" 'DES_STD
                            VTexto = VTexto & "_____________', '" 'MEDIA
                            VTexto = VTexto & "_____________', '" 'DAT_MEN
                            VTexto = VTexto & "_____________', '" 'DAT_MAY
                            VTexto = VTexto & "_____________', '" 'FICHA TECNICA
                            VTexto = VTexto & "_____________', '" 'BATCH NUMERO
                            VTexto = VTexto & "_____________'" 'CPK
                    End If
                    
                            Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"

                            If Err <> 0 Then
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                Err.Clear
                            End If
                                            
                                                    
                'BUSCA LOS DATOS DEL BATCH
                Set RBuscaBatchDatos = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBuscaBatchDatos, "Select * From BatchDatos Where Batch = " & TxtBatch.Text & " And Linea = '" & TxtBatLin.Text & "' Order By Rutina")
                        Else
                            Call Abrir_Recordset(RBuscaBatchDatos, "Select * From BatchDatos Where Batch = " & TxtBatch.Text & " And UPPER(Linea) = '" & UCase(TxtBatLin.Text) & "' Order By Rutina")
                        End If
                                        
                If RBuscaBatchDatos.RecordCount > 0 Then
                    Do Until RBuscaBatchDatos.EOF
                        
                        'RReporteBatch.AddNew
                            'TIPO DE REPORTE DE CLIENTE DE GUATEMALA
                            If OptBatTipRep.Item(0).Value = True Then
                            
                                    VTexto = "'" & TxtBatch.Text & "', '" 'BATCH
                                    VTexto = VTexto & RBuscaBatchDatos(1) & "', '" 'RUTINA
                                    'PULGADAS
                                    If OptBatPulgadas.Value = True Then
                                        VTexto = VTexto & Format(RBuscaBatchDatos(5) / 25.4, "#,###,##0.00") & "', '"
                                        VTexto = VTexto & Format(RBuscaBatchDatos(6) / 25.4, "#,###,##0.00") & "', '"
                                        VTexto = VTexto & Format(RBuscaBatchDatos(9) / 25.4, "#,###,##0.00") & "', '"
                                    'MILIMETROS
                                    Else
                                        VTexto = VTexto & Format(RBuscaBatchDatos(5), "#,###,##0.00") & "', '"
                                        VTexto = VTexto & Format(RBuscaBatchDatos(6), "#,###,##0.00") & "', '"
                                        VTexto = VTexto & Format(RBuscaBatchDatos(9), "#,###,##0.00") & "', '"
                                    End If
                                    
                                        VTexto = VTexto & "', '" 'LIM_ESP_IN
                                        VTexto = VTexto & "', '" 'LIM_ESP_SU
                                        VTexto = VTexto & "', '" 'CP
                                        VTexto = VTexto & "', '" 'DES_STD
                                        VTexto = VTexto & "', '" 'MEDIA
                                        VTexto = VTexto & "', '" 'DAT_MEN
                                        VTexto = VTexto & "', '" 'DAT_MAY
                                        VTexto = VTexto & VFichaTecnica & "', '" 'FICHA TECNICA
                                        VTexto = VTexto & VBatch & "', '" 'BATCH NUMERO
                                        VTexto = VTexto & "'" 'CPK
                                    
                            ''PARA LA EMPRESA Y PARA EL CLIENTE DE MEXICO
                            Else
                                'SI PIDE EL REPORTE EN PULGADAS
                                If OptBatPulgadas.Value = True Then
                                        VTexto = "'" & TxtBatch.Text & "', '" 'BATCH
                                        VTexto = VTexto & RBuscaBatchDatos(1) & "', '" 'RUTINA
                                        VTexto = VTexto & Format(RBuscaBatchDatos(2) / 25.4, "#,###,##0.00") & "', '" 'LIM_PRO_IN
                                        VTexto = VTexto & Format(RBuscaBatchDatos(3) / 25.4, "#,###,##0.00") & "', '" 'LIM_PRO_SU
                                        VTexto = VTexto & Format(RBuscaBatchDatos(4) / 25.4, "#,###,##0.00") & "', '" 'CV
                                        VTexto = VTexto & Format(RBuscaBatchDatos(5) / 25.4, "#,###,##0.00") & "', '" 'LIM_ESP_IN
                                        VTexto = VTexto & Format(RBuscaBatchDatos(6) / 25.4, "#,###,##0.00") & "', '" 'LIM_ESP_SU
                                        VTexto = VTexto & Format(RBuscaBatchDatos(7) / 25.4, "#,###,##0.00") & "', '" 'CP
                                        VTexto = VTexto & Format(RBuscaBatchDatos(8) / 25.4, "#,###,##0.00") & "', '" 'DES_STD
                                        VTexto = VTexto & Format(RBuscaBatchDatos(9) / 25.4, "#,###,##0.00") & "', '" 'MEDIA
                                        VTexto = VTexto & Format(RBuscaBatchDatos(10) / 25.4, "#,###,##0.00") & "', '" 'DAT_MEN
                                        VTexto = VTexto & Format(RBuscaBatchDatos(11) / 25.4, "#,###,##0.00") & "', '" 'DAT_MAY
                                        VTexto = VTexto & VFichaTecnica & "', '" 'FICHA TECNICA
                                        VTexto = VTexto & VBatch & "', '" 'BATCH NUMERO
                                        VTexto = VTexto & Format(RBuscaBatchDatos(13) / 25.4, "#,###,##0.0000") & "'" 'CPK
                                        
                                'SI ES EN MILIMETROS
                                Else
                                        VTexto = "'" & TxtBatch.Text & "', '" 'BATCH
                                        VTexto = VTexto & RBuscaBatchDatos(1) & "', '" 'RUTINA
                                        VTexto = VTexto & Format(RBuscaBatchDatos(2), "#,###,##0.00") & "', '" 'LIM_PRO_IN
                                        VTexto = VTexto & Format(RBuscaBatchDatos(3), "#,###,##0.00") & "', '" 'LIM_PRO_SU
                                        VTexto = VTexto & Format(RBuscaBatchDatos(4), "#,###,##0.00") & "', '" 'CV
                                        VTexto = VTexto & Format(RBuscaBatchDatos(5), "#,###,##0.00") & "', '" 'LIM_ESP_IN
                                        VTexto = VTexto & Format(RBuscaBatchDatos(6), "#,###,##0.00") & "', '" 'LIM_ESP_SU
                                        VTexto = VTexto & Format(RBuscaBatchDatos(7), "#,###,##0.00") & "', '" 'CP
                                        VTexto = VTexto & Format(RBuscaBatchDatos(8), "#,###,##0.00") & "', '" 'DES_STD
                                        VTexto = VTexto & Format(RBuscaBatchDatos(9), "#,###,##0.00") & "', '" 'MEDIA
                                        VTexto = VTexto & Format(RBuscaBatchDatos(10), "#,###,##0.00") & "', '" 'DAT_MEN
                                        VTexto = VTexto & Format(RBuscaBatchDatos(11), "#,###,##0.00") & "', '" 'DAT_MAY
                                        VTexto = VTexto & VFichaTecnica & "', '" 'FICHA TECNICA
                                        VTexto = VTexto & VBatch & "', '" 'BATCH NUMERO
                                        VTexto = VTexto & Format(RBuscaBatchDatos(13) / 25.4, "#,###,##0.0000") & "'" 'CPK
                                End If
                            End If
                            
                                        Conexion.Execute "Insert Into ReporteBatch Values(" & VTexto & ")"

                                        If Err <> 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                            Err.Clear
                                        End If
                         
                        RBuscaBatchDatos.MoveNext
                    Loop
                Else
                       
                End If
                
                                            
                
                GTituloReporte = ""
                GComentarioReporte = ""
                
                'SI PIDE EL REPORTE EN PULGADAS
                If OptBatPulgadas.Value = True Then
                    GTituloReporte = Now & "          Unidad De Medida Pulgadas "
                Else
                    GTituloReporte = Now & "          Unidad De Medida Milimetros "
                End If
                GTituloReporte = GTituloReporte & "       FichaTecnica = " & VFichaTecnica
                GTituloReporte = GTituloReporte & "       Descripcion = " & VDescrip
                
                GComentarioReporte = Left(VNombreComercial & Space(12), 12)
                GComentarioReporte = GComentarioReporte & Space(20) & "Batch = " & VBatch
                
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "\Batch.rpt"
                Else
                    GNombreReporte = "\BatchO.rpt"
                End If
                FrmReporte.Show
                                            
                
                            
    End If ' OPCION DE RUTINAS POR REPORTE
            
            
End Sub

