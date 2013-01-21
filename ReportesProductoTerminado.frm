VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ReportesProductoTerminado 
   BackColor       =   &H00FF8080&
   Caption         =   "Reportes De Inventario"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "ReportesProductoTerminado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   11880
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
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6615
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   11668
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
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         Height          =   195
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Descripcion"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         ToolTipText     =   "Digite sus Datos Para Buscar"
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton CmdSale 
         Height          =   735
         Left            =   10680
         Picture         =   "ReportesProductoTerminado.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Sale De Busqueda"
         Top             =   240
         Width           =   855
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
         TabIndex        =   8
         Top             =   720
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
      Left            =   10200
      Picture         =   "ReportesProductoTerminado.frx":24B4
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "&Vista Preliminar"
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
      Left            =   10200
      Picture         =   "ReportesProductoTerminado.frx":2DE6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin TabDlg.SSTab TabReportes 
      Height          =   7815
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   13785
      _Version        =   393216
      Tabs            =   6
      TabHeight       =   1058
      BackColor       =   16744576
      TabCaption(0)   =   "Inventario"
      TabPicture(0)   =   "ReportesProductoTerminado.frx":3530
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblInvLin2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblInvLin"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblInvOpc"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblInv"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LblInv2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LblInvDesCod"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "LblInvFecIni"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "LblInvFecFin"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Frame2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FrameInvPro"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "OptInv(6)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "FrameTipExi"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TxtInvLin"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "OptInv(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "OptInv(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtInvOpc"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtInv"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "FrameInvResDet"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "FrameInvOpc"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "FrameInvTipBus"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "OptInv(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "OptInv(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "DtpInvFecIni"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "DtpInvFecFin"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).ControlCount=   25
      TabCaption(1)   =   "Entradas"
      TabPicture(1)   =   "ReportesProductoTerminado.frx":384A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame3"
      Tab(1).Control(1)=   "OptEntradas(4)"
      Tab(1).Control(2)=   "OptEntradas(3)"
      Tab(1).Control(3)=   "TxtEntradas"
      Tab(1).Control(4)=   "OptEntradas(2)"
      Tab(1).Control(5)=   "OptEntradas(1)"
      Tab(1).Control(6)=   "OptEntradas(0)"
      Tab(1).Control(7)=   "FrameEntradas"
      Tab(1).Control(8)=   "OptEntradas(5)"
      Tab(1).Control(9)=   "TxtEntradas2"
      Tab(1).Control(10)=   "DtpEntFecFin"
      Tab(1).Control(11)=   "DtpEntFecIni"
      Tab(1).Control(12)=   "LblEntFecFin"
      Tab(1).Control(13)=   "LblEntFecIni"
      Tab(1).Control(14)=   "LblEntEti"
      Tab(1).Control(15)=   "LblEntDes"
      Tab(1).Control(16)=   "LblEntDes2"
      Tab(1).Control(17)=   "LblEntEti2"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Salidas"
      TabPicture(2)   =   "ReportesProductoTerminado.frx":3B64
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).Control(1)=   "OptSalidas(0)"
      Tab(2).Control(2)=   "OptSalidas(1)"
      Tab(2).Control(3)=   "OptSalidas(3)"
      Tab(2).Control(4)=   "OptSalidas(4)"
      Tab(2).Control(5)=   "TxtSalidas"
      Tab(2).Control(6)=   "OptSalidas(6)"
      Tab(2).Control(7)=   "OptSalidas(5)"
      Tab(2).Control(8)=   "FrameDespachos"
      Tab(2).Control(9)=   "OptSalidas(2)"
      Tab(2).Control(10)=   "TxtSalidas2"
      Tab(2).Control(11)=   "OptSalidas(7)"
      Tab(2).Control(12)=   "OptSalidas(8)"
      Tab(2).Control(13)=   "DtpSalFecFin"
      Tab(2).Control(14)=   "DtpSalFecIni"
      Tab(2).Control(15)=   "LblSalEti"
      Tab(2).Control(16)=   "LblSalDes"
      Tab(2).Control(17)=   "LblSalFecIni"
      Tab(2).Control(18)=   "LblSalFecFin"
      Tab(2).Control(19)=   "LblSalEti2"
      Tab(2).Control(20)=   "LblSalDes2"
      Tab(2).ControlCount=   21
      TabCaption(3)   =   "Cierre Bulto/Tarima"
      TabPicture(3)   =   "ReportesProductoTerminado.frx":443E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "OptCierreTarimas(4)"
      Tab(3).Control(1)=   "OptCierreTarimas(3)"
      Tab(3).Control(2)=   "OptCierreTarimas(0)"
      Tab(3).Control(3)=   "OptCierreTarimas(1)"
      Tab(3).Control(4)=   "TxtCieTar"
      Tab(3).Control(5)=   "OptCierreTarimas(2)"
      Tab(3).Control(6)=   "FrameCieTarTipRep"
      Tab(3).Control(7)=   "DtpCieTarFecFin"
      Tab(3).Control(8)=   "DtpCieTarFecIni"
      Tab(3).Control(9)=   "LblCieTarFecIni"
      Tab(3).Control(10)=   "LblCieTarFecFin"
      Tab(3).Control(11)=   "LblCieTarEti"
      Tab(3).Control(12)=   "LblCieTarDes"
      Tab(3).ControlCount=   13
      TabCaption(4)   =   "Desperdicio"
      TabPicture(4)   =   "ReportesProductoTerminado.frx":4758
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "TxtDesperdicio"
      Tab(4).Control(1)=   "FrameDesperdicio"
      Tab(4).Control(2)=   "OptDesperdicio(3)"
      Tab(4).Control(3)=   "OptDesperdicio(2)"
      Tab(4).Control(4)=   "OptDesperdicio(1)"
      Tab(4).Control(5)=   "OptDesperdicio(0)"
      Tab(4).Control(6)=   "DTPDesFecFin"
      Tab(4).Control(7)=   "DTPDesFecIni"
      Tab(4).Control(8)=   "LblDesFecIni"
      Tab(4).Control(9)=   "LblDesFecFin"
      Tab(4).Control(10)=   "LblDesEti"
      Tab(4).Control(11)=   "LblDesDes"
      Tab(4).Control(12)=   "LblCerBulDes"
      Tab(4).Control(13)=   "LblCerBulEti"
      Tab(4).ControlCount=   14
      TabCaption(5)   =   "Traslados"
      TabPicture(5)   =   "ReportesProductoTerminado.frx":9BF2
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "TxtTraslados"
      Tab(5).Control(1)=   "TxtTraslados2"
      Tab(5).Control(2)=   "TxtTraTipDoc"
      Tab(5).Control(3)=   "FrameTraslados"
      Tab(5).Control(4)=   "FrameTraslados2"
      Tab(5).Control(5)=   "OptTraslados(7)"
      Tab(5).Control(6)=   "OptTraslados(6)"
      Tab(5).Control(7)=   "OptTraslados(3)"
      Tab(5).Control(8)=   "OptTraslados(2)"
      Tab(5).Control(9)=   "OptTraslados(1)"
      Tab(5).Control(10)=   "OptTraslados(0)"
      Tab(5).Control(11)=   "OptTraslados(8)"
      Tab(5).Control(12)=   "OptTraslados(9)"
      Tab(5).Control(13)=   "OptTraslados(4)"
      Tab(5).Control(14)=   "OptTraslados(5)"
      Tab(5).Control(15)=   "DtpTraFecFin"
      Tab(5).Control(16)=   "DtpTraFecIni"
      Tab(5).Control(17)=   "LblTraslados"
      Tab(5).Control(18)=   "LblLabel(1)"
      Tab(5).Control(19)=   "LblLabel(0)"
      Tab(5).Control(20)=   "LblTraBod"
      Tab(5).Control(21)=   "LblTraBod2"
      Tab(5).Control(22)=   "LblTraslados2"
      Tab(5).Control(23)=   "LblTraEtiDoc"
      Tab(5).Control(24)=   "LblTraDesDoc"
      Tab(5).Control(25)=   "LblTraDes2"
      Tab(5).Control(26)=   "LblTraDes"
      Tab(5).Control(27)=   "LblTraEti"
      Tab(5).ControlCount=   28
      Begin VB.OptionButton OptCierreTarimas 
         Caption         =   "Fechas y Tipo "
         Height          =   195
         Index           =   4
         Left            =   -74640
         TabIndex        =   178
         Top             =   3000
         Width           =   2175
      End
      Begin VB.OptionButton OptCierreTarimas 
         Caption         =   "Fechas y Descripcion"
         Height          =   195
         Index           =   3
         Left            =   -74640
         TabIndex        =   177
         Top             =   2640
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo De Antiguedad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   1335
         Left            =   6960
         TabIndex        =   173
         Top             =   5640
         Width           =   2775
         Begin VB.OptionButton OptInvFec 
            Caption         =   "Sin Fechas"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   175
            Top             =   840
            Value           =   -1  'True
            Width           =   1935
         End
         Begin VB.OptionButton OptInvFec 
            Caption         =   "Fechas De Produccion/Entrada"
            Height          =   495
            Index           =   0
            Left            =   120
            TabIndex        =   174
            Top             =   240
            Width           =   1935
         End
      End
      Begin MSComCtl2.DTPicker DtpInvFecFin 
         Height          =   285
         Left            =   3840
         TabIndex        =   170
         Top             =   6240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyyy"
         Format          =   89980931
         CurrentDate     =   38478
      End
      Begin MSComCtl2.DTPicker DtpInvFecIni 
         Height          =   285
         Left            =   3840
         TabIndex        =   169
         Top             =   5880
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   38478
      End
      Begin VB.Frame Frame1 
         Caption         =   "Tipo De Inventario"
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
         Height          =   1335
         Left            =   -70440
         TabIndex        =   163
         Top             =   1560
         Width           =   2055
         Begin VB.OptionButton OptSalTipInv 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   166
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton OptSalTipInv 
            Caption         =   "Materia Prima"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   165
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton OptSalTipInv 
            Caption         =   "Producto Terminado"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   164
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo De Inventario"
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
         Height          =   1335
         Left            =   -70440
         TabIndex        =   159
         Top             =   1560
         Width           =   2055
         Begin VB.OptionButton OptEntTipInv 
            Caption         =   "Producto Terminado"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   162
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
         Begin VB.OptionButton OptEntTipInv 
            Caption         =   "Materia Prima"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   161
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton OptEntTipInv 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   160
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.TextBox TxtTraslados 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71280
         TabIndex        =   148
         ToolTipText     =   "Signo '+' o Doble Click para Ayuda"
         Top             =   5160
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtTraslados2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71280
         TabIndex        =   147
         Top             =   5520
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox TxtTraTipDoc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71280
         TabIndex        =   146
         Top             =   5880
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame FrameTraslados 
         Caption         =   "Tipo De Documento"
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
         Left            =   -68160
         TabIndex        =   143
         Top             =   1440
         Width           =   2175
         Begin VB.OptionButton OptTraOpc 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   145
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptTraOpc 
            Caption         =   "Un Tipo Documento"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   144
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame FrameTraslados2 
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
         ForeColor       =   &H000000FF&
         Height          =   1095
         Left            =   -70080
         TabIndex        =   140
         Top             =   1440
         Width           =   1815
         Begin VB.OptionButton OptTraDet 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   240
            TabIndex        =   142
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptTraRes 
            Caption         =   "Resumen"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   240
            TabIndex        =   141
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Orden"
         Height          =   195
         Index           =   7
         Left            =   -74520
         TabIndex        =   139
         Top             =   4320
         Width           =   1695
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Ficha Tecnica"
         Height          =   195
         Index           =   6
         Left            =   -74520
         TabIndex        =   138
         Top             =   3600
         Width           =   2655
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Bodega Entrada y Ficha Tecnica"
         Height          =   195
         Index           =   3
         Left            =   -74520
         TabIndex        =   137
         Top             =   2520
         Width           =   4095
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Bodega Salida y Ficha Tecnica"
         Height          =   195
         Index           =   2
         Left            =   -74520
         TabIndex        =   136
         Top             =   2160
         Width           =   3855
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Numero Documento"
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   135
         Top             =   1800
         Width           =   2295
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74520
         TabIndex        =   134
         Top             =   1440
         Width           =   975
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Ficha Tecnica"
         Height          =   195
         Index           =   8
         Left            =   -74520
         TabIndex        =   133
         Top             =   3960
         Width           =   2655
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Traslados No Liberados"
         Height          =   195
         Index           =   9
         Left            =   -74520
         TabIndex        =   132
         Top             =   4680
         Width           =   2655
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Bodega Salida Y Tipo Ficha Tecnica"
         Height          =   195
         Index           =   4
         Left            =   -74520
         TabIndex        =   131
         Top             =   2880
         Width           =   3855
      End
      Begin VB.OptionButton OptTraslados 
         Caption         =   "Fechas Y Bodega Entrada Y Tipo Ficha Tecnica"
         Height          =   195
         Index           =   5
         Left            =   -74520
         TabIndex        =   130
         Top             =   3240
         Width           =   3855
      End
      Begin VB.TextBox TxtDesperdicio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         TabIndex        =   123
         Top             =   4680
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Frame FrameDesperdicio 
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
         ForeColor       =   &H000000FF&
         Height          =   1815
         Left            =   -68400
         TabIndex        =   119
         Top             =   1320
         Width           =   3015
         Begin VB.OptionButton OptDesResDef 
            Caption         =   "Resumen Con Defectos Proveedor"
            ForeColor       =   &H00C00000&
            Height          =   435
            Left            =   240
            TabIndex        =   176
            Top             =   1320
            Width           =   2535
         End
         Begin VB.OptionButton OptDetalle 
            Caption         =   "Detalle x Catalogo De Productos"
            Height          =   195
            Left            =   240
            TabIndex        =   122
            Top             =   360
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton OptResumen 
            Caption         =   "Resumen"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   240
            TabIndex        =   121
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton OptDesCuaPro 
            Caption         =   "Cuadricula"
            ForeColor       =   &H00008000&
            Height          =   195
            Left            =   240
            TabIndex        =   120
            Top             =   1080
            Width           =   1215
         End
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74520
         TabIndex        =   87
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas Y Cliente Y Ficha Tecnica"
         Height          =   195
         Index           =   1
         Left            =   -74520
         TabIndex        =   86
         Top             =   2280
         Width           =   2895
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas Y Ficha Tecnica"
         Height          =   195
         Index           =   3
         Left            =   -74520
         TabIndex        =   85
         Top             =   3360
         Width           =   2055
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Ficha Tecnica"
         Height          =   195
         Index           =   4
         Left            =   -74520
         TabIndex        =   84
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox TxtSalidas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         TabIndex        =   83
         Top             =   5400
         Width           =   1935
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Salidas No Liberados"
         Height          =   195
         Index           =   6
         Left            =   -74520
         TabIndex        =   82
         Top             =   4440
         Width           =   2295
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Entradas No Liberadas"
         Height          =   195
         Index           =   4
         Left            =   -74400
         TabIndex        =   81
         Top             =   3600
         Width           =   2055
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Ficha Tecnica"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   80
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox TxtEntradas 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72360
         TabIndex        =   79
         Top             =   4800
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Fechas Y Ficha Tecnica"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   78
         Top             =   2520
         Width           =   2655
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Batch Y Linea"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   77
         Top             =   2160
         Width           =   1815
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   76
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton OptInv 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   75
         Top             =   1440
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton OptInv 
         Caption         =   "Descripcion"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   74
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Frame FrameInvTipBus 
         Caption         =   "Tipo De Busqueda"
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
         Height          =   1815
         Left            =   4800
         TabIndex        =   70
         Top             =   2880
         Width           =   2055
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Palabra Inicial"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   73
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Cualquier Palabra"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   72
            Top             =   720
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Igual a"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   71
            Top             =   1080
            Width           =   975
         End
      End
      Begin VB.Frame FrameInvOpc 
         Caption         =   "Opcion Inventario"
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
         Height          =   1335
         Left            =   2640
         TabIndex        =   67
         Top             =   1440
         Width           =   2055
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Bodega Descripcion"
            Height          =   195
            Index           =   2
            Left            =   120
            TabIndex        =   168
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   69
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptInvOpc 
            Caption         =   "Bodega"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   68
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame FrameInvResDet 
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
         ForeColor       =   &H000000FF&
         Height          =   4095
         Left            =   6960
         TabIndex        =   58
         Top             =   1440
         Width           =   2775
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Detalle x Tipo y Codigo"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   179
            Top             =   1080
            Width           =   2175
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen x Bodega y Tipo  y Orden"
            ForeColor       =   &H00C00000&
            Height          =   435
            Index           =   0
            Left            =   120
            TabIndex        =   66
            Top             =   1440
            Width           =   2535
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Detallado x Bodega y Codigo y Orden"
            Height          =   315
            Index           =   1
            Left            =   120
            TabIndex        =   65
            Top             =   240
            Width           =   2415
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Detalle x Tipo y Bodega y Orden"
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   64
            Top             =   720
            Width           =   2415
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen x Tipo y Bodega y Orden"
            ForeColor       =   &H00C00000&
            Height          =   435
            Index           =   3
            Left            =   120
            TabIndex        =   63
            Top             =   1920
            Width           =   2415
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen x Codigo y Bodega"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   62
            Top             =   2760
            Width           =   2415
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen x Bodega y Tipo "
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   61
            Top             =   2400
            Width           =   2415
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen Cuadricula x Codigo"
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   60
            Top             =   3360
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton OptInvResDet 
            Caption         =   "Resumen Cuadricula x Orden"
            ForeColor       =   &H00004000&
            Height          =   195
            Index           =   7
            Left            =   120
            TabIndex        =   59
            Top             =   3720
            Width           =   2535
         End
      End
      Begin VB.TextBox TxtInv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   57
         Top             =   6960
         Width           =   1695
      End
      Begin VB.TextBox TxtInvOpc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   56
         ToolTipText     =   "Signo '+' o Doble Click para Ayuda"
         Top             =   6600
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Batch Y Linea"
         Height          =   195
         Index           =   5
         Left            =   -74520
         TabIndex        =   55
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Frame FrameDespachos 
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
         Height          =   1935
         Left            =   -68280
         TabIndex        =   50
         Top             =   1560
         Width           =   2415
         Begin VB.OptionButton OptSalDet 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   240
            TabIndex        =   54
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptSalRes 
            Caption         =   "Resumen"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   240
            TabIndex        =   53
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton OptSalResCli 
            Caption         =   "Resumen Por Cliente"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   240
            TabIndex        =   52
            Top             =   1080
            Width           =   1815
         End
         Begin VB.OptionButton OptSalCua 
            Caption         =   "Resumen Cuadricula Por Mes y Grafica"
            ForeColor       =   &H00008000&
            Height          =   435
            Left            =   240
            TabIndex        =   51
            Top             =   1440
            Width           =   1815
         End
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas Y Transportista"
         Height          =   195
         Index           =   2
         Left            =   -74520
         TabIndex        =   49
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Frame FrameEntradas 
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
         Height          =   1335
         Left            =   -68280
         TabIndex        =   46
         Top             =   1560
         Width           =   1935
         Begin VB.OptionButton OptEntCua 
            Caption         =   "Cuadricula"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   240
            TabIndex        =   167
            Top             =   1080
            Width           =   1095
         End
         Begin VB.OptionButton OptEntDetalle 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   240
            TabIndex        =   48
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptEntRes 
            Caption         =   "Resumen"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   240
            TabIndex        =   47
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.OptionButton OptCierreTarimas 
         Caption         =   "Fechas "
         Height          =   195
         Index           =   0
         Left            =   -74640
         TabIndex        =   45
         Top             =   1560
         Width           =   975
      End
      Begin VB.OptionButton OptCierreTarimas 
         Caption         =   "Fechas Y Linea"
         Height          =   195
         Index           =   1
         Left            =   -74640
         TabIndex        =   44
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox TxtCieTar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71880
         TabIndex        =   41
         Top             =   4440
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton OptEntradas 
         Caption         =   "Fechas Y Proveedor"
         Height          =   195
         Index           =   5
         Left            =   -74400
         TabIndex        =   40
         Top             =   3240
         Width           =   2295
      End
      Begin VB.OptionButton OptCierreTarimas 
         Caption         =   "Fechas y Ficha Tecnica"
         Height          =   195
         Index           =   2
         Left            =   -74640
         TabIndex        =   39
         Top             =   2280
         Width           =   2175
      End
      Begin VB.OptionButton OptInv 
         Caption         =   "Tipo"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   38
         Top             =   2160
         Width           =   1215
      End
      Begin VB.OptionButton OptInv 
         Caption         =   "Batch y Linea"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   37
         Top             =   2520
         Width           =   1455
      End
      Begin VB.TextBox TxtSalidas2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -71760
         TabIndex        =   36
         Top             =   5760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas Y Cliente Y Tipo Ficha Tecnica"
         Height          =   195
         Index           =   7
         Left            =   -74520
         TabIndex        =   35
         Top             =   2640
         Width           =   3135
      End
      Begin VB.TextBox TxtEntradas2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   -72360
         TabIndex        =   34
         Top             =   5160
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox TxtInvLin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3840
         TabIndex        =   33
         Top             =   7320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame FrameTipExi 
         Caption         =   "Tipo De Existencia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   735
         Left            =   2640
         TabIndex        =   28
         Top             =   4800
         Width           =   4215
         Begin VB.OptionButton OptTipRep 
            Caption         =   "<= 0"
            Height          =   195
            Index           =   1
            Left            =   2040
            TabIndex        =   32
            Top             =   360
            Width           =   615
         End
         Begin VB.OptionButton OptTipRep 
            Caption         =   "> 0"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   31
            Top             =   360
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton OptTipRep 
            Caption         =   "Todos"
            Height          =   195
            Index           =   2
            Left            =   2880
            TabIndex        =   30
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton OptTipRep 
            Caption         =   ">= 0"
            Height          =   195
            Index           =   3
            Left            =   1080
            TabIndex        =   29
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.OptionButton OptInv 
         Caption         =   "Orden"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   27
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Frame FrameInvPro 
         Caption         =   "Tipo De Bodega"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   1815
         Left            =   2640
         TabIndex        =   22
         Top             =   2880
         Width           =   2055
         Begin VB.OptionButton OptInvPro 
            Caption         =   "Otras"
            Height          =   192
            Index           =   3
            Left            =   120
            TabIndex        =   26
            Top             =   1440
            Width           =   735
         End
         Begin VB.OptionButton OptInvPro 
            Caption         =   "No Conforme"
            Height          =   192
            Index           =   2
            Left            =   120
            TabIndex        =   25
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton OptInvPro 
            Caption         =   "Proceso"
            Height          =   192
            Index           =   1
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton OptInvPro 
            Caption         =   "Todas"
            Height          =   192
            Index           =   0
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.OptionButton OptSalidas 
         Caption         =   "Fechas Y Cliente Y Descripcion"
         Height          =   195
         Index           =   8
         Left            =   -74520
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo De Inventario"
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
         Height          =   1335
         Left            =   4800
         TabIndex        =   17
         Top             =   1440
         Width           =   2055
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Todos"
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   20
            Top             =   1080
            Width           =   975
         End
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Materia Prima"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   19
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton OptInvTipBus 
            Caption         =   "Producto Terminado"
            Height          =   195
            Index           =   5
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame FrameCieTarTipRep 
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
         Height          =   1095
         Left            =   -68160
         TabIndex        =   14
         Top             =   1440
         Width           =   1935
         Begin VB.OptionButton OptCieTarRes 
            Caption         =   "Resumen"
            ForeColor       =   &H00C00000&
            Height          =   195
            Left            =   480
            TabIndex        =   16
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton OptCieTarDet 
            Caption         =   "Detalle"
            Height          =   195
            Left            =   480
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.OptionButton OptDesperdicio 
         Caption         =   "Fechas y Grupo"
         Height          =   195
         Index           =   3
         Left            =   -74400
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
      Begin VB.OptionButton OptDesperdicio 
         Caption         =   "Fechas y Ficha Tecnica"
         Height          =   195
         Index           =   2
         Left            =   -74400
         TabIndex        =   12
         Top             =   2520
         Width           =   2175
      End
      Begin VB.OptionButton OptDesperdicio 
         Caption         =   "Fechas y Proceso"
         Height          =   195
         Index           =   1
         Left            =   -74400
         TabIndex        =   11
         Top             =   2160
         Width           =   1695
      End
      Begin VB.OptionButton OptDesperdicio 
         Caption         =   "Fechas"
         Height          =   195
         Index           =   0
         Left            =   -74400
         TabIndex        =   10
         Top             =   1440
         Width           =   975
      End
      Begin MSComCtl2.DTPicker DtpCieTarFecFin 
         Height          =   255
         Left            =   -66960
         TabIndex        =   43
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37449
      End
      Begin MSComCtl2.DTPicker DtpCieTarFecIni 
         Height          =   255
         Left            =   -69600
         TabIndex        =   42
         Top             =   3000
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37449
      End
      Begin MSComCtl2.DTPicker DtpEntFecFin 
         Height          =   255
         Left            =   -68280
         TabIndex        =   88
         Top             =   3480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37336
      End
      Begin MSComCtl2.DTPicker DtpEntFecIni 
         Height          =   255
         Left            =   -70440
         TabIndex        =   89
         Top             =   3480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37336
      End
      Begin MSComCtl2.DTPicker DtpSalFecFin 
         Height          =   255
         Left            =   -69000
         TabIndex        =   90
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37389
      End
      Begin MSComCtl2.DTPicker DtpSalFecIni 
         Height          =   255
         Left            =   -70920
         TabIndex        =   91
         Top             =   4080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37389
      End
      Begin MSComCtl2.DTPicker DTPDesFecFin 
         Height          =   255
         Left            =   -66960
         TabIndex        =   124
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37399
      End
      Begin MSComCtl2.DTPicker DTPDesFecIni 
         Height          =   255
         Left            =   -69480
         TabIndex        =   125
         Top             =   3360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37399
      End
      Begin MSComCtl2.DTPicker DtpTraFecFin 
         Height          =   255
         Left            =   -66720
         TabIndex        =   150
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37330
      End
      Begin MSComCtl2.DTPicker DtpTraFecIni 
         Height          =   255
         Left            =   -68880
         TabIndex        =   149
         Top             =   4800
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   89980931
         CurrentDate     =   37330
      End
      Begin VB.Label LblInvFecFin 
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
         Left            =   2400
         TabIndex        =   172
         Top             =   6240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label LblInvFecIni 
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
         Left            =   2400
         TabIndex        =   171
         Top             =   5880
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label LblTraslados 
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
         Left            =   -74520
         TabIndex        =   158
         Top             =   5160
         Width           =   3135
      End
      Begin VB.Label LblLabel 
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
         Index           =   1
         Left            =   -67320
         TabIndex        =   157
         Top             =   4800
         Width           =   510
      End
      Begin VB.Label LblLabel 
         AutoSize        =   -1  'True
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
         Height          =   195
         Index           =   0
         Left            =   -69600
         TabIndex        =   156
         Top             =   4800
         Width           =   555
      End
      Begin VB.Label LblTraBod 
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
         TabIndex        =   155
         Top             =   5160
         Width           =   4215
      End
      Begin VB.Label LblTraBod2 
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
         TabIndex        =   154
         Top             =   5520
         Width           =   4215
      End
      Begin VB.Label LblTraslados2 
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
         Left            =   -74520
         TabIndex        =   153
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label LblTraEtiDoc 
         AutoSize        =   -1  'True
         Caption         =   "Tipo De Documento"
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
         Left            =   -73080
         TabIndex        =   152
         Top             =   5880
         Visible         =   0   'False
         Width           =   1710
      End
      Begin VB.Label LblTraDesDoc 
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
         TabIndex        =   151
         Top             =   5880
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label LblDesFecIni 
         Alignment       =   1  'Right Justify
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
         Left            =   -70320
         TabIndex        =   129
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label LblDesFecFin 
         Alignment       =   1  'Right Justify
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
         Left            =   -67560
         TabIndex        =   128
         Top             =   3360
         Width           =   510
      End
      Begin VB.Label LblDesEti 
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
         Left            =   -74160
         TabIndex        =   127
         Top             =   4680
         Width           =   2295
      End
      Begin VB.Label LblDesDes 
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
         Left            =   -69720
         TabIndex        =   126
         Top             =   4680
         Width           =   4215
      End
      Begin VB.Label LblSalEti 
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
         Left            =   -74520
         TabIndex        =   118
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Label LblSalDes 
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
         Left            =   -69720
         TabIndex        =   117
         Top             =   5400
         Width           =   4335
      End
      Begin VB.Label LblSalFecIni 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -71640
         TabIndex        =   116
         Top             =   4080
         Width           =   555
      End
      Begin VB.Label LblSalFecFin 
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
         Left            =   -69600
         TabIndex        =   115
         Top             =   4080
         Width           =   510
      End
      Begin VB.Label LblInvDesCod 
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
         Left            =   5640
         TabIndex        =   114
         Top             =   6960
         Width           =   3855
      End
      Begin VB.Label LblCerBulDes 
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
         Left            =   -70080
         TabIndex        =   113
         Top             =   4440
         Width           =   4335
      End
      Begin VB.Label LblCerBulEti 
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
         Left            =   -74280
         TabIndex        =   112
         Top             =   4440
         Width           =   2415
      End
      Begin VB.Label LblEntFecFin 
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
         Left            =   -69000
         TabIndex        =   111
         Top             =   3480
         Width           =   510
      End
      Begin VB.Label LblEntFecIni 
         AutoSize        =   -1  'True
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
         Height          =   195
         Left            =   -71040
         TabIndex        =   110
         Top             =   3480
         Width           =   555
      End
      Begin VB.Label LblEntEti 
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
         Left            =   -74640
         TabIndex        =   109
         Top             =   4800
         Width           =   2175
      End
      Begin VB.Label LblEntDes 
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
         TabIndex        =   108
         Top             =   4800
         Width           =   4935
      End
      Begin VB.Label LblInv2 
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
         Left            =   5640
         TabIndex        =   107
         Top             =   6600
         Width           =   3855
      End
      Begin VB.Label LblInv 
         Alignment       =   1  'Right Justify
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
         Left            =   720
         TabIndex        =   106
         Top             =   6960
         Width           =   3015
      End
      Begin VB.Label LblInvOpc 
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
         Left            =   720
         TabIndex        =   105
         Top             =   6600
         Width           =   3015
      End
      Begin VB.Label LblCieTarFecIni 
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
         Left            =   -70800
         TabIndex        =   104
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label LblCieTarFecFin 
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
         Left            =   -68040
         TabIndex        =   103
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label LblCieTarEti 
         Alignment       =   1  'Right Justify
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
         Left            =   -74640
         TabIndex        =   102
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label LblCieTarDes 
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
         Left            =   -70320
         TabIndex        =   101
         Top             =   4440
         Width           =   4935
      End
      Begin VB.Label LblSalEti2 
         Alignment       =   1  'Right Justify
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
         Left            =   -73095
         TabIndex        =   100
         Top             =   5760
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.Label LblSalDes2 
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
         Left            =   -69720
         TabIndex        =   99
         Top             =   5760
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label LblEntDes2 
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
         TabIndex        =   98
         Top             =   5160
         Width           =   5055
      End
      Begin VB.Label LblEntEti2 
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
         Left            =   -74640
         TabIndex        =   97
         Top             =   5160
         Width           =   2175
      End
      Begin VB.Label LblInvLin 
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
         Left            =   840
         TabIndex        =   96
         Top             =   7320
         Width           =   2895
      End
      Begin VB.Label LblInvLin2 
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
         Left            =   5640
         TabIndex        =   95
         Top             =   7320
         Width           =   3975
      End
      Begin VB.Label LblTraDes2 
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
         Left            =   -69720
         TabIndex        =   94
         Top             =   5760
         Visible         =   0   'False
         Width           =   4335
      End
      Begin VB.Label LblTraDes 
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
         Left            =   -69720
         TabIndex        =   93
         Top             =   5400
         Width           =   4335
      End
      Begin VB.Label LblTraEti 
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
         Left            =   -74520
         TabIndex        =   92
         Top             =   5400
         Width           =   2655
      End
   End
End
Attribute VB_Name = "ReportesProductoTerminado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RBuscaBodega As New ADODB.Recordset
Dim RBuscaFichaTecnica As New ADODB.Recordset
Dim RBuscaCatalogo As New ADODB.Recordset
Dim RBuscaCliente As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaProceso As New ADODB.Recordset
Dim RBuscaTransportista As New ADODB.Recordset
Dim RBuscaTipo As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset
Dim RBuscaDocumento As New ADODB.Recordset

'VARIABLES PARA CARPETA DE INVENTARIO
Dim BInvBodega As Boolean
Dim BInvCatalogo As Boolean
Dim BInvFichaTecnica As Boolean
Dim BInvTipo As Boolean
Dim BInvBodegaGrupo As Boolean
'VARIABLES PARA CARPETA DE ENTRADAS
Dim BEntFichaTecnica As Boolean
Dim BEntProveedor As Boolean
Dim BEntMateriaPrima As Boolean
Dim BEntTipoMateriaPrima As Boolean

'VARIABLES PARA CARPETA DE DESPACHOS
Dim BSalCliente As Boolean
Dim BSalTransportista As Boolean
Dim BSalFichaTecnica As Boolean
Dim BSalTipoFT As Boolean

'VARIABLES PARA CARPETA DE CIERRE TARIMAS
Dim BDesProceso As Boolean
Dim BDesFichaTecnica As Boolean

'VARIABLES PARA TRASLADOS
Dim BTraBodega As Boolean
Dim BTraFichaTecnica As Boolean
Dim BTraDocumentos As Boolean
Dim BTraTipo As Boolean

'CIERRE BULTO
Dim BCieBulFichaTecnica As Boolean
Dim BCieBulLinea As Boolean
Dim BCieBulTipo As Boolean
Dim VDia As String
Dim VMes As String
Dim VAño As String
Dim VDia2 As String
Dim VMes2 As String
Dim VAño2 As String

Dim VTexto As String
Dim VTexto2 As String


Private Sub CmdImprimir_Click()
On Error Resume Next
    MousePointer = 11
    
    'INVENTARIO MATERIA PRIMA
    If TabReportes.Tab = 0 Then
            'VA AL PROCEDIMIENTO DE INVENTARIO
            Inventario
    'ENTRADAS DE MATERIA PRIMA
    ElseIf TabReportes.Tab = 1 Then
            'VA AL PROCEDIMIENTO DE ENTRADAS
            Entradas
    'DESPACHOS DE PRODUCTO TERMINADO
    ElseIf TabReportes.Tab = 2 Then
            'VA AL PROCEDIMIENTO DE SALIDAS
            Despachos
    'CIERRE TARIMA
    ElseIf TabReportes.Tab = 3 Then
            CierreBulto
    'BODEGAS
    ElseIf TabReportes.Tab = 4 Then
            Desperdicio
    'TRASLADOS PRODUCTO TERMINADO
    ElseIf TabReportes.Tab = 5 Then
            Traslados
    End If
    
             'DESPLIEGA EL REPORTE
             FrmReporte.Show
             
             
             MousePointer = 0
             If Err <> 0 Then
                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                Exit Sub
             End If
                
    
End Sub

Private Sub CmdSale_Click()
    'DESABILITA EL GRID
    DBGridBusqueda.AllowUpdate = False
    
    FrameBusqueda.Visible = False
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub

Private Sub DBGridBusqueda_DblClick()
        'INVENTARIO
        If BInvBodega = True Then
                TxtInvOpc.Text = DBGridBusqueda.Columns(0).Text
                TxtInvOpc.SetFocus
        'INVENTARIO
        ElseIf (BInvFichaTecnica = True Or BInvCatalogo = True) Then
                TxtInv.Text = DBGridBusqueda.Columns(0).Text
                TxtInv.SetFocus
        'GRUPO
        ElseIf BInvTipo = True Then
                TxtInv.Text = DBGridBusqueda.Columns(0).Text
                TxtInv.SetFocus
        'ENTRADAS
        ElseIf BEntFichaTecnica = True Then
                TxtEntradas.Text = DBGridBusqueda.Columns(0).Text
                TxtEntradas.SetFocus
        'DESPACHOS
        ElseIf (BSalCliente = True Or BSalTransportista = True) Then
                TxtSalidas.Text = DBGridBusqueda.Columns(0).Text
                TxtSalidas.SetFocus
        'DESPACHOS 2
        ElseIf (BSalFichaTecnica = True Or BSalTipoFT = True) Then
                TxtSalidas2.Text = DBGridBusqueda.Columns(0).Text
                TxtSalidas2.SetFocus
        'TRASLADOS
        ElseIf (BTraBodega = True) Then
                TxtTraslados.Text = DBGridBusqueda.Columns(0).Text
                TxtTraslados.SetFocus
        'TRASLADOS 2
        ElseIf (BTraFichaTecnica = True) Then
                TxtTraslados2.Text = DBGridBusqueda.Columns(0).Text
                TxtTraslados2.SetFocus
        'CIERRE TARIMA
        ElseIf BCieBulFichaTecnica = True Or BCieBulTipo = True Or BCieBulLinea = True Then
                TxtCieTar.Text = DBGridBusqueda.Columns(0).Text
                TxtCieTar.SetFocus
        End If
                'DESABILITA EL GRID
                DBGridBusqueda.AllowUpdate = False
                
                FrameBusqueda.Visible = False
                
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
        'SI PRECIONA LA TECLA DEL SIGNO '+'
        If KeyAscii = 43 Then
                         'INVENTARIO
                        If BInvBodega = True Then
                                TxtInvOpc.Text = DBGridBusqueda.Columns(0).Text
                                TxtInvOpc.SetFocus
                        'INVENTARIO
                        ElseIf (BInvFichaTecnica = True Or BInvCatalogo = True) Then
                                TxtInv.Text = DBGridBusqueda.Columns(0).Text
                                TxtInv.SetFocus
                        'GRUPO
                        ElseIf BInvTipo = True Then
                                TxtInv.Text = DBGridBusqueda.Columns(0).Text
                                TxtInv.SetFocus
                        'ENTRADAS
                        ElseIf BEntFichaTecnica = True Then
                                TxtEntradas.Text = DBGridBusqueda.Columns(0).Text
                                TxtEntradas.SetFocus
                        'DESPACHOS
                        ElseIf (BSalCliente = True Or BSalTransportista = True) Then
                                TxtSalidas.Text = DBGridBusqueda.Columns(0).Text
                                TxtSalidas.SetFocus
                        'DESPACHOS 2
                        ElseIf (BSalFichaTecnica = True Or BSalTipoFT = True) Then
                                TxtSalidas2.Text = DBGridBusqueda.Columns(0).Text
                                TxtSalidas2.SetFocus
                        'TRASLADOS
                        ElseIf (BTraBodega = True) Then
                                TxtTraslados.Text = DBGridBusqueda.Columns(0).Text
                                TxtTraslados.SetFocus
                        'TRASLADOS 2
                        ElseIf (BTraFichaTecnica = True) Then
                                TxtTraslados2.Text = DBGridBusqueda.Columns(0).Text
                                TxtTraslados2.SetFocus
                        'CIERRE TARIMA
                        ElseIf BCieBulFichaTecnica = True Then
                                TxtCieTar.Text = DBGridBusqueda.Columns(0).Text
                                TxtCieTar.SetFocus
                        End If
                        
                                'DESABILITA EL GRID
                                DBGridBusqueda.AllowUpdate = False
                                
                                FrameBusqueda.Visible = False
        End If
        
End Sub

Private Sub Form_Load()
        
        'FECHAS DE TAB DE ENTRADAS
        DtpEntFecIni.Value = Date
        DtpEntFecFin.Value = Date
                
        'FECHAS DE TAB DE DESPACHOS
        DtpSalFecIni.Value = Date
        DtpSalFecFin.Value = Date
        
        'FECHAS DE TAB DE CIERRE TARIMA
        DtpCieTarFecIni.Value = Date
        DtpCieTarFecFin.Value = Date
        
        'FECHAS DE TAB DE TRASLADOS
        DtpTraFecIni.Value = Date
        DtpTraFecFin.Value = Date
        
        DTPDesFecIni.Value = Date
        DTPDesFecFin.Value = Date
                
        
End Sub

Private Sub OptBusqueda_Click(Index As Integer)
        If OptBusqueda.Item(0).Value = True Then
                LblBusqueda.Caption = "Descripcion"
        ElseIf OptBusqueda.Item(1).Value = True Then
            LblBusqueda.Caption = "Codigo"
        End If
            TxtBusqueda.SetFocus
End Sub




Private Sub OptCierreTarimas_Click(Index As Integer)
        If Index = 0 Then
            LblCieTarEti.Caption = ""
            TxtCieTar.Visible = False
        ElseIf Index = 1 Then
            LblCieTarEti.Caption = "Linea"
            TxtCieTar.Visible = True
            TxtCieTar.SetFocus
        ElseIf Index = 2 Then
            LblCieTarEti.Caption = "Ficha Tecnica"
            TxtCieTar.Visible = True
            TxtCieTar.SetFocus
        ElseIf Index = 3 Then
            LblCieTarEti.Caption = "Descripcion"
            TxtCieTar.Visible = True
            TxtCieTar.SetFocus
        ElseIf Index = 4 Then
            LblCieTarEti.Caption = "Tipo Ficha Tecnica"
            TxtCieTar.Visible = True
            TxtCieTar.SetFocus
        End If
        
End Sub


Private Sub OptDesperdicio_Click(Index As Integer)
        If Index = 0 Then
            LblDesEti.Caption = ""
            TxtDesperdicio.Visible = False
        ElseIf Index = 1 Then
            LblDesEti.Caption = "Codigo De Proceso"
            TxtDesperdicio.Visible = True
            TxtDesperdicio.SetFocus
        ElseIf Index = 2 Then
            LblDesEti.Caption = "Codigo Ficha Tecnica"
            TxtDesperdicio.Visible = True
            TxtDesperdicio.SetFocus
        ElseIf Index = 3 Then
            LblDesEti.Caption = "Codigo De Grupo"
            TxtDesperdicio.Visible = True
            TxtDesperdicio.SetFocus

        End If

End Sub

Private Sub OptEntradas_Click(Index As Integer)

        'BATCH Y LINEA
        If Index = 1 Then
                TxtEntradas2.Visible = True
                TxtEntradas2.SetFocus
                LblEntEti2.Caption = "Linea"
        Else
                TxtEntradas2.Visible = False
                LblEntEti2.Caption = ""
        End If

        'FECHAS
        If Index = 0 Then
                DtpEntFecIni.Visible = True
                DtpEntFecFin.Visible = True
                LblEntFecIni.Visible = True
                LblEntFecFin.Visible = True
                TxtEntradas.Text = ""
                TxtEntradas.Visible = False
                LblEntEti.Caption = ""
        'BATCH
        ElseIf Index = 1 Then
                DtpEntFecIni.Visible = False
                DtpEntFecFin.Visible = False
                LblEntFecIni.Visible = False
                LblEntFecFin.Visible = False
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Batch"
        'FECHAS Y FICHA TECNICA
        ElseIf Index = 2 Then
                DtpEntFecIni.Visible = True
                DtpEntFecFin.Visible = True
                LblEntFecIni.Visible = True
                LblEntFecFin.Visible = True
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Ficha Tecnica"
        'FICHA TECNICA
        ElseIf Index = 3 Then
                DtpEntFecIni.Visible = False
                DtpEntFecFin.Visible = False
                LblEntFecIni.Visible = False
                LblEntFecFin.Visible = False
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Ficha Tecnica"
        'NO LIBERADAS
        ElseIf Index = 4 Then
                DtpEntFecIni.Visible = False
                DtpEntFecFin.Visible = False
                LblEntFecIni.Visible = False
                LblEntFecFin.Visible = False
                TxtEntradas.Text = ""
                TxtEntradas.Visible = False
                LblEntEti.Caption = ""
        'FECHAS Y LINEA
        ElseIf Index = 5 Then
                DtpEntFecIni.Visible = True
                DtpEntFecFin.Visible = True
                LblEntFecIni.Visible = True
                LblEntFecFin.Visible = True
                TxtEntradas.Visible = True
                TxtEntradas.SetFocus
                LblEntEti.Caption = "Proveedor"
        End If
End Sub

Private Sub OptInv_Click(Index As Integer)
    
    'BATCH
    If Index = 3 Then
        TxtInvLin.Visible = True
        LblInvLin.Caption = "Linea"
    Else
        TxtInvLin.Visible = True
        LblInvLin.Caption = ""
    End If
    
    If Index = 0 Then
        LblInv.Caption = "Codigo "
    ElseIf Index = 1 Then
        LblInv.Caption = "Descripcion"
    ElseIf Index = 2 Then
        LblInv.Caption = "Tipo"
    ElseIf Index = 3 Then
        LblInv.Caption = "Batch"
    ElseIf Index = 4 Then
        LblInv.Caption = "Catalogo"
    ElseIf Index = 5 Then
        LblInv.Caption = "Interno o Externo"
    ElseIf Index = 6 Then
        LblInv.Caption = "Orden"
    ElseIf Index = 7 Then
        LblInv.Caption = "Pasillo"
    End If
        TxtInv.SetFocus
        
End Sub

Private Sub OptInvFec_Click(Index As Integer)
        If Index = 0 Then
            LblInvFecIni.Visible = True
            LblInvFecFin.Visible = True
            DtpInvFecIni.Visible = True
            DtpInvFecFin.Visible = True
        Else
            LblInvFecIni.Visible = False
            LblInvFecFin.Visible = False
            DtpInvFecIni.Visible = False
            DtpInvFecFin.Visible = False
        End If
End Sub

Private Sub OptInvOpc_Click(Index As Integer)
    If Index = 0 Then
        LblInvOpc.Caption = ""
        TxtInvOpc.Text = ""
        TxtInvOpc.Visible = False
    ElseIf Index = 1 Then
        LblInvOpc.Caption = "Codigo Bodega"
        TxtInvOpc.Visible = True
        TxtInvOpc.Text = ""
        TxtInvOpc.SetFocus
    ElseIf Index = 2 Then
        LblInvOpc.Caption = "Descrip. Bodega"
        TxtInvOpc.Visible = True
        TxtInvOpc.Text = ""
        TxtInvOpc.SetFocus

    End If
    
End Sub

Private Sub OptSalidas_Click(Index As Integer)
        'POR FECHAS
        If Index = 0 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = False
                LblSalEti.Caption = ""
                LblSalDes.Caption = ""
        'POR FECHAS Y CODIGO DE CLIENTE
        ElseIf Index = 1 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = True
                TxtSalidas.SetFocus
                LblSalEti.Caption = "Codigo Cliente"
        'POR FECHAS Y CODIGO DE TRANSPORTISTA
        ElseIf Index = 2 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = True
                TxtSalidas.SetFocus
                LblSalEti.Caption = "Codigo Transportista"
        'POR FECHAS Y FICHA TECNICA
        ElseIf Index = 3 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = False
                TxtSalidas2.Visible = True
                TxtSalidas2.SetFocus
                LblSalEti.Caption = ""
                TxtSalidas.Text = ""
                LblSalDes.Caption = ""
                LblSalEti2.Caption = "Ficha Tecnica"
        'SOLO POR FICHA TECNICA
        ElseIf Index = 4 Then
                DtpSalFecIni.Visible = False
                DtpSalFecFin.Visible = False
                LblSalFecIni.Visible = False
                LblSalFecFin.Visible = False
                TxtSalidas.Visible = False
                TxtSalidas2.Visible = True
                TxtSalidas2.SetFocus
                LblSalEti2.Caption = "Ficha Tecnica"
                LblSalEti.Caption = ""
                TxtSalidas.Text = ""
                LblSalDes.Caption = ""
        'BATCH
        ElseIf Index = 5 Then
                DtpSalFecIni.Visible = False
                DtpSalFecFin.Visible = False
                LblSalFecIni.Visible = False
                LblSalFecFin.Visible = False
                TxtSalidas.Visible = True
                TxtSalidas.SetFocus
                LblSalEti.Caption = "Batch"
        'NO LIBERADAS
        ElseIf Index = 6 Then
                DtpSalFecIni.Visible = False
                DtpSalFecFin.Visible = False
                LblSalFecIni.Visible = False
                LblSalFecFin.Visible = False
                TxtSalidas.Visible = False
                TxtSalidas.Visible = False
                LblSalEti.Caption = ""
                TxtSalidas.Text = ""
                LblSalDes.Caption = ""
                LblSalEti2.Caption = ""
                TxtSalidas2.Text = ""
                LblSalDes2.Caption = ""
        'POR FECHAS Y CODIGO DE CLIENTE Y Tipo De Ficha Tecnica
        ElseIf Index = 7 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = True
                TxtSalidas.SetFocus
                LblSalEti.Caption = "Cliente"
        'POR FECHAS Y CODIGO DE CLIENTE
        ElseIf Index = 8 Then
                DtpSalFecIni.Visible = True
                DtpSalFecFin.Visible = True
                LblSalFecIni.Visible = True
                LblSalFecFin.Visible = True
                TxtSalidas.Visible = True
                TxtSalidas.SetFocus
                LblSalEti.Caption = "Codigo Cliente"
        End If
        
        'SI ELIGE LAS OPCIONES DE FICHA TECNICA ADICIONAL
        If (Index = 1 Or Index = 7 Or Index = 3 Or Index = 4 Or Index = 5 Or Index = 8) Then
                LblSalEti2.Visible = True
                TxtSalidas2.Visible = True
                LblSalDes2.Visible = True
        Else
                LblSalEti2.Visible = False
                TxtSalidas2.Visible = False
                LblSalDes2.Visible = False
            
        End If
        
        'SI ES POR FICHA TECNICA
        If (Index = 1 Or Index = 3 Or Index = 4) Then
                LblSalEti2.Caption = "Ficha Tecnica"
        ElseIf Index = 5 Then
                LblSalEti2.Caption = "Linea"
        ElseIf Index = 7 Then
                LblSalEti2.Caption = "Tipo De Ficha Tecnica"
        ElseIf Index = 8 Then
                LblSalEti2.Caption = "Descripcion"
        End If
                
        

End Sub


Private Sub OptTraOpc_Click(Index As Integer)
        If Index = 0 Then
            LblTraEtiDoc.Visible = False
            TxtTraTipDoc.Visible = False
            LblTraDesDoc.Visible = False
        Else
            LblTraEtiDoc.Visible = True
            TxtTraTipDoc.Visible = True
            LblTraDesDoc.Visible = True
        End If
End Sub

Private Sub OptTraslados_Click(Index As Integer)
        'FECHAS
        If OptTraslados.Item(0).Value = True Then
            TxtTraslados.Visible = False
            TxtTraslados2.Visible = False
            LblTraslados.Caption = ""
            LblLabel.Item(0).Visible = True
            LblLabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'DOCUMENTO
        ElseIf OptTraslados.Item(1).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = False
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Numero De Documento"
            LblLabel.Item(0).Visible = False
            LblLabel.Item(1).Visible = False
            DtpTraFecIni.Visible = False
            DtpTraFecFin.Visible = False
        'FECHAS Y BODEGA DE SALIDA Y CODIGO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(2).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = True
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Materia Prima"
            LblTraslados2.Caption = "Codigo Bodega Salida"
            LblLabel.Item(0).Visible = True
            LblLabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'FECHAS Y BODEGA DE ENTRADA Y CODIGO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(3).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = True
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Materia Prima"
            LblTraslados2.Caption = "Codigo Bodega Entrada"
            LblLabel.Item(0).Visible = True
            LblLabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'FECHAS Y BODEGA DE SALIDA Y TIPO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(4).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = True
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Bodega Salida"
            LblTraslados2.Caption = "Codigo Tipo De Materia Prima"
            LblLabel.Item(0).Visible = True
            LblLabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'FECHAS Y BODEGA DE ENTRADA Y TIPO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(5).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = True
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Bodega Entrada"
            LblTraslados2.Caption = "Codigo Tipo De Materia Prima"
            LblLabel.Item(0).Visible = True
            LblLabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'FECHAS Y CODIGO DE MATERIA PRIMA
        ElseIf OptTraslados.Item(6).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = False
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Materia Prima"
            LblLabel.Item(0).Visible = True
            LblLabel.Item(1).Visible = True
            DtpTraFecIni.Visible = True
            DtpTraFecFin.Visible = True
        'NUMERO DE INGRESO
        ElseIf OptTraslados.Item(7).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = False
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Numero De Documento"
            LblLabel.Item(0).Visible = False
            LblLabel.Item(1).Visible = False
            DtpTraFecIni.Visible = False
            DtpTraFecFin.Visible = False
        'MATERIA PRIMA
        ElseIf OptTraslados.Item(8).Value = True Then
            TxtTraslados.Visible = True
            TxtTraslados2.Visible = False
            TxtTraslados.SetFocus
            LblTraslados.Caption = "Codigo Materia Prima"
            LblLabel.Item(0).Visible = False
            LblLabel.Item(1).Visible = False
            DtpTraFecIni.Visible = False
            DtpTraFecFin.Visible = False
        'NO LIBERADO
        ElseIf OptTraslados.Item(9).Value = True Then
            TxtTraslados.Visible = False
            TxtTraslados2.Visible = False
            LblTraslados.Caption = ""
            LblLabel.Item(0).Visible = False
            LblLabel.Item(1).Visible = False
            DtpTraFecIni.Visible = False
            DtpTraFecFin.Visible = False
        End If
                
        

End Sub

Private Sub TabReportes_Click(PreviousTab As Integer)
            'INVENTARIO
            If TabReportes.Tab = 0 Then
                    OptInv.Item(0).Value = True
            'ENTRADAS
            ElseIf TabReportes.Tab = 1 Then
                    OptEntradas.Item(0).Value = True
            'DESPACHOS
            ElseIf TabReportes.Tab = 2 Then
                    OptSalidas.Item(0).Value = True
            'CIERRE TARIMAS
            ElseIf TabReportes.Tab = 3 Then
                    OptCierreTarimas.Item(0).Value = True
            'BODEGAS
            ElseIf TabReportes.Tab = 4 Then
                    OptDesperdicio(0).Value = True
            'TRASLADOS
            ElseIf TabReportes.Tab = 5 Then
                    OptTraslados.Item(0).Value = True
            End If
End Sub

Private Sub TxtBusqueda_Change()
        Set RBusqueda = New ADODB.Recordset
        
        'BODEGAS EN CARPETA DE INVENTARIO
        If (BInvBodega = True Or BTraBodega = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select * From BodegasInventario Where CodigoBodega Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select * From BodegasInventario Where UPPER(CodigoBodega) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
        
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select * From BodegasInventario Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select * From BodegasInventario Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
        
        'LINEAS
        ElseIf BCieBulLinea = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where Lineas Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where Descrip Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
            End If
        
        'CATALOGO
        ElseIf BInvCatalogo = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select * From VariablesDescripcion Where CodigoVariable Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select * From VariablesDescripcion Where UPPER(CodigoVariable) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select * From VariablesDescripcion Where DescripcionVariable '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select * From VariablesDescripcion Where UPPER(DescripcionVariable) '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
            End If
        'FICHA TECNICA
        ElseIf (BSalFichaTecnica = True Or BTraFichaTecnica = True Or BInvFichaTecnica = True Or BInvTipo = True Or BEntFichaTecnica = True Or BCieBulFichaTecnica = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Esp_Tec Like '%" & TxtBusqueda.Text & "%' And Activa = -1")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%' And Activa = -1")
                    End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, Tipo, Envases, Origen From FichaTecnica Where Descrip Like '%" & TxtBusqueda.Text & "%' And Activa = -1")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, Tipo, Envases, Origen From FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%' And Activa = -1")
                    End If
            End If
        'TIPO DE FICHA TECNICA
        ElseIf (BInvTipo = True Or BSalTipoFT = True Or BCieBulTipo = True) Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos Where CodigoTipo Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos Where UPPER(CodigoTipo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            End If
        'CLIENTES
        ElseIf BSalCliente = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where CodigoCliente Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where UPPER(CodigoCliente) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                    Else
                        Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                    End If
        
            End If
        'TRANSPORTISTA
        ElseIf BSalTransportista = True Then
            'CODIGO
            If OptBusqueda.Item(0).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Transportistas Where Codigo Like '%" & TxtBusqueda.Text & "%'")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Transportistas Where UPPER(Codigo) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                End If
            'DESCRIPCION
            ElseIf OptBusqueda.Item(1).Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Transportistas Where Descripcion Like '%" & TxtBusqueda.Text & "%'")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Transportistas Where UPPER(Descripcion) Like '%" & UCase(TxtBusqueda.Text) & "%'")
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


Private Sub TxtCieTar_Change()
        'BUSCA CODIGO DE FICHA TECNICA
        If OptCierreTarimas.Item(2).Value = True Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCieTar.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCieTar.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblCieTarDes.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblCieTarDes.Caption = ""
                End If
        ElseIf OptCierreTarimas.Item(4).Value = True Then
            Set RBuscaTipo = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicaTipos Where CodigoTipo = '" & TxtCieTar.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicaTipos Where UPPER(CodigoTipo) = '" & UCase(TxtCieTar.Text) & "'")
                End If
                If RBuscaTipo.RecordCount > 0 Then
                    LblCieTarDes.Caption = RBuscaTipo!Descripcion
                Else
                    LblCieTarDes.Caption = ""
                End If
        Else
                    LblCieTarDes.Caption = ""
        End If
End Sub

Private Sub TxtCieTar_DblClick()
                Set RBusqueda = New ADODB.Recordset
                'OPCION POR FICHA TECNICA
                If OptCierreTarimas.Item(1).Value = True Then
                            'INVENTARIO
                            BInvBodega = False
                            BInvCatalogo = False
                            BInvFichaTecnica = False
                            BInvTipo = False
                            'ENTRADAS
                            BEntFichaTecnica = False
                            'DESPACHOS
                            BSalCliente = False
                            BSalTransportista = False
                            BSalFichaTecnica = False
                            BSalTipoFT = False
                            'CIERRE TARIMA
                            BCieBulLinea = True
                            BCieBulFichaTecnica = False
                            BCieBulTipo = False
                            
                            'TRASLADOS
                            BTraBodega = False
                            BTraFichaTecnica = False
                            BTraDocumentos = False
                            'DESPERDICIO
                            BDesProceso = False
                            BDesFichaTecnica = False
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
                ElseIf OptCierreTarimas.Item(2).Value = True Then
                            'INVENTARIO
                            BInvBodega = False
                            BInvCatalogo = False
                            BInvFichaTecnica = False
                            BInvTipo = False
                            'ENTRADAS
                            BEntFichaTecnica = False
                            'DESPACHOS
                            BSalCliente = False
                            BSalTransportista = False
                            BSalFichaTecnica = False
                            BSalTipoFT = False
                            'CIERRE TARIMA
                            BCieBulLinea = False
                            BCieBulFichaTecnica = True
                            BCieBulTipo = False
                            
                            'TRASLADOS
                            BTraBodega = False
                            BTraFichaTecnica = False
                            BTraDocumentos = False
                            'DESPERDICIO
                            BDesProceso = False
                            BDesFichaTecnica = False
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Activa = -1")
                ElseIf OptCierreTarimas.Item(4).Value = True Then
                            'INVENTARIO
                            BInvBodega = False
                            BInvCatalogo = False
                            BInvFichaTecnica = False
                            BInvTipo = False
                            'ENTRADAS
                            BEntFichaTecnica = False
                            'DESPACHOS
                            BSalCliente = False
                            BSalTransportista = False
                            BSalFichaTecnica = False
                            BSalTipoFT = False
                            'CIERRE TARIMA
                            BCieBulLinea = False
                            BCieBulFichaTecnica = False
                            BCieBulTipo = True
                            
                            'TRASLADOS
                            BTraBodega = False
                            BTraFichaTecnica = False
                            BTraDocumentos = False
                            'DESPERDICIO
                            BDesProceso = False
                            BDesFichaTecnica = False
                            Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos")
                End If
        
                If OptCierreTarimas.Item(1).Value = True Or OptCierreTarimas.Item(2).Value = True Or OptCierreTarimas.Item(4).Value = True Then
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                End If

End Sub

Private Sub TxtCieTar_GotFocus()
        TxtCieTar.SelStart = 0
        TxtCieTar.SelLength = Len(TxtCieTar.Text)
End Sub



Private Sub TxtCieTar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
                Set RBusqueda = New ADODB.Recordset
                'OPCION POR FICHA TECNICA
                If OptCierreTarimas.Item(1).Value = True Then
                            'INVENTARIO
                            BInvBodega = False
                            BInvCatalogo = False
                            BInvFichaTecnica = False
                            BInvTipo = False
                            'ENTRADAS
                            BEntFichaTecnica = False
                            'DESPACHOS
                            BSalCliente = False
                            BSalTransportista = False
                            BSalFichaTecnica = False
                            BSalTipoFT = False
                            'CIERRE TARIMA
                            BCieBulLinea = True
                            BCieBulFichaTecnica = False
                            BCieBulTipo = False
                            
                            'TRASLADOS
                            BTraBodega = False
                            BTraFichaTecnica = False
                            BTraDocumentos = False
                            'DESPERDICIO
                            BDesProceso = False
                            BDesFichaTecnica = False
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas")
                ElseIf OptCierreTarimas.Item(2).Value = True Then
                            'INVENTARIO
                            BInvBodega = False
                            BInvCatalogo = False
                            BInvFichaTecnica = False
                            BInvTipo = False
                            'ENTRADAS
                            BEntFichaTecnica = False
                            'DESPACHOS
                            BSalCliente = False
                            BSalTransportista = False
                            BSalFichaTecnica = False
                            BSalTipoFT = False
                            'CIERRE TARIMA
                            BCieBulLinea = False
                            BCieBulFichaTecnica = True
                            BCieBulTipo = False
                            
                            'TRASLADOS
                            BTraBodega = False
                            BTraFichaTecnica = False
                            BTraDocumentos = False
                            'DESPERDICIO
                            BDesProceso = False
                            BDesFichaTecnica = False
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica Where Activa = -1")
                ElseIf OptCierreTarimas.Item(4).Value = True Then
                            'INVENTARIO
                            BInvBodega = False
                            BInvCatalogo = False
                            BInvFichaTecnica = False
                            BInvTipo = False
                            'ENTRADAS
                            BEntFichaTecnica = False
                            'DESPACHOS
                            BSalCliente = False
                            BSalTransportista = False
                            BSalFichaTecnica = False
                            BSalTipoFT = False
                            'CIERRE TARIMA
                            BCieBulLinea = False
                            BCieBulFichaTecnica = False
                            BCieBulTipo = True
                            
                            'TRASLADOS
                            BTraBodega = False
                            BTraFichaTecnica = False
                            BTraDocumentos = False
                            'DESPERDICIO
                            BDesProceso = False
                            BDesFichaTecnica = False
                            Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos")
                End If
        
                If OptCierreTarimas.Item(1).Value = True Or OptCierreTarimas.Item(2).Value = True Or OptCierreTarimas.Item(4).Value = True Then
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                End If
        End If

End Sub

Private Sub TxtDesperdicio_Change()
        'PROCESO
        If OptDesperdicio.Item(1).Value = True Then
                Set RBuscaProceso = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaProceso, "Select Descripcion From ProcesosMateriaPrima Where CodigoProceso = '" & TxtDesperdicio.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaProceso, "Select Descripcion From ProcesosMateriaPrima Where UPPER(CodigoProceso) = '" & UCase(TxtDesperdicio.Text) & "'")
                    End If
                    If RBuscaProceso.RecordCount > 0 Then
                        LblDesDes.Caption = RBuscaProceso!Descripcion
                    Else
                        LblDesDes.Caption = ""
                    End If
        'FICHA TECNICA
        ElseIf OptDesperdicio.Item(2).Value = True Then
                Set RBuscaFichaTecnica = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtDesperdicio.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtDesperdicio.Text) & "'")
                    End If
                    If RBuscaFichaTecnica.RecordCount > 0 Then
                        LblDesDes.Caption = RBuscaFichaTecnica!Descrip
                    Else
                        LblDesDes.Caption = ""
                    End If
        End If

End Sub

Private Sub TxtDesperdicio_DblClick()
        Set RBusqueda = New ADODB.Recordset
        'OPCION POR PROCESO
        If OptDesperdicio.Item(1).Value = True Then
                    'INVENTARIO
                    BInvBodega = True
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    BInvTipo = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = False
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BCieBulTipo = False
                   'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = True
                    BDesFichaTecnica = False
                    
        'FICHA TECNICA
        ElseIf OptDesperdicio.Item(2).Value = True Then
                    'INVENTARIO
                    BInvBodega = True
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    BInvTipo = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = False
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BCieBulTipo = False
                   'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = True
        End If
    
        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
        If BDesProceso = True Then
                Call Abrir_Recordset(RBusqueda, "Select * From ProcesosMateriaPrima")
        'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
        ElseIf BDesFichaTecnica = True Then
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
        End If
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus

End Sub

Private Sub TxtDesperdicio_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
                SendKeys "{tab}"
            End If
            
            If KeyAscii = 43 Then
                        Set RBusqueda = New ADODB.Recordset
                        'OPCION POR PROCESO
                        If OptDesperdicio.Item(1).Value = True Then
                                    BInvBodega = False
                                    BInvTipo = False
                                    BInvFichaTecnica = False
                                    BInvBodegaGrupo = False
                                    BTraBodega = False
                                    BTraFichaTecnica = False
                                    BTraDocumentos = False
                                    'CIERRE TARIMA
                                    BCieBulLinea = False
                                    BCieBulFichaTecnica = False
                                    BCieBulTipo = False
                                    BEntProveedor = False
                                    BEntMateriaPrima = False
                                    BEntTipoMateriaPrima = False
                                    BSalFichaTecnica = False
                                    BSalCliente = False
                                    BCieBulLinea = False
                                    BCieBulFichaTecnica = False
                                    BDesProceso = True
                                    BDesFichaTecnica = False
                                    BTraTipo = False
                                    BCieBulTipo = False
                        'FICHA TECNICA
                        ElseIf OptDesperdicio.Item(2).Value = True Then
                                    BInvBodega = False
                                    BInvTipo = False
                                    BInvFichaTecnica = False
                                    BInvBodegaGrupo = False
                                    BTraBodega = False
                                    BTraFichaTecnica = False
                                    BTraDocumentos = False
                                    'CIERRE TARIMA
                                    BCieBulLinea = False
                                    BCieBulFichaTecnica = False
                                    BCieBulTipo = False
                                    BEntProveedor = False
                                    BEntMateriaPrima = False
                                    BEntTipoMateriaPrima = False
                                    BSalFichaTecnica = False
                                    BSalCliente = False
                                    BCieBulLinea = False
                                    BCieBulFichaTecnica = False
                                    BDesProceso = False
                                    BDesFichaTecnica = True
                                    BTraTipo = False
                                    BCieBulTipo = False
                        End If
                    
                        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
                                If BDesProceso = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select * From ProcesosMateriaPrima")
                                'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
                                ElseIf BDesFichaTecnica = True Then
                                        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica")
                                End If
                                
                                Set DBGridBusqueda.DataSource = RBusqueda
                                DBGridBusqueda.Columns(1).Width = "4000"
                                FrameBusqueda.Visible = True
                                TxtBusqueda.SetFocus

            End If
End Sub

Private Sub TxtEntradas_Change()
        'BUSCA CODIGO DE FICHA TECNICA
        If (OptEntradas.Item(2).Value = True Or OptEntradas.Item(3).Value = True) Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtEntradas.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtEntradas.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblEntDes.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblEntDes.Caption = ""
                End If
        'PROVEEDOR
        ElseIf OptEntradas.Item(5).Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descripcion From Proveedores Where CodigoProveedor = '" & TxtEntradas.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaLinea, "Select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtEntradas.Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblEntDes.Caption = RBuscaLinea!Descripcion
                Else
                    LblEntDes.Caption = ""
                End If
        Else
                    LblEntDes.Caption = ""
        End If
End Sub

Private Sub TxtEntradas_DblClick()
        'OPCION POR FICHA TECNICA
        If (OptEntradas.Item(2).Value = True Or OptEntradas.Item(3).Value = True) Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    BInvTipo = False
                    'ENTRADAS
                    BEntFichaTecnica = True
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = False
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BCieBulTipo = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False
                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "4000"
                    
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If

End Sub

Private Sub TxtEntradas_GotFocus()
        TxtEntradas.SelStart = 0
        TxtEntradas.SelLength = Len(TxtEntradas.Text)
        
End Sub

Private Sub TxtEntradas_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
                
        If KeyAscii = 43 Then
                'OPCION POR FICHA TECNICA
                If (OptEntradas.Item(2).Value = True Or OptEntradas.Item(3).Value = True) Then
                            'INVENTARIO
                            BInvBodega = False
                            BInvCatalogo = False
                            BInvFichaTecnica = False
                            BInvTipo = False
                            'ENTRADAS
                            BEntFichaTecnica = True
                            'DESPACHOS
                            BSalCliente = False
                            BSalTransportista = False
                            BSalFichaTecnica = False
                            BSalTipoFT = False
                            'CIERRE TARIMA
                            BCieBulLinea = False
                            BCieBulFichaTecnica = False
                            BCieBulTipo = False
                            'TRASLADOS
                            BTraBodega = False
                            BTraFichaTecnica = False
                            BTraDocumentos = False
                            
                            'DESPERDICIO
                            BDesProceso = False
                            BDesFichaTecnica = False

                            Set RBusqueda = New ADODB.Recordset
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
                            Set DBGridBusqueda.DataSource = RBusqueda
                            DBGridBusqueda.Columns(1).Width = "4000"
                            FrameBusqueda.Visible = True
                            TxtBusqueda.SetFocus
                End If
        End If

        
End Sub

Private Sub TxtEntradas2_Change()
        If OptEntradas.Item(1).Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtEntradas2.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtEntradas2.Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblEntDes2.Caption = RBuscaLinea!Descrip
                Else
                    LblEntDes2.Caption = ""
                End If
        End If
           
End Sub

Private Sub TxtInv_Change()
        'CODIGO DE FICHA TECNICA
        If OptInv.Item(0).Value = True Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtInv.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtInv.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblInvDesCod.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblInvDesCod.Caption = ""
                End If
        'TIPO DE FICHA TECNICA
        ElseIf OptInv.Item(2).Value = True Then
            Set RBuscaTipo = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicaTipos Where CodigoTipo = '" & TxtInv.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicaTipos Where UPPER(CodigoTipo) = '" & UCase(TxtInv.Text) & "'")
                End If
                
                If RBuscaTipo.RecordCount > 0 Then
                    LblInvDesCod.Caption = RBuscaTipo!Descripcion
                Else
                    LblInvDesCod.Caption = ""
                End If
        'BUSCA CATALOGO
        End If

End Sub

Private Sub TxtInv_DblClick()
        Set RBusqueda = New ADODB.Recordset
        If OptInv.Item(0).Value = True Then
                'OPCION DE FICHA TECNICA CARPETA DE INVENTARIO
                BInvBodega = False
                BInvCatalogo = False
                BInvFichaTecnica = True
                BInvTipo = False
                'ENTRADAS
                BEntFichaTecnica = False
                'DESPACHOS
                BSalCliente = False
                BSalTransportista = False
                BSalFichaTecnica = False
                BSalTipoFT = False
                'CIERRE TARIMA
                BCieBulLinea = False
                BCieBulFichaTecnica = False
                BCieBulTipo = False
                'TRASLADOS
                BTraBodega = False
                BTraFichaTecnica = False
                BTraDocumentos = False
                'DESPERDICIO
                BDesProceso = False
                BDesFichaTecnica = False
                
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
        ElseIf OptInv.Item(2).Value = True Then
                'OPCION DE TIPO CARPETA DE INVENTARIO
                BInvBodega = False
                BInvCatalogo = False
                BInvFichaTecnica = False
                BInvTipo = True
                'ENTRADAS
                BEntFichaTecnica = False
                'DESPACHOS
                BSalCliente = False
                BSalTransportista = False
                BSalFichaTecnica = False
                BSalTipoFT = False
                'CIERRE TARIMA
                BCieBulFichaTecnica = False
                'TRASLADOS
                BTraBodega = False
                BTraFichaTecnica = False
                
                'DESPERDICIO
                BDesProceso = False
                BDesFichaTecnica = False
                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos")
        
        End If
        
        If (OptInv.Item(0).Value = True Or OptInv.Item(2).Value = True) Then
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        End If
                        
End Sub

Private Sub TxtInv_GotFocus()
                TxtInv.SelStart = 0
                TxtInv.SelLength = Len(TxtInv.Text)
End Sub

Private Sub TxtInv_KeyPress(KeyAscii As Integer)
        'SI PRECIONA LA TECLA DE ENTER
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        'SI PRECIONA LA TECLA DEL SIGNO '+'
        If KeyAscii = 43 Then
            Set RBusqueda = New ADODB.Recordset
            If OptInv.Item(0).Value = True Then
                'OPCION DE FICHA TECNICA CARPETA DE INVENTARIO
                BInvBodega = False
                BInvCatalogo = False
                BInvFichaTecnica = True
                BInvTipo = False
                'ENTRADAS
                BEntFichaTecnica = False
                'DESPACHOS
                BSalCliente = False
                BSalTransportista = False
                BSalFichaTecnica = False
                BSalTipoFT = False
                'CIERRE TARIMA
                BCieBulLinea = False
                BCieBulFichaTecnica = False
                BCieBulTipo = False
                'TRASLADOS
                BTraBodega = False
                BTraFichaTecnica = False
                BTraDocumentos = False
                'DESPERDICIO
                BDesProceso = False
                BDesFichaTecnica = False
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
            ElseIf OptInv.Item(2).Value = True Then
                'OPCION DE TIPO CARPETA DE INVENTARIO
                BInvBodega = False
                BInvCatalogo = False
                BInvFichaTecnica = False
                BInvTipo = True
                'ENTRADAS
                BEntFichaTecnica = False
                'DESPACHOS
                BSalCliente = False
                BSalTransportista = False
                BSalFichaTecnica = False
                BSalTipoFT = False
                'CIERRE TARIMA
                BCieBulFichaTecnica = False
                'TRASLADOS
                BTraBodega = False
                BTraFichaTecnica = False
                
                'DESPERDICIO
                BDesProceso = False
                BDesFichaTecnica = False
                
                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos")
                
            End If
        
            If (OptInv.Item(0).Value = True Or OptInv.Item(2).Value = True) Then
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
            End If
            
        End If

End Sub

Private Sub TxtInvLin_Change()
 'BUSCA LA LINEA
        If OptInv.Item(3).Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtInvLin.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtInvLin.Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblInvLin2.Caption = RBuscaLinea!Descrip
                Else
                    LblInvLin2.Caption = ""
                End If
        End If
End Sub

Private Sub TxtInvOpc_Change()
        'BUSCA LA BODEGA
        If OptInvOpc.Item(1).Value = True Then
                Set RBuscaBodega = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtInvOpc.Text & "'")
                    Else
                        Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtInvOpc.Text) & "'")
                    End If
                    If RBuscaBodega.RecordCount > 0 Then
                        LblInv2.Caption = RBuscaBodega!Descripcion
                    Else
                        LblInv2.Caption = ""
                    End If
        Else
                    LblInv2.Caption = ""
        End If
        
End Sub

Private Sub TxtInvOpc_DblClick()
        'OPCION POR BODEGA
        If OptInvOpc.Item(1).Value = True Then
                    'INVENTARIO
                    BInvBodega = True
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    BInvTipo = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = False
                    BSalTipoFT = False
                     'CIERRE TARIMA
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BCieBulTipo = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False

                    Set RBusqueda = New ADODB.Recordset
                    Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If
        
End Sub

Private Sub TxtInvOpc_GotFocus()
        TxtInvOpc.SelStart = 0
        TxtInvOpc.SelLength = Len(TxtInvOpc.Text)
End Sub

Private Sub TxtInvOpc_KeyPress(KeyAscii As Integer)
            
    'SI PRECIONA LA TECLA ENTER
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
        
    'SI PRECIONA LA TECLA DEL SIGNO '+'
    If KeyAscii = 43 Then
            'OPCION POR BODEGA
            If OptInvOpc.Item(1).Value = True Then
                        'INVENTARIO
                        BInvBodega = True
                        BInvCatalogo = False
                        BInvFichaTecnica = False
                        BInvTipo = False
                        'ENTRADAS
                        BEntFichaTecnica = False
                        'DESPACHOS
                        BSalCliente = False
                        BSalTransportista = False
                        BSalFichaTecnica = False
                        BSalTipoFT = False
                         'CIERRE TARIMA
                        BCieBulLinea = False
                        BCieBulFichaTecnica = False
                        BCieBulTipo = False
                        'TRASLADOS
                        BTraBodega = False
                        BTraFichaTecnica = False
                        BTraDocumentos = False
                        
                        'DESPERDICIO
                        BDesProceso = False
                        BDesFichaTecnica = False
                        
                        Set RBusqueda = New ADODB.Recordset
                        Call Abrir_Recordset(RBusqueda, "Select CodigoBodega, Descripcion From BodegasInventario")
                        Set DBGridBusqueda.DataSource = RBusqueda
                        DBGridBusqueda.Columns(1).Width = "3000"
                        FrameBusqueda.Visible = True
                        TxtBusqueda.SetFocus
            End If

    End If
End Sub


Private Sub TxtSalidas_Change()
        'BUSCA CLIENTE
        If (OptSalidas.Item(1).Value = True Or OptSalidas.Item(7).Value = True) Then
            Set RBuscaCliente = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where CodigoCliente = '" & TxtSalidas.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaCliente, "Select Descripcion From Clientes Where UPPER(CodigoCliente) = '" & UCase(TxtSalidas.Text) & "'")
                End If
                If RBuscaCliente.RecordCount > 0 Then
                    LblSalDes.Caption = RBuscaCliente!Descripcion
                Else
                    LblSalDes.Caption = ""
                End If
        'BUSCA TRANSPORTISTA
        ElseIf OptSalidas.Item(2).Value = True Then
            Set RBuscaTransportista = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTransportista, "Select Descripcion From Transportistas Where Codigo = '" & TxtSalidas.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaTransportista, "Select Descripcion From Transportistas Where UPPER(Codigo) = '" & UCase(TxtSalidas.Text) & "'")
                End If
                If RBuscaTransportista.RecordCount > 0 Then
                    LblSalDes.Caption = RBuscaTransportista!Descripcion
                Else
                    LblSalDes.Caption = ""
                End If
        Else
                    LblSalDes.Caption = ""
        End If

End Sub

Private Sub TxtSalidas_DblClick()

        'OPCION FECHAS Y CLIENTE
        If (OptSalidas.Item(1).Value = True Or OptSalidas.Item(7).Value = True) Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    BInvTipo = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = True
                    BSalTransportista = False
                    BSalFichaTecnica = False
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BCieBulTipo = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False

        'OPCION POR TRANSPORTISTA
        ElseIf OptSalidas.Item(2).Value = True Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = True
                    BSalFichaTecnica = False
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulFichaTecnica = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False

        'OPCION POR FICHA TECNICA
        ElseIf (OptSalidas.Item(3).Value = True Or OptSalidas.Item(4).Value = True) Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = True
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulFichaTecnica = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False

                    
        End If
    
        Set RBusqueda = New ADODB.Recordset
    
        'OPCION DE CLIENTES
        If BSalCliente = True Then
                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        'OPCION DE TRANSPORTISTAS
        ElseIf BSalTransportista = True Then
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Transportistas")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        'OPCION DE FICHAS TECNICAS
        ElseIf BSalFichaTecnica = True Then
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        End If

End Sub

Private Sub TxtSalidas_GotFocus()
        TxtSalidas.SelStart = 0
        TxtSalidas.SelLength = Len(TxtSalidas.Text)
End Sub

Private Sub TxtSalidas_KeyPress(KeyAscii As Integer)
    'SI PRECIONA LA TECLA ENTER
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
        
    'SI PRECIONA LA TECLA DEL SIGNO '+'
    If KeyAscii = 43 Then
        'OPCION FECHAS Y CLIENTE
        If (OptSalidas.Item(1).Value = True Or OptSalidas.Item(7).Value = True) Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    BInvTipo = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = True
                    BSalTransportista = False
                    BSalFichaTecnica = False
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BCieBulTipo = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False

        'OPCION POR TRANSPORTISTA
        ElseIf OptSalidas.Item(2).Value = True Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = True
                    BSalFichaTecnica = False
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulFichaTecnica = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False
        'OPCION POR FICHA TECNICA
        ElseIf (OptSalidas.Item(3).Value = True Or OptSalidas.Item(4).Value = True) Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = True
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulFichaTecnica = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False
        End If
        
        Set RBusqueda = New ADODB.Recordset
    
        'OPCION DE CLIENTES
        If BSalCliente = True Then
                Call Abrir_Recordset(RBusqueda, "Select CodigoCliente, Descripcion From Clientes")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        'OPCION DE TRANSPORTISTAS
        ElseIf BSalTransportista = True Then
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion From Transportistas")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        'OPCION DE FICHAS TECNICAS
        ElseIf BSalFichaTecnica = True Then
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial From FichaTecnica Where Activa = -1")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        End If
    End If
End Sub


Sub Inventario()

            TxtInv.Text = UCase(TxtInv.Text)
            TxtInvOpc.Text = UCase(TxtInvOpc.Text)
            TxtInvLin.Text = UCase(TxtInvLin.Text)

            If OptInvTipBus.Item(0).Value = True Then
                    VTexto = " Like '" & TxtInv.Text & "*'"
                    VTexto2 = " Like '" & TxtInvOpc.Text & "*'"
            ElseIf OptInvTipBus.Item(1).Value = True Then
                    VTexto = " Like '*" & TxtInv.Text & "*'"
                    VTexto2 = " Like '*" & TxtInvOpc.Text & "*'"
            ElseIf OptInvTipBus.Item(2).Value = True Then
                    VTexto = " = '" & TxtInv.Text & "'"
                    VTexto2 = " = '" & TxtInvOpc.Text & "'"
            End If
            
            'CODIGO
            If OptInv.Item(0).Value = True Then
                        'TODOS
                        If OptInvOpc.Item(0).Value = True Then
                            'TITULO
                            GTituloReporte = "Todas Las Bodegas y Codigo " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "UPPERCASE({DetalleEntradasInventario.FichaTecnica}) " & UCase(VTexto)
                        'Bodega
                        ElseIf OptInvOpc.Item(1).Value = True Then
                            GTituloReporte = "Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Codigo " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "UPPERCASE({DetalleEntradasInventario.FichaTecnica}) " & UCase(VTexto) & " AND UPPERCASE({DetalleEntradasInventario.Bodega}) " & UCase(VTexto2)
                        'Descripcion Bodega
                        ElseIf OptInvOpc.Item(2).Value = True Then
                            GTituloReporte = "Por Descripcion Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption
                            GCriteriaReporte = "UPPERCASE({DetalleEntradasInventario.FichaTecnica}) " & UCase(VTexto) & " AND UPPERCASE({DetalleEntradasInventario.Bodega}) = UPPERCASE({BodegasInventario.CodigoBodega}) And UPPERCASE({BodegasInventario.Descripcion}) " & UCase(VTexto2)
                        End If
            'DESCRIPCION DE MATERIA PRIMA
            ElseIf OptInv.Item(1).Value = True Then
                    'TODOS
                        If OptInvOpc.Item(0).Value = True Then
                            GTituloReporte = "Todas Las Bodegas y Descripcion " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "Ucase({FichaTecnica.Descrip}) " & UCase(VTexto)
                        'Bodega
                        ElseIf OptInvOpc.Item(1).Value = True Then
                            GTituloReporte = "Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Descripcion " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "UCase({FichaTecnica.Descrip}) " & UCase(VTexto) & " AND UCase({DetalleEntradasInventario.Bodega}) " & UCase(VTexto2)
                        'Descripcion Bodega
                        ElseIf OptInvOpc.Item(2).Value = True Then
                            GTituloReporte = "Por Descripcion Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption
                            GCriteriaReporte = "Ucase({FichaTecnica.Descrip}) " & UCase(VTexto) & " AND ucase({DetalleEntradasInventario.Bodega}) = UCase({BodegasInventario.CodigoBodega}) And UCase({BodegasInventario.Descripcion}) " & UCase(VTexto2)
                        End If
            'TIPO
            ElseIf OptInv.Item(2).Value = True Then
                        'TODOS
                        If OptInvOpc.Item(0).Value = True Then
                            'TITULO
                            GTituloReporte = "Todas Las Bodegas y Tipo " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "UCase({DetalleEntradasInventario.FichaTecnica}) = UCase({FichaTecnica.Esp_Tec}) And Ucase({FichaTecnica.Tipo}) " & VTexto
                        'Bodega
                        ElseIf OptInvOpc.Item(1).Value = True Then
                            GTituloReporte = "Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Tipo " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "UCase({DetalleEntradasInventario.FichaTecnica}) = UCase({FichaTecnica.Esp_Tec}) And Ucase({FichaTecnica.Tipo}) " & VTexto & " And UCase({DetalleEntradasInventario.Bodega}) " & VTexto2
                        'Descripcion Bodega
                        ElseIf OptInvOpc.Item(2).Value = True Then
                            GTituloReporte = "Por Descripcion Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption
                            GCriteriaReporte = "Ucase({DetalleEntradasInventario.FichaTecnica}) = UCase({FichaTecnica.Esp_Tec}) And Ucase({FichaTecnica.Tipo}) " & VTexto & " AND ucase({DetalleEntradasInventario.Bodega}) = UCase({BodegasInventario.CodigoBodega}) And UCase({BodegasInventario.Descripcion}) " & VTexto2
                        End If
            'BATCH (aqui no utilizo la variable VTexto ya que es un campo numerico)
            ElseIf OptInv.Item(3).Value = True Then
                        'TODOS
                        If OptInvOpc.Item(0).Value = True Then
                            'TITULO
                            GTituloReporte = "Todas Las Bodegas y Batch " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "{DetalleEntradasInventario.Batch} = " & TxtInv.Text & " And Ucase({DetalleEntradasInventario.Linea}) = '" & TxtInvLin.Text & "'"
                        'Bodega
                        ElseIf OptInvOpc.Item(1).Value = True Then
                            GTituloReporte = "Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Batch " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "{DetalleEntradasInventario.Batch} = " & TxtInv.Text & " And Ucase({DetalleEntradasInventario.Linea}) = '" & TxtInvLin.Text & "' And UCase({DetalleEntradasInventario.Bodega}) " & VTexto2
                        'Descripcion Bodega
                        ElseIf OptInvOpc.Item(2).Value = True Then
                            GTituloReporte = "Por Descripcion Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Batch " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "{DetalleEntradasInventario.Batch} = " & TxtInv.Text & " And Ucase({DetalleEntradasInventario.Linea}) = '" & TxtInvLin.Text & "' And UCase({DetalleEntradasInventario.Bodega}) = UCase({BodegasInventario.CodigoBodega}) And UCase({BodegasInventario.Descripcion}) " & VTexto2
                        
                        End If
            'CATALOGO
            'ElseIf OptInv.Item(4).Value = True Then
                        'TODOS
            '            If OptInvOpc.Item(0).Value = True Then
                            'TITULO
            '                GTituloReporte = "Texto = 'Todas Las Bodegas y Catalogo " & TxtInv.Text & " " & LblInvDesCod.Caption & "'"
            '                GCriteriaReporte = "{DetalleEntradasInventario.FichaTecnica} = {FichaTecnica.Esp_Tec} And {FichaTecnica.Variables} " & VTexto
            '            'Bodega
            '            ElseIf OptInvOpc.Item(1).Value = True Then
            '                GTituloReporte = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Catalogo " & TxtInv.Text & " " & LblInvDesCod.Caption & "'"
            '                GCriteriaReporte = "{FichaTecnica.Variables} " & VTexto & " And {DetalleEntradasInventario.Bodega} " & VTexto2
            '            End If
            ''INTERNO Y EXTERNO
            'ElseIf OptInv.Item(5).Value = True Then
            '            'TODOS
            '            If OptInvOpc.Item(0).Value = True Then
            '                'TITULO
            '                GTituloReporte = "Texto = 'Todas Las Bodegas y Ficha Tecnica " & TxtInv.Text & " " & LblInvDesCod.Caption & "'"
            '                GCriteriaReporte = "{DetalleEntradasInventario.FichaTecnica} = {FichaTecnica.Esp_Tec} And {FichaTecnica.Origen} " & VTexto
            '            'Bodega
            '            ElseIf OptInvOpc.Item(1).Value = True Then
            '                GTituloReporte = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Ficha Tecnica " & TxtInv.Text & " " & LblInvDesCod.Caption & "'"
            '                GCriteriaReporte = "{DetalleEntradasInventario.FichaTecnica} = {FichaTecnica.Esp_Tec} And {FichaTecnica.Origen} " & VTexto & " And {DetalleEntradasInventario.Bodega} " & VTexto2
            '            End If
            'ORDEN
            ElseIf OptInv.Item(6).Value = True Then
                        'TODOS
                        If OptInvOpc.Item(0).Value = True Then
                            'TITULO
                            GTituloReporte = "Todas Las Bodegas y Codigo " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "UCase({DetalleEntradasInventario.OrdenProduccion}) " & VTexto
                        'Bodega
                        ElseIf OptInvOpc.Item(1).Value = True Then
                            GTituloReporte = "Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Codigo " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "Ucase({DetalleEntradasInventario.OrdenProduccion}) " & VTexto & " AND Ucase({DetalleEntradasInventario.Bodega}) " & VTexto2
                        ' Descripcion Bodega
                        ElseIf OptInvOpc.Item(2).Value = True Then
                            GTituloReporte = "Por Descripcion Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Codigo " & TxtInv.Text & " " & LblInvDesCod.Caption
                            GCriteriaReporte = "Ucase({DetalleEntradasInventario.OrdenProduccion}) " & VTexto & " AND Ucase({DetalleEntradasInventario.Bodega}) = UCase({BodegasInventario.CodigoBodega}) And UCase({BodegasInventario.Descripcion}) " & VTexto2
                        End If
            'PASILLO
            'ElseIf OptInv.Item(7).Value = True Then
            '            'TODOS
            '            If OptInvOpc.Item(0).Value = True Then
            '                'TITULO
            '                GTituloReporte = "Texto = 'Todas Las Bodegas y Pasillo" & TxtInv.Text & " " & LblInvDesCod.Caption & "'"
            '                GCriteriaReporte = "{DetalleEntradasInventario.Pasillo} " & VTexto
            '            'Bodega
            '            ElseIf OptInvOpc.Item(1).Value = True Then
            '                GTituloReporte = "Texto = 'Por Bodega " & TxtInvOpc.Text & " " & LblInv2.Caption & " Y Pasillo" & TxtInv.Text & " " & LblInvDesCod.Caption & "'"
            '                GCriteriaReporte = "{DetalleEntradasInventario.Pasillo} " & VTexto & " AND {DetalleEntradasInventario.Bodega} " & VTexto2
            '            End If
            
            
            End If 'FIN DE IF DE CODIGO Y DESCRIPCION
            
                            'ORDENAR POR
                            ''CrReportes.SortFields(0) = "+{DetalleEntradasInventario.Batch}"
                            ''CrReportes.SortFields(1) = "+{DetalleEntradasInventario.Tarima}"
                            ''CrReportes.SortFields(2) = "+{DetalleEntradasInventario.FechaProduccion}"
                            
                            'TIPO DE REPORTE CON EXISTENCIA MAYOR QUE CERO
                            If OptTipRep.Item(0).Value = True Then
                                    GCriteriaReporte = GCriteriaReporte & " And {DetalleEntradasInventario.Saldo} > 0"
                            'MENOR O IGUAL QUE CERO
                            ElseIf OptTipRep.Item(1).Value = True Then
                                    GCriteriaReporte = GCriteriaReporte & " And {DetalleEntradasInventario.Saldo} <= 0"
                            'TODOS
                            ElseIf OptTipRep.Item(2).Value = True Then
                            
                            'MAYOR O IGUAL QUE CERO
                            ElseIf OptTipRep.Item(3).Value = True Then
                                    GCriteriaReporte = GCriteriaReporte & " And {DetalleEntradasInventario.Saldo} >= 0"
                            End If

                            'SI ES MAYOR QUE CERO
                            If OptTipRep.Item(0).Value = True Then
                            Else
                                'TITULO
                                GTituloReporte = GTituloReporte & "Existencia A Cero"
                            End If
                            
                            
                            'TIPO DE BODEGA _____________________________________________
                                'TODAS LAS BODEGAS
                                If OptInvPro.Item(0).Value = True Then
                                'BODEGAS DE PROCESO
                                ElseIf OptInvPro.Item(1).Value = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        GCriteriaReporte = GCriteriaReporte & " And {BodegasInventario.EsBodegaDeProceso} = true"
                                    Else
                                        GCriteriaReporte = GCriteriaReporte & " And {BodegasInventario.EsBodegaDeProceso} = -1"
                                    End If
                                'BODEGAS DE NO CONFORME
                                ElseIf OptInvPro.Item(2).Value = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        GCriteriaReporte = GCriteriaReporte & " And {BodegasInventario.EsBodegaDeNoConforme} = true"
                                    Else
                                        GCriteriaReporte = GCriteriaReporte & " And {BodegasInventario.EsBodegaDeNoConforme} = -1"
                                    End If
                                'BODEGAS QUE NO ESTEN EN PROCESO NI EN NO CONFORME
                                ElseIf OptInvPro.Item(3).Value = True Then
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        GCriteriaReporte = GCriteriaReporte & " And {BodegasInventario.EsBodegaDeProceso} = false And {BodegasInventario.EsBodegaDeNoConforme} = false"
                                    Else
                                        GCriteriaReporte = GCriteriaReporte & " And {BodegasInventario.EsBodegaDeProceso} = false And {BodegasInventario.EsBodegaDeNoConforme} = -1"
                                    End If
                                End If
                                
                                
                                
                            'TIPO DE INVENTARIO _____________________________________________________________
                                    If OptInvTipBus.Item(3).Value = True Then
                                    
                                    ElseIf OptInvTipBus.Item(4).Value = True Then
                                            GCriteriaReporte = GCriteriaReporte & " And {DetalleEntradasInventario.FichaTecnica} = {FichaTecnica.Esp_Tec} And {FichaTecnica.TipoInventario} = 'MATERIA PRIMA'"
                                            GTituloReporte = GTituloReporte & "     INVENTARIO DE MATERIA PRIMA"
                                    ElseIf OptInvTipBus.Item(5).Value = True Then
                                            GCriteriaReporte = GCriteriaReporte & " And {DetalleEntradasInventario.FichaTecnica} = {FichaTecnica.Esp_Tec} And {FichaTecnica.TipoInventario} = 'PRODUCTO TERMINADO'"
                                            GTituloReporte = GTituloReporte & "     INVENTARIO DE PRODUCTO TERMINADO"
                                    End If
                                    
                            'TIPO DE ANTIGUEDAD_____________________________________________________________
                                    If OptInvFec.Item(0).Value = True Then
                                            VDia = Day(DtpInvFecIni.Value)
                                            VMes = Month(DtpInvFecIni.Value)
                                            VAño = Year(DtpInvFecIni.Value)
                                            VDia2 = Day(DtpInvFecFin.Value)
                                            VMes2 = Month(DtpInvFecFin.Value)
                                            VAño2 = Year(DtpInvFecFin.Value)
                                            GCriteriaReporte = GCriteriaReporte & " And {DetalleEntradasInventario.FechaProduccion} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                                            GTituloReporte = GTituloReporte & "     Desde " & DtpInvFecIni.Value & " Hasta " & DtpInvFecFin.Value
                                    Else
                                            
                                            
                                    End If


                            
                            
            
                            'TIPO DE REPORTE
                            'If OptTipRep.Item(0).Value = True Then
                                'RESUMEN POR BODEGA
                                If OptInvResDet.Item(0).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvResxBodega.rpt"
                                        Else
                                            GNombreReporte = "RepInvResxBodegaO.rpt"
                                        End If
                                    
                                'DETALLADO
                                ElseIf OptInvResDet.Item(1).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvDetxBodega.rpt"
                                        Else
                                            GNombreReporte = "RepInvDetxBodegaO.rpt"
                                        End If
                                'DETALLADO GENERAL
                                ElseIf OptInvResDet.Item(2).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvDetxTipo.rpt"
                                        Else
                                            GNombreReporte = "RepInvDetxTipoO.rpt"
                                        End If
                                'RESUMEN POR CODIGO
                                ElseIf OptInvResDet.Item(3).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvResxTipo.rpt"
                                        Else
                                            GNombreReporte = "RepInvResxTipoO.rpt"
                                        End If
                                'RESUMEN X CODIGO Y BODEGA
                                ElseIf OptInvResDet.Item(4).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvResxCodigoyBodega.rpt"
                                        Else
                                            GNombreReporte = "RepInvResxCodigoyBodegaO.rpt"
                                        End If
                                'RESUMEN POR TIPO
                                
                                ElseIf OptInvResDet.Item(5).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvResxBodegayTipo.rpt"
                                        Else
                                            GNombreReporte = "RepInvResxBodegayTipoO.rpt"
                                        End If
                                'RESUMEN CUADRICULA
                                ElseIf OptInvResDet.Item(6).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvProTerExiResCuadricula.rpt"
                                        Else
                                            GNombreReporte = "RepInvProTerExiResCuadriculaO.rpt"
                                        End If
                                'RESUMEN CUADRICULA X ORDEN
                                ElseIf OptInvResDet.Item(7).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvProTerExiResCuadriculaOrden.rpt"
                                        Else
                                            GNombreReporte = "RepInvProTerExiResCuadriculaOrdenO.rpt"
                                        End If
                                'DETALLE X TIPO Y CODIGO
                                ElseIf OptInvResDet.Item(8).Value = True Then
                                        If GOrigenDeDatos = "AmaproAccess" Then
                                            GNombreReporte = "RepInvDetxTipoyCodigo.rpt"
                                        Else
                                            GNombreReporte = "RepInvDetxTipoyCodigoO.rpt"
                                        End If
                                End If
                                
                            'End If
End Sub

Sub Entradas()

            TxtEntradas.Text = UCase(TxtEntradas.Text)
            TxtEntradas2.Text = UCase(TxtEntradas2.Text)
            
            VDia = Day(DtpEntFecIni.Value)
            VMes = Month(DtpEntFecIni.Value)
            VAño = Year(DtpEntFecIni.Value)
            VDia2 = Day(DtpEntFecFin.Value)
            VMes2 = Month(DtpEntFecFin.Value)
            VAño2 = Year(DtpEntFecFin.Value)
            'FECHAS
            If OptEntradas.Item(0).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpEntFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpEntFecFin.Value, "dd/mm/yyyy")
                 GCriteriaReporte = "{EncabezadoEntradasInventario.FechaEntrada} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            'BATCH Y LINEA
            ElseIf OptEntradas.Item(1).Value = True Then
                 GTituloReporte = "Batch " & TxtEntradas.Text & " Y Linea " & TxtEntradas2.Text
                 GCriteriaReporte = "{EncabezadoEntradasInventario.Batch} = " & TxtEntradas.Text & " And UPPERCASE({EncabezadoEntradasInventario.Linea}) = '" & UCase(TxtEntradas2.Text) & "'"
            'FECHAS Y FICHA TECNICA
            ElseIf OptEntradas.Item(2).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpEntFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpEntFecFin.Value, "dd/mm/yyyy") & " Por Ficha Tecnica " & TxtEntradas.Text & " " & LblEntDes.Caption
                 GCriteriaReporte = "{EncabezadoEntradasInventario.FechaEntrada} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({DetalleEntradasInventario.FichaTecnica}) Like '" & UCase(TxtEntradas.Text) & "*'"
            'FICHA TECNICA
            ElseIf OptEntradas.Item(3).Value = True Then
                 GTituloReporte = "Por Ficha Tecnica " & TxtEntradas.Text & " " & LblEntDes.Caption
                 GCriteriaReporte = "UPPERCASE({DetalleEntradasInventario.FichaTecnica}) Like '" & UCase(TxtEntradas.Text) & "*'"
            'ENTRADAS NO LIBERADAS
            ElseIf OptEntradas.Item(4).Value = True Then
                 GTituloReporte = "Entradas No Liberadas"
                 GCriteriaReporte = "{EncabezadoEntradasInventario.Estado} = 'NO LIBERADA'"
            'FECHAS Y PROVEEDOR
            ElseIf OptEntradas.Item(5).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpEntFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpEntFecFin.Value, "dd/mm/yyyy") & " Por Linea " & TxtEntradas.Text & " " & LblEntDes.Caption
                 GCriteriaReporte = "{EncabezadoEntradasInventario.FechaEntrada} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({EncabezadoEntradasInventario.Proveedor}) Like '" & UCase(TxtEntradas.Text) & "*'"
            End If
            
                'TIPO DE INVENTARIO _____________________________________________________________
                If OptEntTipInv.Item(0).Value = True Then
                                           
                ElseIf OptEntTipInv.Item(1).Value = True Then
                        GCriteriaReporte = GCriteriaReporte & " And UPPERCASE({DetalleEntradasInventario.FichaTecnica}) = UPPERCASE({FichaTecnica.Esp_Tec}) And {FichaTecnica.TipoInventario} = 'MATERIA PRIMA'"
                        GTituloReporte = GTituloReporte & "         DE MATERIA PRIMA"
                ElseIf OptEntTipInv.Item(2).Value = True Then
                        GCriteriaReporte = GCriteriaReporte & " And UPPERCASE({DetalleEntradasInventario.FichaTecnica}) = UPPERCASE({FichaTecnica.Esp_Tec}) And {FichaTecnica.TipoInventario} = 'PRODUCTO TERMINADO'"
                        GTituloReporte = GTituloReporte & "         DE PRODUCTO TERMINADO"
                End If
            
            
                'TIPO DE REPORTE
                If OptEntDetalle.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioEntradas.rpt"
                    Else
                        GNombreReporte = "InventarioEntradasO.rpt"
                    End If
                ElseIf OptEntRes.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioEntradasResumen.rpt"
                    Else
                        GNombreReporte = "InventarioEntradasResumenO.rpt"
                    End If
                ElseIf OptEntCua.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioEntradasCuadricula.rpt"
                    Else
                        GNombreReporte = "InventarioEntradasCuadriculaO.rpt"
                    End If
                End If
End Sub

Sub CierreBulto()
            VDia = Day(DtpCieTarFecIni.Value)
            VMes = Month(DtpCieTarFecIni.Value)
            VAño = Year(DtpCieTarFecIni.Value)
            VDia2 = Day(DtpCieTarFecFin.Value)
            VMes2 = Month(DtpCieTarFecFin.Value)
            VAño2 = Year(DtpCieTarFecFin.Value)
            'FECHAS
            If OptCierreTarimas.Item(0).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpCieTarFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCieTarFecFin.Value, "dd/mm/yyyy")
                 GCriteriaReporte = "{CierreBulto.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            'FECHAS Y LINEA
            ElseIf OptCierreTarimas.Item(1).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpCieTarFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCieTarFecFin.Value, "dd/mm/yyyy") & " Y Linea " & TxtCieTar.Text & " " & LblCieTarDes.Caption
                 GCriteriaReporte = "{CierreBulto.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({CierreBulto.Linea}) Like '" & UCase(TxtCieTar.Text) & "*'"
            'FECHAS Y FICHA TECNICA
            ElseIf OptCierreTarimas.Item(2).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpCieTarFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCieTarFecFin.Value, "dd/mm/yyyy") & " Y Ficha Tecnica " & TxtCieTar.Text & " " & LblCieTarDes.Caption
                 GCriteriaReporte = "{CierreBulto.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({CierreBulto.FichaTecnica}) Like '" & UCase(TxtCieTar.Text) & "*'"
            'FECHAS Y ESCRIPCION
            ElseIf OptCierreTarimas.Item(3).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpCieTarFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCieTarFecFin.Value, "dd/mm/yyyy") & " Y Ficha Tecnica " & TxtCieTar.Text & " " & LblCieTarDes.Caption
                 GCriteriaReporte = "{CierreBulto.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And uppercase({FichaTecnica.Descrip}) Like '*" & UCase(TxtCieTar.Text) & "*'"
            'FECHAS Y TIPO
            ElseIf OptCierreTarimas.Item(4).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpCieTarFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpCieTarFecFin.Value, "dd/mm/yyyy") & " Y Ficha Tecnica " & TxtCieTar.Text & " " & LblCieTarDes.Caption
                 GCriteriaReporte = "{CierreBulto.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({FichaTecnica.Tipo}) Like '*" & UCase(TxtCieTar.Text) & "*'"
            End If
            
                'DETALLE
                If OptCieTarDet.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "CierreBulto.rpt"
                    Else
                        GNombreReporte = "CierreBultoO.rpt"
                    End If
                'RESUMEN
                Else
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "CierreBultoResumen.rpt"
                    Else
                        GNombreReporte = "CierreBultoResumenO.rpt"
                    End If
                End If
            
          
End Sub

Sub Despachos()
            TxtSalidas.Text = UCase(TxtSalidas.Text)
            TxtSalidas2.Text = UCase(TxtSalidas2.Text)
            
            VTexto = " Like '" & TxtSalidas.Text & "*'"
            VTexto2 = " Like '" & TxtSalidas2.Text & "*'"
                                    
            
            VDia = Day(DtpSalFecIni.Value)
            VMes = Month(DtpSalFecIni.Value)
            VAño = Year(DtpSalFecIni.Value)
            VDia2 = Day(DtpSalFecFin.Value)
            VMes2 = Month(DtpSalFecFin.Value)
            VAño2 = Year(DtpSalFecFin.Value)
            
            'FECHAS
            If OptSalidas.Item(0).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpSalFecFin.Value, "dd/mm/yyyy")
                 GCriteriaReporte = "{EncabezadoSalidasInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            'FECHAS Y CLIENTE Y FICHA TECNICA
            ElseIf OptSalidas.Item(1).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpSalFecFin.Value, "dd/mm/yyyy") & " Por Cliente " & TxtSalidas.Text & " " & LblSalDes.Caption
                 GCriteriaReporte = "{EncabezadoSalidasInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({EncabezadoSalidasInventario.Cliente})" & UCase(VTexto) & " And UPPERCASE({DetalleSalidasInventario.FichaTecnica}" & UCase(VTexto2)
            'FECHAS Y CLIENTE Y DESCRIPCION
            ElseIf OptSalidas.Item(8).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpSalFecFin.Value, "dd/mm/yyyy") & " Por Cliente " & TxtSalidas.Text & " " & LblSalDes.Caption
                 GCriteriaReporte = "{EncabezadoSalidasInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({EncabezadoSalidasInventario.Cliente})" & UCase(VTexto) & " And {EncabezadoSalidasInventario.Documento} = {DetalleSalidasInventario.Documento} And UPPERCASE({DetalleSalidasInventario.FichaTecnica}) = UPPERCASE({FichaTecnica.Esp_Tec}) And UPPERCASE({FichaTecnica.Descrip})" & UCase(VTexto2)
            'FECHAS Y CLIENTE Y TIPO DE FICHA TECNICA
            ElseIf OptSalidas.Item(7).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpSalFecFin.Value, "dd/mm/yyyy") & " Por Cliente " & TxtSalidas.Text & " " & LblSalDes.Caption
                 GCriteriaReporte = "{EncabezadoSalidasInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({EncabezadoSalidasInventario.Cliente})" & UCase(VTexto) & " And UPPERCASE({DetalleSalidasInventario.FichaTecnica}) = UPPERCASE({FichaTecnica.Esp_Tec}) And UPPERCASE({FichaTecnica.Tipo})" & UCase(VTexto2)
            'FECHAS Y TRANSPORTISTA
            ElseIf OptSalidas.Item(2).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpSalFecFin.Value, "dd/mm/yyyy") & " Por Transportista " & TxtSalidas.Text & " " & LblSalDes.Caption
                 GCriteriaReporte = "{EncabezadoSalidasInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({EncabezadoSalidasInventario.CodigoTransportista})" & UCase(VTexto)
            'FECHAS Y FICHA TECNICA
            ElseIf OptSalidas.Item(3).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpSalFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpSalFecFin.Value, "dd/mm/yyyy") & " Por Ficha Tecnica " & TxtSalidas.Text & " " & LblSalDes.Caption
                 GCriteriaReporte = "{EncabezadoSalidasInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({DetalleSalidasInventario.FichaTecnica})" & UCase(VTexto2)
            'FICHA TECNICA
            ElseIf OptSalidas.Item(4).Value = True Then
                 GTituloReporte = "Por Ficha Tecnica " & TxtSalidas2.Text & " " & LblSalDes2.Caption
                 GCriteriaReporte = "UPPERCASE({DetalleSalidasInventario.FichaTecnica})" & UCase(VTexto2)
            'BATCH Y LINEA
            ElseIf OptSalidas.Item(5).Value = True Then
                 GTituloReporte = "Por Batch " & TxtSalidas.Text & " Y Linea " & TxtSalidas2.Text
                 GCriteriaReporte = "{DetalleSalidasInventario.Batch} = " & TxtSalidas.Text & " And UPPERCASE({DetalleSalidasInventario.Linea}) = '" & UCase(TxtSalidas2.Text) & "'"
            'NO LIBERADO
            ElseIf OptSalidas.Item(6).Value = True Then
                 GTituloReporte = "Despachos No Liberados"
                 GCriteriaReporte = "{EncabezadoSalidasInventario.Estado} = 'NO LIBERADA'"
            End If
            
            'TIPO DE INVENTARIO _____________________________________________________________
                If OptSalTipInv.Item(0).Value = True Then
                                           
                ElseIf OptSalTipInv.Item(1).Value = True Then
                        GCriteriaReporte = GCriteriaReporte & " And UPPERCASE({DetalleSalidasInventario.FichaTecnica}) = UPPERCASE({FichaTecnica.Esp_Tec}) And {FichaTecnica.TipoInventario} = 'MATERIA PRIMA'"
                        GTituloReporte = GTituloReporte & "         DE MATERIA PRIMA"
                ElseIf OptSalTipInv.Item(2).Value = True Then
                        GCriteriaReporte = GCriteriaReporte & " And UPPERCASE({DetalleSalidasInventario.FichaTecnica}) = UPPERCASE({FichaTecnica.Esp_Tec}) And {FichaTecnica.TipoInventario} = 'PRODUCTO TERMINADO'"
                        GTituloReporte = GTituloReporte & "         DE PRODUCTO TERMINADO"
                End If
            
                'TIPO DE REPORTE
                If OptSalDet.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioSalidas.rpt"
                    Else
                        GNombreReporte = "InventarioSalidasO.rpt"
                    End If
                    'CrReportes.SubreportToChange = "SalidasMateriaPrima"
                    'CrReportes.ConnectionString = "pwd=metal"
                    'CrReportes.SubreportToChange = ""
                ElseIf OptSalRes.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioSalidasResumen.rpt"
                    Else
                        GNombreReporte = "InventarioSalidasResumenO.rpt"
                    End If
                ElseIf OptSalResCli.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioSalidasResumenCliente.rpt"
                    Else
                        GNombreReporte = "InventarioSalidasResumenClienteO.rpt"
                    End If
                ElseIf OptSalCua.Value = True Then
                    If GOrigenDeDatos = "AmaproAccess" Then
                        GNombreReporte = "InventarioSalidasCuadriculaMes.rpt"
                    Else
                        GNombreReporte = "InventarioSalidasCuadriculaMesO.rpt"
                    End If
                End If

End Sub


Sub Desperdicio()
            VDia = Day(DTPDesFecIni.Value)
            VMes = Month(DTPDesFecIni.Value)
            VAño = Year(DTPDesFecIni.Value)
            VDia2 = Day(DTPDesFecFin.Value)
            VMes2 = Month(DTPDesFecFin.Value)
            VAño2 = Year(DTPDesFecFin.Value)
                        
                    'FECHAS
                    If OptDesperdicio.Item(0).Value = True Then
                         GTituloReporte = "Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy")
                         GCriteriaReporte = "{CapturaDesperdicio.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                    'FECHAS Y PROCESO
                    ElseIf OptDesperdicio.Item(1).Value = True Then
                         GTituloReporte = "Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy") & " Del Proceso " & TxtDesperdicio.Text & " " & LblDesDes.Caption
                         GCriteriaReporte = "{CapturaDesperdicio.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({CapturaDesperdicio.CodigoProceso}) like '" & UCase(TxtDesperdicio.Text) & "*'"
                    'FECHAS Y FICHA TECNICA
                    ElseIf OptDesperdicio.Item(2).Value = True Then
                         GTituloReporte = "Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy") & " De Ficha Tecnica " & TxtDesperdicio.Text & " " & LblDesDes.Caption
                         GCriteriaReporte = "{CapturaDesperdicio.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({CapturaDesperdicio.FichaTecnica}) Like '" & UCase(TxtDesperdicio.Text) & "*'"
                    'FECHAS Y Grupo
                    ElseIf OptDesperdicio.Item(3).Value = True Then
                         GTituloReporte = "Desde " & Format(DTPDesFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DTPDesFecFin.Value, "dd/mm/yyyy") & " Del Grupo " & TxtDesperdicio.Text & " " & LblDesDes.Caption
                         GCriteriaReporte = "{CapturaDesperdicio.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And UPPERCASE({CapturaDesperdicio.CodigoProceso}) = UPPERCASE({ProcesosMateriaPrima.CodigoProceso})  And UPPERCASE({ProcesosMateriaPrima.Grupo}) like '" & UCase(TxtDesperdicio.Text) & "*'"
                    End If
            
                
                    If OptDetalle.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "DesperdicioProcesoMateriaPrimaDetalle.rpt"
                        Else
                            GNombreReporte = "DesperdicioProcesoMateriaPrimaDetalleO.rpt"
                        End If
                    ElseIf OptResumen.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "DesperdicioProcesoMateriaPrimaResumen.rpt"
                        Else
                            GNombreReporte = "DesperdicioProcesoMateriaPrimaResumenO.rpt"
                        End If
                    'CUADRICULA
                    ElseIf OptDesCuaPro.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "DesperdicioProcesoMateriaPrimaCuadricula.rpt"
                        Else
                            GNombreReporte = "DesperdicioProcesoMateriaPrimaCuadriculaO.rpt"
                        End If
                    ElseIf OptDesResDef.Value = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            GNombreReporte = "DesperdicioProcesoMateriaPrimaResumenDefectos.rpt"
                        Else
                            GNombreReporte = "DesperdicioProcesoMateriaPrimaResumenDefectosO.rpt"
                        End If
                    End If
                


End Sub

Private Sub TxtSalidas2_Change()
        'BUSCA FICHA TECNICA
        If (OptSalidas.Item(1).Value = True Or OptSalidas.Item(3).Value = True Or OptSalidas.Item(4).Value = True) Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtSalidas2.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtSalidas2.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblSalDes2.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblSalDes2.Caption = ""
                End If
        'BUSCA TIPO DE FICHA TECNICA
        ElseIf OptSalidas.Item(7).Value = True Then
            Set RBuscaTipo = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicaTipos Where CodigoTipo = '" & TxtSalidas2.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicaTipos Where UPPER(CodigoTipo) = '" & UCase(TxtSalidas2.Text) & "'")
                End If
                If RBuscaTipo.RecordCount > 0 Then
                    LblSalDes2.Caption = RBuscaTipo!Descripcion
                Else
                    LblSalDes2.Caption = ""
                End If
        'LINEA
        ElseIf OptSalidas.Item(5).Value = True Then
            Set RBuscaLinea = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtSalidas2.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtSalidas2.Text) & "'")
                End If
                If RBuscaLinea.RecordCount > 0 Then
                    LblSalDes2.Caption = RBuscaLinea!Descrip
                Else
                    LblSalDes2.Caption = ""
                End If
        

        End If

End Sub

Private Sub TxtSalidas2_DblClick()
        'OPCION POR FICHA TECNICA
        If (OptSalidas.Item(3).Value = True Or OptSalidas.Item(4).Value = True Or OptSalidas.Item(1).Value = True) Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = True
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulFichaTecnica = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False
        'TIPO DE FICHA TECNICA
        ElseIf OptSalidas.Item(7).Value = True Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = False
                    BSalTipoFT = True
                    'CIERRE TARIMA
                    BCieBulFichaTecnica = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False
        End If
        
        Set RBusqueda = New ADODB.Recordset
        'OPCION DE FICHAS TECNICAS
        If BSalFichaTecnica = True Then
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, Envases, Origen From FichaTecnica")
        'TIPO DE FICHA TECNICA
        ElseIf BSalTipoFT = True Then
                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos")
        End If
        
        If (BSalFichaTecnica = True Or BSalTipoFT = True) Then
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtSalidas2.SetFocus
        End If

End Sub

Private Sub TxtSalidas2_KeyPress(KeyAscii As Integer)
    'SI PRECIONA LA TECLA ENTER
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
    
        
       
    If KeyAscii = 43 Then
        'OPCION POR FICHA TECNICA
        If (OptSalidas.Item(3).Value = True Or OptSalidas.Item(4).Value = True Or OptSalidas.Item(1).Value = True) Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = True
                    BSalTipoFT = False
                    'CIERRE TARIMA
                    BCieBulFichaTecnica = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False
        'TIPO DE FICHA TECNICA
        ElseIf OptSalidas.Item(7).Value = True Then
                    'INVENTARIO
                    BInvBodega = False
                    BInvCatalogo = False
                    BInvFichaTecnica = False
                    'ENTRADAS
                    BEntFichaTecnica = False
                    'DESPACHOS
                    BSalCliente = False
                    BSalTransportista = False
                    BSalFichaTecnica = False
                    BSalTipoFT = True
                    'CIERRE TARIMA
                    BCieBulFichaTecnica = False
                    'TRASLADOS
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'DESPERDICIO
                    BDesProceso = False
                    BDesFichaTecnica = False
        End If
        
        Set RBusqueda = New ADODB.Recordset
        
        'OPCION DE FICHAS TECNICAS
        If BSalFichaTecnica = True Then
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, Envases, Origen From FichaTecnica")
        'TIPO DE FICHA TECNICA
        ElseIf BSalTipoFT = True Then
                Call Abrir_Recordset(RBusqueda, "Select CodigoTipo, Descripcion From FichaTecnicaTipos")
        End If
        
        If (BSalFichaTecnica = True Or BSalTipoFT = True) Then
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "4000"
                FrameBusqueda.Visible = True
                TxtSalidas2.SetFocus
        End If
        
    End If

End Sub

Public Sub Traslados()
            VDia = Day(DtpTraFecIni.Value)
            VMes = Month(DtpTraFecIni.Value)
            VAño = Year(DtpTraFecIni.Value)
            VDia2 = Day(DtpTraFecFin.Value)
            VMes2 = Month(DtpTraFecFin.Value)
            VAño2 = Year(DtpTraFecFin.Value)
            'FECHAS
            If OptTraslados.Item(0).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy")
                 GCriteriaReporte = "{EncabezadoTrasladosInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
                 'gcriteriareporte = "{EncabezadoDevolucionesMateriaPrima.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ")"
            'NUMERO DOCUMENTO
            ElseIf OptTraslados.Item(1).Value = True Then
                    GTituloReporte = "Numero De Documento " & TxtTraslados.Text
                    GCriteriaReporte = "ucase({EncabezadoTrasladosInventario.NumeroDocumento}) = '" & UCase(TxtTraslados.Text) & "'"
            'FECHAS Y BODEGA SALIDA Y CODIGO DE MATERIA PRIMA
            ElseIf OptTraslados.Item(2).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Bodega Salida " & TxtTraslados.Text & " " & LblTraBod.Caption
                 GCriteriaReporte = "{EncabezadoTrasladosInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And ucase({DetalleTrasladosInventario.FichaTecnica}) Like '" & UCase(TxtTraslados.Text) & "*' And ucase({EncabezadoTrasladosInventario.BodegaSalida}) Like '" & UCase(TxtTraslados2.Text) & "*'"
            'FECHAS BODEGA ENTRADA Y CODIGO DE MATERIA PRIMA
            ElseIf OptTraslados.Item(3).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Bodega Entrada " & TxtTraslados.Text & " " & LblTraBod.Caption
                 GCriteriaReporte = "{EncabezadoTrasladosInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And ucase({DetalleTrasladosInventario.FichaTecnica}) Like '" & UCase(TxtTraslados.Text) & "*' And ucase({DetalleTrasladosInventario.BodegaEntrada}) Like '" & UCase(TxtTraslados2.Text) & "*'"
            'FECHAS Y BODEGA SALIDA Y TIPO DE MATERIA PRIMA
            ElseIf OptTraslados.Item(4).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Bodega Salida " & TxtTraslados.Text & " " & LblTraBod.Caption & " Materia Prima " & LblTraBod2.Caption
                 GCriteriaReporte = "{EncabezadoTrasladosInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And ucase({EncabezadoTrasladosInventario.BodegaSalida}) Like '" & UCase(TxtTraslados.Text) & "*' And ucase({FichaTecnica.Tipo}) Like '" & UCase(TxtTraslados2.Text) & "*'"
            'FECHAS BODEGA ENTRADA Y TIPO DE MATERIA PRIMA
            ElseIf OptTraslados.Item(5).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Bodega Entrada " & TxtTraslados.Text & " " & LblTraBod.Caption & " Materia Prima " & LblTraBod2.Caption
                 GCriteriaReporte = "{EncabezadoTrasladosInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And ucase({DetalleTrasladosInventario.BodegaEntrada}) Like '" & UCase(TxtTraslados.Text) & "*' And ucase({FichaTecnica.Tipo}) Like '" & UCase(TxtTraslados2.Text) & "*'"
            'FECHAS CODIGO MATERIA PRIMA
            ElseIf OptTraslados.Item(6).Value = True Then
                 GTituloReporte = "Desde " & Format(DtpTraFecIni.Value, "dd/mm/yyyy") & " Hasta " & Format(DtpTraFecFin.Value, "dd/mm/yyyy") & " Y Codigo " & TxtTraslados.Text & " " & LblTraBod.Caption
                 GCriteriaReporte = "{EncabezadoTrasladosInventario.Fecha} in date (" & VAño & "," & VMes & "," & VDia & ") to date (" & VAño2 & "," & VMes2 & "," & VDia2 & ") And ucase({DetalleTrasladosInventario.FichaTecnica}) Like '" & UCase(TxtTraslados.Text) & "*'"
            'ORDEN
            ElseIf OptTraslados.Item(7).Value = True Then
                 GTituloReporte = "Orden " & TxtTraslados.Text
                 GCriteriaReporte = "ucase({DetalleTrasladosInventario.Orden}) = '" & UCase(TxtTraslados.Text) & "'"
            'MATERIA PRIMA
            ElseIf OptTraslados.Item(8).Value = True Then
                 GTituloReporte = "Codigo Materia Prima " & TxtTraslados.Text & " " & LblTraBod.Caption
                 GCriteriaReporte = "ucase({DetalleTrasladosInventario.FichaTecnica}) Like '" & UCase(TxtTraslados.Text) & "*'"
            'NO LIBERADO
            ElseIf OptTraslados.Item(9).Value = True Then
                 GTituloReporte = "Traslados No Liberados"
                 GCriteriaReporte = "{EncabezadoTrasladosInventario.Estado} = 'NO LIBERADO'"
            End If
            
            'OPCION DE TODOS LOS TIPOS DE DOCUMENTO
            If OptTraOpc.Item(0).Value = True Then
                'POR UN TIPO DE DOCUMENTO
            Else
                GCriteriaReporte = GCriteriaReporte & " And ucase({EncabezadoTrasladosInventario.TipoDeDocumento}) = '" & UCase(TxtTraTipDoc.Text) & "'"
            End If
        
            
            If OptTraDet.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "InventarioTrasladosDetalle.rpt"
                Else
                    GNombreReporte = "InventarioTrasladosDetalleO.rpt"
                End If
            Else
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "InventarioTrasladosResumen.rpt"
                Else
                    GNombreReporte = "InventarioTrasladosResumenO.rpt"
                End If
            End If
            'gnombrereporte = App.Path & "\ReporteDevolucionesMateriaPrima.rpt"

End Sub

Private Sub TxtTraslados_Change()
        'BUSCA LA BODEGA SI LA OPCION DE BODEGA DE SALIDA O ENTRADA SON VERDADERAS
        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
            Set RBuscaBodega = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtTraslados.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtTraslados.Text) & "'")
                End If
                If RBuscaBodega.RecordCount > 0 Then
                    LblTraBod.Caption = RBuscaBodega!Descripcion
                Else
                    LblTraBod.Caption = ""
                End If
        'BUSCA CODIGO DE MATERIA PRIMA
        ElseIf (OptTraslados.Item(4).Value = True Or OptTraslados.Item(6).Value = True Or OptTraslados.Item(2).Value = True Or OptTraslados.Item(3).Value = True) Then
            Set RBuscaFichaTecnica = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtTraslados.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaFichaTecnica, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtTraslados.Text) & "'")
                End If
                If RBuscaFichaTecnica.RecordCount > 0 Then
                    LblTraBod.Caption = RBuscaFichaTecnica!Descrip
                Else
                    LblTraBod.Caption = ""
                End If
        Else
                    LblTraBod.Caption = ""
        End If

End Sub

Private Sub TxtTraslados_DblClick()
        Set RBusqueda = New ADODB.Recordset
        'OPCION POR BODEGA
        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
                    BInvBodega = False
                    BInvTipo = False
                    BInvFichaTecnica = False
                    BInvBodegaGrupo = False
                    BTraBodega = True
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    'CIERRE TARIMA
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BCieBulTipo = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalFichaTecnica = False
                    BSalCliente = False
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipo = False
                    BCieBulTipo = False
                    
        'OPCION POR MATERIA PRIMA
        ElseIf (OptTraslados.Item(2).Value = True Or OptTraslados.Item(3).Value = True Or OptTraslados.Item(6).Value = True Or OptTraslados.Item(8).Value = True) Then
                    BInvBodega = False
                    BInvTipo = False
                    BInvFichaTecnica = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraFichaTecnica = True
                    BTraDocumentos = False
                    'CIERRE TARIMA
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BCieBulTipo = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalFichaTecnica = False
                    BSalCliente = False
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipo = False
                    BCieBulTipo = False
                    
        End If
    
        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
        If BTraBodega = True Then
                Call Abrir_Recordset(RBusqueda, "Select * From BodegasInventario")
        'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
        ElseIf BTraFichaTecnica = True Then
                Call Abrir_Recordset(RBusqueda, "Select * From FichaTecnica")
        End If
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "3000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus

End Sub

Private Sub TxtTraslados_GotFocus()
            TxtTraslados.SelStart = 0
            TxtTraslados.SelLength = Len(TxtTraslados.Text)
End Sub

Private Sub TxtTraslados_KeyPress(KeyAscii As Integer)

                If KeyAscii = 13 Then
                    SendKeys "{tab}"
                End If

                If KeyAscii = 43 Then
                        Set RBusqueda = New ADODB.Recordset
                        'OPCION POR BODEGA
                        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
                                    BInvBodega = False
                                    BInvTipo = False
                                    BInvFichaTecnica = False
                                    BInvBodegaGrupo = False
                                    BTraBodega = True
                                    BTraFichaTecnica = False
                                    BTraDocumentos = False
                                    'CIERRE TARIMA
                                    BCieBulLinea = False
                                    BCieBulFichaTecnica = False
                                    BCieBulTipo = False
                                    BEntProveedor = False
                                    BEntMateriaPrima = False
                                    BEntTipoMateriaPrima = False
                                    BSalFichaTecnica = False
                                    BSalCliente = False
                                    BCieBulLinea = False
                                    BCieBulFichaTecnica = False
                                    BDesProceso = False
                                    BDesFichaTecnica = False
                                    BTraTipo = False
                                    BCieBulTipo = False
                                    
                        'OPCION POR MATERIA PRIMA
                        ElseIf (OptTraslados.Item(2).Value = True Or OptTraslados.Item(3).Value = True Or OptTraslados.Item(6).Value = True Or OptTraslados.Item(8).Value = True) Then
                                    BInvBodega = False
                                    BInvTipo = False
                                    BInvFichaTecnica = False
                                    BInvBodegaGrupo = False
                                    BTraBodega = False
                                    BTraFichaTecnica = True
                                    BTraDocumentos = False
                                    'CIERRE TARIMA
                                    BCieBulLinea = False
                                    BCieBulFichaTecnica = False
                                    BCieBulTipo = False
                                    BEntProveedor = False
                                    BEntMateriaPrima = False
                                    BEntTipoMateriaPrima = False
                                    BSalFichaTecnica = False
                                    BSalCliente = False
                                    BCieBulLinea = False
                                    BCieBulFichaTecnica = False
                                    BDesProceso = False
                                    BDesFichaTecnica = False
                                    BTraTipo = False
                                    BCieBulTipo = False
                                    
                        End If
                    
                        'OPCION DE BODEGA EN CARPETA DE TRASLADOS
                        If BTraBodega = True Then
                                Call Abrir_Recordset(RBusqueda, "Select * From BodegasInventario")
                        'OPCION DE MATERIA PRIMA EN CARPETA DE TRASLADOS
                        ElseIf BTraFichaTecnica = True Then
                                Call Abrir_Recordset(RBusqueda, "Select * From FichaTecnica")
                        End If
                                Set DBGridBusqueda.DataSource = RBusqueda
                                DBGridBusqueda.Columns(1).Width = "3000"
                                FrameBusqueda.Visible = True
                                TxtBusqueda.SetFocus

                End If

End Sub

Private Sub TxtTraslados2_Change()
    If (OptTraslados.Item(2).Value = True Or OptTraslados.Item(3).Value = True) Then
        Set RBuscaBodega = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtTraslados2.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtTraslados2.Text) & "'")
            End If
                        If RBuscaBodega.RecordCount > 0 Then
                            LblTraBod2.Caption = RBuscaBodega!Descripcion
                        Else
                            LblTraBod2.Caption = ""
                        End If
    
    Else
            'BUSCA TIPO DE MATERIA PRIMA
            Set RBuscaTipo = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicaTipos Where CodigoTipo = '" & TxtTraslados2.Text & "'")
                Else
                    Call Abrir_Recordset(RBuscaTipo, "Select Descripcion From FichaTecnicaTipos Where UPPER(CodigoTipo) = '" & UCase(TxtTraslados2.Text) & "'")
                End If
                If RBuscaTipo.RecordCount > 0 Then
                    LblTraBod2.Caption = RBuscaTipo!Descripcion
                Else
                    LblTraBod2.Caption = ""
                End If
    End If
        

End Sub

Private Sub TxtTraslados2_DblClick()
  'OPCION POR TIPO DE MATERIA PRIMA
        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
                    Set RBusqueda = New ADODB.Recordset
                    BInvBodega = False
                    BInvTipo = False
                    BInvFichaTecnica = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalFichaTecnica = False
                    BSalCliente = False
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipo = True
                    BCieBulTipo = False
                                        
                    Call Abrir_Recordset(RBusqueda, "Select * From FichaTecnicaTipos")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If

End Sub

Private Sub TxtTraslados2_KeyPress(KeyAscii As Integer)
        
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    End If
        
    If KeyAscii = 43 Then
  'OPCION POR TIPO DE MATERIA PRIMA
        If (OptTraslados.Item(4).Value = True Or OptTraslados.Item(5).Value = True) Then
                    Set RBusqueda = New ADODB.Recordset
                    BInvBodega = False
                    BInvTipo = False
                    BInvFichaTecnica = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = False
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalFichaTecnica = False
                    BSalCliente = False
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipo = True
                    BCieBulTipo = False
                                        
                    Call Abrir_Recordset(RBusqueda, "Select * From FichaTecnicaTipos")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If
    End If

End Sub

Private Sub TxtTraTipDoc_Change()
        Set RBuscaDocumento = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaDocumento, "Select Descripcion From Documentos Where CodigoDocumento = '" & TxtTraTipDoc.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaDocumento, "Select Descripcion From Documentos Where UPPER(CodigoDocumento) = '" & UCase(TxtTraTipDoc.Text) & "'")
            End If
            If RBuscaDocumento.RecordCount > 0 Then
                LblTraDesDoc.Caption = RBuscaDocumento!Descripcion
            Else
                LblTraDesDoc.Caption = ""
            End If

End Sub

Private Sub TxtTraTipDoc_DblClick()
                    Set RBusqueda = New ADODB.Recordset
                    BInvBodega = False
                    BInvTipo = False
                    BInvFichaTecnica = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = True
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalFichaTecnica = False
                    BSalCliente = False
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipo = False
                    BCieBulTipo = False
    
                    Call Abrir_Recordset(RBusqueda, "Select * From Documentos")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus

End Sub

Private Sub TxtTraTipDoc_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
        If KeyAscii = 43 Then
                    Set RBusqueda = New ADODB.Recordset
                    BInvBodega = False
                    BInvTipo = False
                    BInvFichaTecnica = False
                    BInvBodegaGrupo = False
                    BTraBodega = False
                    BTraFichaTecnica = False
                    BTraDocumentos = True
                    BEntProveedor = False
                    BEntMateriaPrima = False
                    BEntTipoMateriaPrima = False
                    BSalFichaTecnica = False
                    BSalCliente = False
                    BCieBulLinea = False
                    BCieBulFichaTecnica = False
                    BDesProceso = False
                    BDesFichaTecnica = False
                    BTraTipo = False
                    BCieBulTipo = False
    
                    Call Abrir_Recordset(RBusqueda, "Select * From Documentos")
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "3000"
                    FrameBusqueda.Visible = True
                    TxtBusqueda.SetFocus
        End If
End Sub
