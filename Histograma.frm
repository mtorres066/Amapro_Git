VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form GeneraHistograma 
   Caption         =   "Histograma De Captura De Rutinas"
   ClientHeight    =   8340
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Histograma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
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
      Height          =   8175
      Left            =   600
      TabIndex        =   54
      Top             =   5520
      Visible         =   0   'False
      Width           =   11775
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   6975
         Left            =   120
         TabIndex        =   57
         Top             =   1080
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12303
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
         Caption         =   "Descripcion"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   59
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptBusqueda 
         Caption         =   "Codigo"
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   1680
         TabIndex        =   58
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBusqueda 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   56
         ToolTipText     =   "digite los datos a buscar"
         Top             =   720
         Width           =   4455
      End
      Begin VB.CommandButton CmdSale 
         Height          =   615
         Left            =   10560
         Picture         =   "Histograma.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Sale De Busqueda"
         Top             =   360
         Width           =   615
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   14631
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   1058
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "Histograma.frx":237C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbltitulo"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "LblRutina"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "LblFichaTecnica"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Shape1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label15"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label13"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label12"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label9"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label6"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label5"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label16"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "LblLinea"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "LblLineas"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "ImageFoto"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblcab"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "LblCab2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "FGridHistograma"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "PFecRutFin"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "PFecRutIni"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "TxtRut"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "TxtFicTec"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "CmdGenerar"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "TxtLIC"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "TxtLSC"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "TxtDesviacionInterno"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Txtmediainterno"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Txtb"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "TxtCv"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "TxtDes"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "TxtMed"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "TxtDatMay"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "TxtDatMen"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "TxtInt"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "TxtGru"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "TxtRan"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "TxtCP"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "TxtLin"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "ChkLin"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "DbGridHistograma"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Framecabezal"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "txtcab"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtcab2"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "ChkReporte"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).ControlCount=   52
      TabCaption(1)   =   "Grafica"
      TabPicture(1)   =   "Histograma.frx":27D6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CboVerGra"
      Tab(1).Control(1)=   "CmdImprimirGrafica"
      Tab(1).Control(2)=   "CmdCopiar"
      Tab(1).Control(3)=   "CmdGrabar"
      Tab(1).Control(4)=   "Grafica"
      Tab(1).Control(5)=   "Label2"
      Tab(1).ControlCount=   6
      Begin VB.CheckBox ChkReporte 
         Caption         =   "Imprime Reporte"
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
         Left            =   240
         TabIndex        =   66
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox txtcab2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtcab 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame Framecabezal 
         Caption         =   "Tipo Cabezal"
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
         Left            =   2880
         TabIndex        =   61
         Top             =   840
         Width           =   1455
         Begin VB.OptionButton optcab 
            Caption         =   "Un Cabezal"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   63
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optcab 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   62
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DbGridHistograma 
         Height          =   6015
         Left            =   240
         TabIndex        =   60
         Top             =   2040
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   10610
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
      Begin VB.ComboBox CboVerGra 
         Height          =   315
         ItemData        =   "Histograma.frx":2C28
         Left            =   -68280
         List            =   "Histograma.frx":2C50
         TabIndex        =   52
         Text            =   "3dBar"
         Top             =   600
         Width           =   1935
      End
      Begin VB.CommandButton CmdImprimirGrafica 
         Height          =   375
         Left            =   -64440
         Picture         =   "Histograma.frx":2CC5
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Imprimir Grafica"
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton CmdCopiar 
         Height          =   375
         Left            =   -65400
         Picture         =   "Histograma.frx":31F7
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Copia Grafica"
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton CmdGrabar 
         Height          =   375
         Left            =   -66240
         Picture         =   "Histograma.frx":3729
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Graba Grafica"
         Top             =   600
         Width           =   735
      End
      Begin VB.CheckBox ChkLin 
         Caption         =   "X Linea"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   1800
         Width           =   975
      End
      Begin VB.TextBox TxtLin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TxtCP 
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
         Left            =   10680
         TabIndex        =   24
         Top             =   7800
         Width           =   855
      End
      Begin VB.TextBox TxtRan 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   7080
         Width           =   975
      End
      Begin VB.TextBox TxtGru 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   7440
         Width           =   975
      End
      Begin VB.TextBox TxtInt 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   7800
         Width           =   975
      End
      Begin VB.TextBox TxtDatMen 
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
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   5640
         Width           =   855
      End
      Begin VB.TextBox TxtDatMay 
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
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   6000
         Width           =   855
      End
      Begin VB.TextBox TxtMed 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   6360
         Width           =   975
      End
      Begin VB.TextBox TxtDes 
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
         Left            =   8520
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   6720
         Width           =   975
      End
      Begin VB.TextBox TxtCv 
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
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   7080
         Width           =   855
      End
      Begin VB.TextBox Txtb 
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
         Left            =   10680
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   7440
         Width           =   855
      End
      Begin VB.TextBox Txtmediainterno 
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
         Left            =   10680
         TabIndex        =   22
         Top             =   6360
         Width           =   855
      End
      Begin VB.TextBox TxtDesviacionInterno 
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
         Left            =   10680
         TabIndex        =   21
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox TxtLSC 
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
         Left            =   8520
         TabIndex        =   20
         Top             =   6000
         Width           =   975
      End
      Begin VB.TextBox TxtLIC 
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
         Left            =   8520
         TabIndex        =   19
         Top             =   5640
         Width           =   975
      End
      Begin VB.CommandButton CmdGenerar 
         Caption         =   "Genera Grafica"
         Height          =   975
         Left            =   10680
         Picture         =   "Histograma.frx":3C5B
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Genera Histograma"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox TxtFicTec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtRut 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   9
         Top             =   840
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Caption         =   "Opciones"
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
         Height          =   735
         Left            =   1560
         TabIndex        =   13
         Top             =   840
         Width           =   1215
         Begin VB.OptionButton OptCatalogo 
            Caption         =   "Catalogo"
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptFicha 
            Caption         =   "Ficha"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   975
         End
      End
      Begin MSChart20Lib.MSChart Grafica 
         Height          =   7095
         Left            =   -74880
         OleObjectBlob   =   "Histograma.frx":5CCD
         TabIndex        =   1
         Top             =   1080
         Width           =   11655
      End
      Begin MSComCtl2.DTPicker PFecRutIni 
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58720259
         CurrentDate     =   36919
      End
      Begin MSComCtl2.DTPicker PFecRutFin 
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   58720259
         CurrentDate     =   36919
      End
      Begin MSFlexGridLib.MSFlexGrid FGridHistograma 
         Height          =   3495
         Left            =   7320
         TabIndex        =   16
         Top             =   1920
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   6165
         _Version        =   393216
         Rows            =   5
         Cols            =   5
         BackColor       =   12632319
         ForeColor       =   0
         BackColorFixed  =   12632256
         ForeColorFixed  =   0
         ForeColorSel    =   12632256
         GridColor       =   0
         BorderStyle     =   0
      End
      Begin MSComDlg.CommonDialog CDDialogo 
         Left            =   120
         Top             =   -480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DefaultExt      =   "bmp"
         DialogTitle     =   "Grabar Grafica"
         Filter          =   "Pictures (*.bmp)|*.bmp"
         FilterIndex     =   3
      End
      Begin VB.Label LblCab2 
         AutoSize        =   -1  'True
         Caption         =   "Al"
         Height          =   195
         Left            =   4560
         TabIndex        =   65
         Top             =   1680
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label lblcab 
         AutoSize        =   -1  'True
         Caption         =   "Cabezal Del"
         Height          =   195
         Left            =   3000
         TabIndex        =   64
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Vistas De Grafica"
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
         TabIndex        =   53
         Top             =   720
         Width           =   1575
      End
      Begin VB.Image ImageFoto 
         BorderStyle     =   1  'Fixed Single
         Height          =   7455
         Left            =   120
         Stretch         =   -1  'True
         Top             =   720
         Width           =   11655
      End
      Begin VB.Label LblLineas 
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
         Left            =   6600
         TabIndex        =   48
         Top             =   1560
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label LblLinea 
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
         Left            =   5400
         TabIndex        =   47
         Top             =   1560
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FF8080&
         Caption         =   "CP"
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
         Left            =   9600
         TabIndex        =   46
         Top             =   7800
         Width           =   735
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FF8080&
         Caption         =   "Rango"
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
         Left            =   7440
         TabIndex        =   45
         Top             =   7080
         Width           =   735
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FF8080&
         Caption         =   "Grupos"
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
         Left            =   7440
         TabIndex        =   44
         Top             =   7440
         Width           =   855
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FF8080&
         Caption         =   "Intervalo"
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
         Left            =   7440
         TabIndex        =   43
         Top             =   7800
         Width           =   855
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FF8080&
         Caption         =   "Dato Menor"
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
         Left            =   9600
         TabIndex        =   42
         Top             =   5640
         Width           =   1575
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF8080&
         Caption         =   "Dato Mayor"
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
         Left            =   9600
         TabIndex        =   41
         Top             =   6000
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF8080&
         Caption         =   "Media"
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
         Left            =   7440
         TabIndex        =   40
         Top             =   6360
         Width           =   975
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FF8080&
         Caption         =   "Desviacion"
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
         Left            =   7440
         TabIndex        =   39
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FF8080&
         Caption         =   "CV%"
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
         Left            =   9600
         TabIndex        =   38
         Top             =   7080
         Width           =   615
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FF8080&
         Caption         =   "b"
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
         Left            =   9600
         TabIndex        =   37
         Top             =   7440
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Med. Datos"
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
         Left            =   9600
         TabIndex        =   36
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "Des. Datos"
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
         Left            =   9600
         TabIndex        =   35
         Top             =   6720
         Width           =   975
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FF8080&
         Caption         =   "Lim.Sup.Cli."
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
         Left            =   7440
         TabIndex        =   34
         Top             =   6000
         Width           =   1095
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FF8080&
         Caption         =   "Lim. Inf. Cli."
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
         Left            =   7440
         TabIndex        =   33
         Top             =   5640
         Width           =   1335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FF8080&
         BackStyle       =   1  'Opaque
         Height          =   2655
         Left            =   7320
         Top             =   5520
         Width           =   4335
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
         Left            =   6600
         TabIndex        =   18
         Top             =   1200
         Width           =   3855
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
         Height          =   255
         Left            =   6600
         TabIndex        =   17
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label lbltitulo 
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
         Left            =   4440
         TabIndex        =   15
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Rutina"
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
         Left            =   4440
         TabIndex        =   14
         Top             =   840
         Width           =   615
      End
   End
End
Attribute VB_Name = "GeneraHistograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RMinimo As New ADODB.Recordset
Dim RMaximo As New ADODB.Recordset
Dim Rcontador As New ADODB.Recordset
Dim RBuscaCantidadDatos As New ADODB.Recordset
Dim RMedia As New ADODB.Recordset
Dim RMedia2 As New ADODB.Recordset
Dim RDesviacion As New ADODB.Recordset
Dim RDesviacion2 As New ADODB.Recordset
Dim RDatos As New ADODB.Recordset
Dim RVariables As New ADODB.Recordset
Dim RFichaTecnica As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaRutinas As New ADODB.Recordset
Dim RDatosHistograma As New ADODB.Recordset
Dim RBusqueda As New ADODB.Recordset

Dim VCabezales As Long
Dim VMinimo As Currency
Dim VMaximo As Currency
Dim VRango As Currency
Dim VGrupos As Integer
Dim VIntervalo As Currency
Dim VMedia As Currency
Dim VMedia2 As Currency
Dim VDesviacion As Currency
Dim VDesviacion2 As Currency

Dim VCV As Currency
Dim Vb As Currency

Dim VLSC As Currency
Dim VLSI As Currency
Dim VLIC As Currency
Dim VLII As Currency

Dim VContador As Integer
Dim VContGrafica As Integer

Dim Xi As Currency
Dim Fi As Currency
Dim P1 As Currency
Dim P2 As Currency

Dim Cont As Integer
Dim ContFilas As Integer

Dim RBuscaRutina As New ADODB.Recordset
Dim RBuscaFicha As New ADODB.Recordset

Dim RHistograma As New ADODB.Recordset

Dim BRutina As Boolean
Dim BFicha As Boolean
Dim BCatalogo As Boolean
Dim BLinea As Boolean

Dim VPrimerValor As Currency
Dim VSegundoValor As Currency
Dim VTercerValor As Currency
Dim VCuartoValor As Currency
Dim VQuintoValor As Currency
Dim VSextoValor As Currency
Dim VSeptimoValor As Currency

Dim VPrimerDato As Currency
Dim VSegundoDato As Currency
Dim VTercerDato As Currency
Dim VCuartoDato As Currency
Dim VQuintoDato As Currency
Dim VSextoDato As Currency
Dim VSeptimoDato As Currency

Dim VPrimerDato2 As Currency
Dim VSegundoDato2 As Currency
Dim VTercerDato2 As Currency
Dim VCuartoDato2 As Currency
Dim VQuintoDato2 As Currency
Dim VSextoDato2 As Currency
Dim VSeptimoDato2 As Currency


Dim ContadorColumnas As Integer

Dim RBuscaFoto As New ADODB.Recordset

Dim LSP As Currency
Dim LIP As Currency
Dim CP As Currency

Dim VTexto As String



Private Sub CboVerGra_Click()

    If CboVerGra.Text = "2dArea" Then
        Grafica.chartType = VtChChartType2dArea
    ElseIf CboVerGra.Text = "2dBar" Then
        Grafica.chartType = VtChChartType2dBar
    ElseIf CboVerGra.Text = "2dCombination" Then
        Grafica.chartType = VtChChartType2dCombination
    ElseIf CboVerGra.Text = "2dLine" Then
        Grafica.chartType = VtChChartType2dLine
    ElseIf CboVerGra.Text = "type2dpie" Then
        Grafica.chartType = VtChChartType2dPie
    ElseIf CboVerGra.Text = "2dStep" Then
        Grafica.chartType = VtChChartType2dStep
    ElseIf CboVerGra.Text = "Type2dXY" Then
        Grafica.chartType = VtChChartType2dXY
    ElseIf CboVerGra.Text = "3dArea" Then
        Grafica.chartType = VtChChartType3dArea
    ElseIf CboVerGra.Text = "3dBar" Then
        Grafica.chartType = VtChChartType3dBar
    ElseIf CboVerGra.Text = "3dCombination" Then
        Grafica.chartType = VtChChartType3dCombination
    ElseIf CboVerGra.Text = "3dLine" Then
        Grafica.chartType = VtChChartType3dLine
    ElseIf CboVerGra.Text = "3dStep" Then
        Grafica.chartType = VtChChartType3dStep
    End If
End Sub

Private Sub ChkLin_Click()
    If ChkLin.Value = 1 Then
        LblLinea.Visible = True
        LblLineas.Visible = True
        TxtLin.Visible = True
        TxtLin.SetFocus
    Else
        LblLinea.Visible = False
        LblLineas.Visible = False
        TxtLin.Visible = False
    End If
End Sub

Private Sub CmdCopiar_Click()
    Grafica.EditCopy
    MsgBox "Grafica Copiada Al Clipboard", vbOKOnly
End Sub

Private Sub CmdGenerar_Click()
On Error Resume Next

MousePointer = 11

        'Set RHistograma = New ADODB.Recordset
        '    Call Abrir_Recordset(RHistograma, "Select * from Histograma")

VCabezales = 0
FGridHistograma.Clear
Grafica.RowCount = 0
ContadorColumnas = 0
Cont = 1
ContFilas = 2

    'VALIDA EL CABEZAL
    If optcab.Item(1).Value = True Then
            If IsNumeric(txtcab.Text) Then
            Else
                MsgBox "Cabezal De Inicio Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
            End If
            
            If IsNumeric(txtcab2.Text) Then
            Else
                MsgBox "Cabezal Final Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
                Exit Sub
            End If
            
    End If
            
    Set RDatos = New ADODB.Recordset
    'SI QUIEREN INCLUIR EN REPORTE LAS LINEAS
    If ChkLin.Value = 1 Then
                'SELECCIONA DATOS POR FICHA TENCNICA O POR CATALOGO
                If OptFicha.Value = True Then
                    'CABEZAL
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDatos, "Select Linea, Fec_rut, Hor_rut, Cabezal, Valor from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "' Order By Valor")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDatos, "Select Linea, Fec_rut, Hor_rut, Cabezal, Valor from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' Order By Valor")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDatos, "Select Linea, Fec_rut, Hor_rut, Cabezal, Valor from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "' and Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text & " Order By Valor")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDatos, "Select Linea, Fec_rut, Hor_rut, Cabezal, Valor from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text & "  Order By Valor")
                            End If
                    
                            
                    End If
                Else
                    'CABEZAL
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDatos, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And Cr.Linea = '" & TxtLin.Text & "' Order By CR.Valor ")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDatos, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(Cr.Linea) = '" & UCase(TxtLin.Text) & "' Order By CR.Valor ")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDatos, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And Cr.Linea = '" & TxtLin.Text & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Valor ")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDatos, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(Cr.Linea) = '" & UCase(TxtLin.Text) & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & "  Order By CR.Valor ")
                            End If
                    End If
                End If
    'TODOS LOS DATOS
    Else
                'SELECCIONA DATOS POR FICHA TENCNICA O POR CATALOGO
                If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDatos, "Select Linea, Fec_rut, Hor_rut, Cabezal, Valor from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 Order By Valor")
                            Else
                                Call Abrir_Recordset(RDatos, "Select Linea, Fec_rut, Hor_rut, Cabezal, Valor from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 Order By Valor")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDatos, "Select Linea, Fec_rut, Hor_rut, Cabezal, Valor from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 and Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text & " Order By Valor")
                            Else
                                Call Abrir_Recordset(RDatos, "Select Linea, Fec_rut, Hor_rut, Cabezal, Valor from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 and Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text & " Order By Valor")
                            End If
                    End If
                Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDatos, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 Order By CR.Valor")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDatos, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 Order By CR.Valor")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDatos, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal >= " & txtcab2.Text & " Order By CR.Valor")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDatos, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Valor")
                            End If
                    End If
                End If
    End If
      
      
      If RDatos.RecordCount > 0 Then
      Else
            If ChkLin.Value = 1 Then
                MsgBox "No Hay Datos En Estas Fechas Con Esta Rutina Y Linea", vbOKOnly + vbInformation, "Informacion"
            Else
                MsgBox "No Hay Datos En Estas Fechas", vbOKOnly + vbInformation, "Informacion"
            End If
            MousePointer = 0
            Exit Sub
      End If
      
                Set RDatosHistograma = New ADODB.Recordset
                     'SI PIDE POR LINEA
                     If ChkLin.Value = 1 Then
                               'MUESTRA TODAS LAS RUTINAS SELECCIONADAS POR FECHA Y FICHA TECNICA Y RUTINA
                               If OptFicha.Value = True Then
                                     If optcab.Item(0).Value = True Then
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and Cr.Esp_Tec = Ft.Esp_Tec And CR.Esp_Tec = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Cr.Esp_Tec) = UPPER(Ft.Esp_Tec) And UPPER(CR.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                End If
                                     Else
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and Cr.Esp_Tec = Ft.Esp_Tec And CR.Esp_Tec = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Cr.Esp_Tec) = UPPER(Ft.Esp_Tec) And UPPER(CR.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                End If
                                     End If
                               Else
                                    If optcab.Item(0).Value = True Then
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                End If
                                    Else
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                End If
                                    End If
                               End If
                     'TODOS
                     Else
                               'MUESTRA TODAS LAS RUTINAS SELECCIONADAS POR FECHA Y FICHA TECNICA Y RUTINA
                               If OptFicha.Value = True Then
                                      If optcab.Item(0).Value = True Then
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and Cr.Esp_Tec = Ft.Esp_Tec And CR.Esp_Tec = '" & TxtFicTec.Text & "' And CR.Valor > 0 Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Cr.Esp_Tec) = UPPER(Ft.Esp_Tec) And UPPER(CR.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                End If
                                      Else
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and Cr.Esp_Tec = Ft.Esp_Tec And CR.Esp_Tec = '" & TxtFicTec.Text & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Cr.Esp_Tec) = UPPER(Ft.Esp_Tec) And UPPER(CR.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                End If
                                      End If
                               Else
                                      If optcab.Item(0).Value = True Then
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                End If
                                      Else
                                                If GOrigenDeDatos = "AmaproAccess" Then
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                Else 'ORACLE
                                                    Call Abrir_Recordset(RDatosHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text & " Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                                End If
                                      End If
                               End If
                    End If
     
     'BUSCA CUANTOS CABEZALES TIENE LA RUTINA
     Set RBuscaRutinas = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaRutinas, "Select Cabezal From Rutinas Where Rutina = '" & TxtRut.Text & "'")
            Else
                Call Abrir_Recordset(RBuscaRutinas, "Select Cabezal From Rutinas Where UPPER(Rutina) = '" & UCase(TxtRut.Text) & "'")
            End If
            If RBuscaRutinas.RecordCount > 0 Then
                VCabezales = RBuscaRutinas!cabezal
            Else
                MsgBox "Codigo De Rutina No Existe", vbOKOnly + vbInformation, "Informacion"
                MousePointer = 0
                Exit Sub
            End If
     
      Set DbGridHistograma.DataSource = RDatosHistograma
      
      DbGridHistograma.Columns(0).Width = "300"
      DbGridHistograma.Columns(1).Width = "1000"
      DbGridHistograma.Columns(1).NumberFormat = "dd/mm/yyyy"
      DbGridHistograma.Columns(2).Width = "500"
      DbGridHistograma.Columns(3).Width = "300"
      DbGridHistograma.Columns(4).Width = "700"
      DbGridHistograma.Columns(5).Width = "1000"
      DbGridHistograma.Columns(6).Width = "5000"
        
      If OptCatalogo.Value = True Then
        DbGridHistograma.Columns(5).Width = "1000"
        DbGridHistograma.Columns(6).Width = "5000"
      End If
      
      
      'SACA EL DATO MINIMO DE LAS RUTINAS SELECCIONADAS POR FECHA Y FICHA TECNICA Y RUTINA Y Linea
      'SI ESCOJE POR LINEA
      Set RMinimo = New ADODB.Recordset
      If ChkLin.Value = 1 Then
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMinimo, "Select Min(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "'")
                            Else
                                Call Abrir_Recordset(RMinimo, "Select Min(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMinimo, "Select Min(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "' And Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RMinimo, "Select Min(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMinimo, "Select Min(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RMinimo, "Select Min(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMinimo, "Select Min(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' And CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else 'ORACLE
                                Call Abrir_Recordset(RMinimo, "Select Min(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' And CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      'TODAS LAS LINEAS
      Else
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMinimo, "Select Min(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0")
                            Else
                                Call Abrir_Recordset(RMinimo, "Select Min(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMinimo, "Select Min(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RMinimo, "Select Min(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMinimo, "Select Min(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0")
                            Else
                                Call Abrir_Recordset(RMinimo, "Select Min(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMinimo, "Select Min(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RMinimo, "Select Min(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And CR.Cabezal = " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      End If
            
            If RMinimo.RecordCount > 0 Then
               If Not IsNull(RMinimo(0)) Then
                  VMinimo = Format(RMinimo(0), "#,###,##0.00")
               Else
                  VMinimo = Format(0, "#,###,##0.00")
               End If
            End If
        
        
      Set RMaximo = New ADODB.Recordset
      'SACA EL DATO MAXIMO DE LAS RUTINAS SELECCIONADAS POR FECHA Y FICHA TECNICA Y RUTINA Y LINEA
      'POR LINEA
      If ChkLin.Value = 1 Then
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMaximo, "Select Max(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "'")
                            Else
                                Call Abrir_Recordset(RMaximo, "Select Max(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMaximo, "Select Max(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "' And cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RMaximo, "Select Max(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' And cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMaximo, "Select Max(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RMaximo, "Select Max(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMaximo, "Select Max(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' And CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else 'ORACLE
                                Call Abrir_Recordset(RMaximo, "Select Max(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' And CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      'TODAS LAS LINEAS
      Else
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMaximo, "Select Max(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0")
                            Else
                                Call Abrir_Recordset(RMaximo, "Select Max(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMaximo, "Select Max(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RMaximo, "Select Max(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RMaximo, "Select Max(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0")
                                Else
                                    Call Abrir_Recordset(RMaximo, "Select Max(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0")
                                End If
                    Else
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RMaximo, "Select Max(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And Cr.Cabezal >= " & txtcab.Text & " And cR.Cabezal <= " & txtcab2.Text)
                                Else
                                    Call Abrir_Recordset(RMaximo, "Select Max(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 and cr.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                                End If
                    End If
            End If
      End If
        
            If RMaximo.RecordCount > 0 Then
               If Not IsNull(RMaximo(0)) Then
                  VMaximo = Format(RMaximo(0), "#,###,##0.00")
               Else
                  VMaximo = Format(0, "#,###,##0.00")
               End If
            End If
        
        'SACA EL TOTAL DE REGISTROS DE LAS RUTINAS SELECCIONADAS POR FECHA Y FICHA TECNICA Y RUTINA
        'POR UNA LINEA
        
      Set Rcontador = New ADODB.Recordset
      If ChkLin.Value = 1 Then
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(Rcontador, "Select Count(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "'")
                            Else
                                Call Abrir_Recordset(Rcontador, "Select Count(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(Rcontador, "Select Count(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(Rcontador, "Select Count(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(Rcontador, "Select Count(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(Rcontador, "Select Count(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(Rcontador, "Select Count(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else 'ORACLE
                                Call Abrir_Recordset(Rcontador, "Select Count(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      'TODAS LAS LINEAS
      Else
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(Rcontador, "Select Count(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0")
                            Else
                                Call Abrir_Recordset(Rcontador, "Select Count(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(Rcontador, "Select Count(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(Rcontador, "Select Count(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(Rcontador, "Select Count(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0")
                            Else
                                Call Abrir_Recordset(Rcontador, "Select Count(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(Rcontador, "Select Count(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " CR.And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(Rcontador, "Select Count(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      End If
        
            If Rcontador.RecordCount > 0 Then
               If Not IsNull(Rcontador(0)) Then
                  VContador = Rcontador(0)
               Else
                  VContador = 0
               End If
            End If
            
            
        'SACA LA DESVIACION DE REGISTROS DE LAS RUTINAS SELECCIONADAS POR FECHA Y FICHA TECNICA Y RUTINA
        'POR UNA LINEA
        Set RDesviacion2 = New ADODB.Recordset
        If ChkLin.Value = 1 Then
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDesviacion2, "Select Stdevp(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "'")
                            Else
                                Call Abrir_Recordset(RDesviacion2, "Select Stddev(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDesviacion2, "Select Stdevp(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RDesviacion2, "Select Stddev(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDesviacion2, "Select Stdevp(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RDesviacion2, "Select Stddev(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDesviacion2, "Select Stdevp(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' And CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else 'ORACLE
                                Call Abrir_Recordset(RDesviacion2, "Select Stddev(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' And CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      'TODAS LAS LINEAS
      Else
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDesviacion2, "Select Stdevp(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0")
                            Else
                                Call Abrir_Recordset(RDesviacion2, "Select Stddev(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDesviacion2, "Select Stdevp(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RDesviacion2, "Select Stddev(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDesviacion2, "Select Stdevp(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0")
                            Else
                                Call Abrir_Recordset(RDesviacion2, "Select Stddev(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RDesviacion2, "Select Stdevp(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 and Cr.cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RDesviacion2, "Select Stddev(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      End If
      
      
            If RDesviacion2.RecordCount > 0 Then
               If Not IsNull(RDesviacion2(0)) Then
                  VDesviacion2 = RDesviacion2(0)
               Else
                  VDesviacion2 = 0
               End If
            End If
            
                
        'SACA LA MEDIA DE REGISTROS DE LAS RUTINAS SELECCIONADAS POR FECHA Y FICHA TECNICA Y RUTINA
        'POR UNA LINEA
        Set RMedia2 = New ADODB.Recordset
        If ChkLin.Value = 1 Then
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMedia2, "Select Avg(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "'")
                            Else
                                Call Abrir_Recordset(RMedia2, "Select Avg(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMedia2, "Select Avg(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 And Linea = '" & TxtLin.Text & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RMedia2, "Select Avg(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMedia2, "Select Avg(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RMedia2, "Select Avg(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMedia2, "Select Avg(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' and CR.cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else 'ORACLE
                                Call Abrir_Recordset(RMedia2, "Select Avg(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      'TODAS LAS LINEAS
      Else
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMedia2, "Select Avg(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0")
                            Else
                                Call Abrir_Recordset(RMedia2, "Select Avg(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMedia2, "Select Avg(Valor) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RMedia2, "Select Avg(Valor) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMedia2, "Select Avg(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0")
                            Else
                                Call Abrir_Recordset(RMedia2, "Select Avg(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RMedia2, "Select Avg(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 and CR.cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RMedia2, "Select Avg(CR.Valor) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 and CR.cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
      End If
        
      
            If RMedia2.RecordCount > 0 Then
               If Not IsNull(RMedia2(0)) Then
                  VMedia2 = RMedia2(0)
               Else
                  VMedia2 = 0
               End If
            End If
            
            
       'buscamos el codigo de variable que le pertenece a esta ficha tecnica
       Set RFichaTecnica = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RFichaTecnica, "Select Variables From FichaTecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
            Else
                Call Abrir_Recordset(RFichaTecnica, "Select Variables From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
       
                'SACAMOS LOS STANDARES DEL CLIENTE DE ACUERDO A LA FICHA TECNICA
                Set RVariables = New ADODB.Recordset
                If OptFicha.Value = True Then
                         If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RVariables, "Select MaximoClienteMilimetros, MinimoClienteMilimetros, MaximoInternoMilimetros, MinimoInternoMilimetros From VariablesMedia Where Codigo = '" & RFichaTecnica(0) & "' and Rutina = '" & TxtRut.Text & "'")
                         Else
                            Call Abrir_Recordset(RVariables, "Select MaximoClienteMilimetros, MinimoClienteMilimetros, MaximoInternoMilimetros, MinimoInternoMilimetros From VariablesMedia Where UPPER(Codigo) = '" & UCase(RFichaTecnica(0)) & "' and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "'")
                         End If
                Else
                         If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RVariables, "Select MaximoClienteMilimetros, MinimoClienteMilimetros, MaximoInternoMilimetros, MinimoInternoMilimetros From VariablesMedia Where Codigo = '" & TxtFicTec.Text & "' and Rutina = '" & TxtRut.Text & "'")
                         Else
                            Call Abrir_Recordset(RVariables, "Select MaximoClienteMilimetros, MinimoClienteMilimetros, MaximoInternoMilimetros, MinimoInternoMilimetros From VariablesMedia Where UPPER(Codigo) = '" & UCase(TxtFicTec.Text) & "' and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "'")
                         End If
                End If
                
        
        'SI PIDE EL REPORTE EN PULGADAS SE HACE LA CONVERSION
        'If OptPulgadas.Value = True Then
        '    If RVariables.RecordCount > 0 Then
        '        RVariables(0) = RVariables(0) / 25.4
        '        RVariables(1) = RVariables(1) / 25.4
        '        RVariables(2) = RVariables(2) / 25.4
        '        RVariables(3) = RVariables(3) / 25.4
        '    End If
        'End If
        
       
       If RVariables.RecordCount > 0 Then
           VLSC = RVariables(0)
           VLIC = RVariables(1)
           VLSI = RVariables(2)
           VLII = RVariables(3)
       Else
           VLSC = 0
           VLSI = 0
           VLIC = 0
           VLII = 0
       End If

       'SACAMOS LA MEDIA
        VMedia = (Val(VLSC) + Val(VLIC)) / 2
        'VMedia2 = (Val(VLSI) + Val(VLII)) / 2
       'SACAMOS LA DESVIACION STANDARD
        VDesviacion = (Val(VLSC) - Val(VLIC)) / VCabezales
        'VDesviacion2 = (Val(VLSI) - Val(VLII)) / 6

'-----------------------------------------------------------------------
'SACA FORMULAS
            
     'RANGO
      VRango = Format(VMaximo - VMinimo, "#,###,##0.00")
            
      'FORMA GRUPOS
      VGrupos = Round(1 + 3.3 * Log(VContador) / Log(10), 0)
      'INTERVALO
      VIntervalo = Format(VRango / VGrupos, "#,###,##0.00")
      
      TxtRan.Text = VRango
      TxtGru.Text = VGrupos
      TxtInt.Text = VIntervalo
      
      TxtDatMen.Text = VMinimo
      TxtDatMay.Text = VMaximo
      
      TxtLSC.Text = VLSC
      TxtLIC.Text = VLIC
      
      'MEDIA Y DESVIACION CLIENTE
      TxtMed.Text = VMedia
      TxtDes.Text = VDesviacion
      
      'MEDIA Y DESVIACION INTERNA
      Txtmediainterno.Text = Format(VMedia2, "#,###,##0.00")
      TxtDesviacionInterno.Text = Format(VDesviacion2, "#,###,##0.00")
      
      
      If VDesviacion = 0 Then
         TxtCv.Text = Format(0, "#,###,##0.00")
      Else
        TxtCv.Text = Format(((VDesviacion / VMedia) * 100), "#,###,##0.00")
      End If
      
      If VDesviacion = 0 Then
         Vb = 0
      Else
         Vb = Format(1 / Val(VDesviacion) ^ 2, "#,###,##0.000")
      End If
      
      Txtb.Text = Vb
      
       'CP
       LSP = 3 * TxtDesviacionInterno.Text + Txtmediainterno.Text
       LIP = -3 * TxtDesviacionInterno.Text + Txtmediainterno.Text
       
       CP = (TxtLSC.Text - TxtLIC.Text) / (LSP - LIP)
       TxtCP.Text = CP

        '-----------------------------------------------------------------------------
        'titulos del grid
        FGridHistograma.Rows = ContFilas
        FGridHistograma.Row = 0
        FGridHistograma.ColWidth(0) = "0"
        FGridHistograma.Col = 1
        FGridHistograma.Text = "P1"
        FGridHistograma.Col = 2
        FGridHistograma.Text = "P2"
        FGridHistograma.Col = 3
        FGridHistograma.Text = "Xi"
        FGridHistograma.Col = 4
        FGridHistograma.Text = "Fi"


        'CUANTAS COLUMNAS VA A TENER LA GRAFICA
        If VMinimo = VMaximo Then
            Grafica.ColumnCount = 2
            Grafica.RowCount = 1
            
            Grafica.Row = 1
            Grafica.Column = 1
            Grafica.Data = 0
           
            Grafica.Column = 2
            Grafica.Data = 0
            
        Else
            Grafica.ColumnCount = 2
            Grafica.RowCount = VGrupos + 1
        End If

        'LLENAMOS LA PRIMERA LINEA
        FGridHistograma.Rows = ContFilas
        FGridHistograma.Row = Cont
        FGridHistograma.Col = 1
        FGridHistograma.Text = VMinimo
        P1 = VMinimo
        
        FGridHistograma.Col = 2
        FGridHistograma.Text = (Val(VMinimo) + Val(VIntervalo)) - Val(0.01)
        P2 = (Val(VMinimo) + Val(VIntervalo)) - Val(0.01)

        'CALCULA LA MITAD DENTRO DE LOS RANGOS P1 Y P2
        Xi = Format(((Val(P1) + Val(P2)) / 2), "#,###,##0.00")
        FGridHistograma.Col = 3
        FGridHistograma.Text = Xi
    
        'CALCULA LA CANTIDAD DE DATOS DENTRO DE LOS RANGOS P1 Y P2
        'POR UNA LINEA
        Set RBuscaCantidadDatos = New ADODB.Recordset
        If ChkLin.Value = 1 Then
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And Linea = '" & TxtLin.Text & "'")
                            Else
                                Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And Linea = '" & TxtLin.Text & "' And Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' and CR.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "'")
                            Else
                                Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' and CR.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' and CR.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else
                                Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' and CR.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' And CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
        'POR TODAS LAS LINEAS
        Else
            If OptFicha.Value = True Then
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0")
                            Else 'ORACLE
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            Else 'ORACLE
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            Else
                    If optcab.Item(0).Value = True Then
                            If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' and CR.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0")
                            Else
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' and CR.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0")
                            End If
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' and CR.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            Else
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' and CR.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                            End If
                    End If
            End If
        End If

    
    If RBuscaCantidadDatos.RecordCount > 0 Then
        If Not IsNull(RBuscaCantidadDatos(0)) Then
            Fi = RBuscaCantidadDatos(0)
        Else
            Fi = 0
        End If
    End If
    
'LLENA DE DATOS LA GRAFICA PARA LA PRIMERA COLUMNA _________________________________________________________________
    If VMinimo = VMaximo Then
    Else
        Grafica.Row = 1
        Grafica.Column = 1
        Grafica.Data = Fi
        
        Grafica.Row = 1
        Grafica.RowLabel = Xi
        Grafica.Column = 1
        
        'guarda valores
        VPrimerValor = Xi
        VPrimerDato = Fi
                                          
        Grafica.Column = 2
        'formula para sacar dato de tendencia normal
        Grafica.Data = ((VContador / 2) * Exp(-Vb * (VMedia - Xi) ^ 2))
        
        'Guarda valores
        VPrimerDato2 = Grafica.Data
    End If
    
    Cont = Cont + 1
    ContFilas = ContFilas + 1

          
    FGridHistograma.Col = 4
    FGridHistograma.Text = Fi
    

      'SE LLENA EL GRID CON LOS DATOS CREA UN CONTADOR PARA QUE EMPIESE DESDE LA SEGUNDA LINEA HASTA EL FINAL

      If VMinimo = VMaximo Then
      
      Else
            Do Until VMinimo > VMaximo
                        
                ContadorColumnas = ContadorColumnas + 1
                
                FGridHistograma.Rows = ContFilas
                VMinimo = Val(VMinimo) + Val(VIntervalo)
                FGridHistograma.Row = Cont
                
                FGridHistograma.Col = 1
                FGridHistograma.Text = VMinimo
                P1 = VMinimo
                
                FGridHistograma.Col = 2
                FGridHistograma.Text = (Val(VMinimo) + Val(VIntervalo)) - Val(0.01)
                P2 = (Val(VMinimo) + Val(VIntervalo)) - Val(0.01)
                
                Xi = Format(((Val(P1) + Val(P2)) / 2), "#,###,##0.00")
                
                                        'SI SE FORMAN 10 REGISTROS EN EL GRID GUARDAMOS LOS VALORES EN LAS VARIABLES
                                        'PARA DE ULTIMO COLOCARLOS EN LA ROWS 2 Y TRES
                                            If Cont = 2 Then
                                                VSegundoValor = Xi
                                            End If
                                            If Cont = 3 Then
                                                VTercerValor = Xi
                                            End If
                                            If Cont = 4 Then
                                                VCuartoValor = Xi
                                            End If
                                            If Cont = 5 Then
                                                VQuintoValor = Xi
                                            End If
                                            If Cont = 6 Then
                                                VSextoValor = Xi
                                            End If
                                            If Cont = 7 Then
                                                VSeptimoValor = Xi
                                            End If
                
                FGridHistograma.Col = 3
                FGridHistograma.Text = Xi
                
                'CALCULA LA CANTIDAD DE DATOS DENTRO DE LOS RANGOS P1 Y P2
                'POR UNA LINEA
                Set RBuscaCantidadDatos = New ADODB.Recordset
                If ChkLin.Value = 1 Then
                    If OptFicha.Value = True Then
                        If optcab.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And Linea = '" & TxtLin.Text & "'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date(#" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And Linea = '" & TxtLin.Text & "'")
                                End If
                        Else
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And Linea = '" & TxtLin.Text & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date(#" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And Linea = '" & TxtLin.Text & "' and cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                                End If
                        End If
                    Else
                        If optcab.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' and Cr.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "'")
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' and Cr.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "'")
                                End If
                        Else
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' and Cr.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 And CR.Linea = '" & TxtLin.Text & "' and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                                Else 'ORACLE
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' and Cr.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 And UPPER(CR.Linea) = '" & UCase(TxtLin.Text) & "' AND CR.cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                                End If
                        End If
                    End If
                'POR TODAS LAS LINEAS
                Else
                    If OptFicha.Value = True Then
                        If optcab.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0")
                                Else
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0")
                                End If
                        Else
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and Rutina = '" & TxtRut.Text & "' and Esp_Tec = '" & TxtFicTec.Text & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 And Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                                Else
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select count(*) from CapturaRutinas where Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' and Valor >= " & P1 & " and Valor <= " & P2 & " And Valor > 0 and Cabezal >= " & txtcab.Text & " And Cabezal <= " & txtcab2.Text)
                                End If
                        End If
                    Else
                        If optcab.Item(0).Value = True Then
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' and Cr.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0")
                                Else
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' and Cr.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0")
                                End If
                        Else
                                If GOrigenDeDatos = "AmaproAccess" Then
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' and Cr.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                                Else
                                    Call Abrir_Recordset(RBuscaCantidadDatos, "Select COUNT(*) from CapturaRutinas CR, FichaTecnica FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(CR.Esp_Tec) = UPPER(FT.Esp_tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' and Cr.Valor >= " & P1 & " and Cr.Valor <= " & P2 & " And CR.Valor > 0 and CR.Cabezal >= " & txtcab.Text & " And CR.Cabezal <= " & txtcab2.Text)
                                End If
                        End If
                    End If
                End If
                
                If RBuscaCantidadDatos.RecordCount > 0 Then
                    If Not IsNull(RBuscaCantidadDatos(0)) Then
                        Fi = RBuscaCantidadDatos(0)
                    Else
                        Fi = 0
                    End If
                End If
                
                FGridHistograma.Col = 4
                FGridHistograma.Text = Fi
                
                
                                        'SI SE FORMAN 2 O MAS REGISTROS APARTE DE LOS GRUPOS ACETADOS
                                        'EN EL GRID GUARDAMOS LOS VALORES EN LAS VARIABLES
                                        'PARA DE ULTIMO COLOCARLOS EN LA ROWS CORRESPONDIENTES
                                            
                                            If Cont = 2 Then
                                                VSegundoDato = Fi
                                            End If
                                            If Cont = 3 Then
                                                VTercerDato = Fi
                                            End If
                                            If Cont = 4 Then
                                                VCuartoDato = Fi
                                            End If
                                            If Cont = 5 Then
                                                VQuintoDato = Fi
                                            End If
                                            If Cont = 6 Then
                                                VSextoDato = Fi
                                            End If
                                            If Cont = 7 Then
                                                VSeptimoDato = Fi
                                            End If
                
            '--------------------------------------- GRAFICA ---------------------------------------
                'LLENA DE DATOS LA GRAFICA
                If P1 > VMaximo Then
                'If Cont > VGrupos Then
                                Grafica.Row = Cont
                                Grafica.Column = 1
                                Grafica.Data = 0
                                Grafica.RowLabel = Xi
                                Grafica.Row = Cont
                                Grafica.Column = 2
                                Grafica.Data = 0
                                Grafica.RowLabel = Xi
                
                Else
                        If Err = 1117 Then
                                
                                Grafica.Row = Cont
                                Grafica.Column = 1
                                Grafica.Data = Fi
                               
                                Grafica.Row = Cont
                                Grafica.RowLabel = Xi
                                Grafica.Row = Cont
                                Grafica.Column = 2
                                'formula para sacar dato de tendencia normal
                                Grafica.Data = ((VContador / 2) * Exp(-Vb * (VMedia - Xi) ^ 2))
                                 
                        Else
                                
                                
                                Grafica.Row = Cont
                                Grafica.Column = 1
                                Grafica.Data = Fi
                                
                                Grafica.Row = Cont
                                Grafica.RowLabel = Xi
                                Grafica.Row = Cont
                                
                                Grafica.Column = 2
                                'formula para sacar dato de tendencia normal
                                Grafica.Data = ((VContador / 2) * Exp(-Vb * (VMedia - Xi) ^ 2))
                                
                        End If
                                
               End If
                        Cont = Cont + 1
                        ContFilas = ContFilas + 1
                        
            Loop
                        If (ContadorColumnas + 1) = VGrupos Then
                            Grafica.RowCount = VGrupos
                            Grafica.Row = VGrupos
                            Grafica.RowLabel = Xi
                        ElseIf ContadorColumnas = VGrupos Then
                            Grafica.RowCount = VGrupos
                        ElseIf ContadorColumnas > VGrupos Then
                            Grafica.RowCount = VGrupos + 1
                            
                            Grafica.Column = 1
                            Grafica.Row = 1
                            Grafica.RowLabel = VPrimerValor
                            Grafica.Data = VPrimerDato
                            Grafica.Row = 2
                            Grafica.RowLabel = VSegundoValor
                            Grafica.Data = VSegundoDato
                            Grafica.Row = 3
                            Grafica.RowLabel = VTercerValor
                            Grafica.Data = VTercerDato
                            Grafica.Row = 4
                            Grafica.RowLabel = VCuartoValor
                            Grafica.Data = VCuartoDato
                            Grafica.Row = 5
                            Grafica.RowLabel = VQuintoValor
                            Grafica.Data = VQuintoDato
                            Grafica.Row = 6
                            Grafica.RowLabel = VSextoValor
                            Grafica.Data = VSextoDato
                            Grafica.Row = 7
                            Grafica.RowLabel = VSeptimoValor
                            Grafica.Data = VSeptimoDato
                            If Err <> 0 Then
                            End If
                            
                        Else
                            
                        End If
                        Grafica.Row = 1
                        Grafica.RowLabel = VPrimerValor
            
       End If
  
'-------------------------------------------------------------------------------------------
'GENERA REPORTES
If ChkReporte.Value = 1 Then
        Call ImprimeHistograma
End If

  
MousePointer = 0
End Sub



Private Sub CmdGrabar_Click()

       
   CDDialogo.CancelError = True
   On Error GoTo ErrHandler
       
    CDDialogo.InitDir = App.Path
    CDDialogo.ShowSave
    
    
            Grafica.EditCopy
             
            SavePicture Clipboard.GetData, CDDialogo.FileName
            MsgBox "La gráfica ha sido guardada ", vbInformation, "Guardar gráfica"
    
ErrHandler:
  'User pressed the Cancel button
  Exit Sub


End Sub

Private Sub CmdImprimirGrafica_Click()
On Error Resume Next
    MousePointer = 11
        Printer.Font = "Courier New"
        
        If Err <> 0 Then
            MsgBox Err.Description
            MousePointer = 0
            Exit Sub
        End If
        
        Grafica.ShowLegend = False
        Grafica.EditCopy
        Printer.PaintPicture Clipboard.GetData, 100, 100
        
        Cont = 1
        Do Until Cont = 50
                Printer.Print
            Cont = Cont + 1
        Loop
        
        Printer.FontBold = True
        Printer.ForeColor = vbGreen
        Printer.FontSize = "18"
        Printer.Print Space(10) & "HISTOGRAMA DE EVALUACION DE VARIABLE"
        Printer.FontSize = "14"
        Printer.Print Space(6) & "_____________________________________________________________________________________________________________"
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        Printer.Print
        Printer.Print Space(6) & "PERIODO DESDE ";
        Printer.FontBold = True
        Printer.ForeColor = vbBlue
        Printer.Print Format(PFecRutIni.Value, "dd/mm/yyyy");
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        Printer.Print " HASTA ";
        Printer.FontBold = True
        Printer.ForeColor = vbBlue
        Printer.Print Format(PFecRutFin.Value, "dd/mm/yyyy")
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        Printer.Print Space(6) & "RUTINA ";
        Printer.FontBold = True
        Printer.ForeColor = vbBlue
        Printer.Print TxtRut.Text;
        Printer.Print Space(5) & LblRutina.Caption
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        If OptFicha.Value = True Then
            Printer.Print Space(6) & "FICHA TECNICA ";
            Printer.FontBold = True
            Printer.ForeColor = vbBlue
                Printer.Print TxtFicTec.Text & Space(5) & LblFichaTecnica.Caption
            Printer.ForeColor = &H80000008
            Printer.FontBold = False
        Else
            Printer.Print Space(6) & "CATALOGO ";
            Printer.FontBold = True
            Printer.ForeColor = vbBlue
                Printer.Print TxtFicTec.Text & Space(5) & LblFichaTecnica.Caption
            Printer.ForeColor = &H80000008
            Printer.FontBold = False
        End If
        
        'SI PIDE POR LINEA
        If ChkLin.Value = 1 Then
            Printer.Print Space(6) & "LINEA ";
            Printer.FontBold = True
            Printer.ForeColor = vbBlue
                Printer.Print TxtLin.Text & Space(5) & LblLineas.Caption
            Printer.ForeColor = &H80000008
            Printer.FontBold = False
        End If
        
        Printer.FontSize = "8"
        Printer.Print
        Printer.Print
        Printer.Print Space(6) & "RANGO                   ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtRan.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "GRUPOS                  ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtGru.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "INTERVALO               ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtInt.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "DATO MENOR              ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtDatMen.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "DATO MAYOR              ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtDatMay.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "LIMITE SUPERIOR CLIENTE ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtLSC.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "LIMITE INFERIOR CLIENTE ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtLIC.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "MEDIA                   ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtMed.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "MEDIA DE DATOS          ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print Txtmediainterno.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "DESVIACION              ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtDes.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "DESVIACION DE DATOS     ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtDesviacionInterno.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "CV                      ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtCv.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "B                       ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print Txtb.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.Print Space(6) & "CP                      ";
        Printer.FontBold = True
        Printer.ForeColor = vbRed
            Printer.Print TxtCP.Text
        Printer.ForeColor = &H80000008
        Printer.FontBold = False
        
        Printer.EndDoc
        
    MousePointer = 0


        Grafica.ShowLegend = True
End Sub

Private Sub CmdSale_Click()
    FrameBusqueda.Visible = False
End Sub


Private Sub DBGridBusqueda_DblClick()
    If BRutina = True Then
        TxtRut.Text = DBGridBusqueda.Columns(0)
        TxtRut.SetFocus
    ElseIf BLinea = True Then
        TxtLin.Text = DBGridBusqueda.Columns(0)
        TxtLin.SetFocus
    Else
        TxtFicTec.Text = DBGridBusqueda.Columns(0)
        TxtFicTec.SetFocus
    End If
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 43 Then
        If BRutina = True Then
            TxtRut.Text = DBGridBusqueda.Columns(0)
            TxtRut.SetFocus
        ElseIf BLinea = True Then
            TxtLin.Text = DBGridBusqueda.Columns(0)
            TxtLin.SetFocus
        Else
            TxtFicTec.Text = DBGridBusqueda.Columns(0)
            TxtFicTec.SetFocus
        End If
    End If
    FrameBusqueda.Visible = False
End Sub

Private Sub DBGridHistograma_HeadClick(ByVal ColIndex As Integer)
        RDatosHistograma.Sort = RDatosHistograma.Fields(ColIndex).Name
                  
End Sub

Private Sub Form_Load()
    PFecRutIni.Value = Date
    PFecRutFin.Value = Date
End Sub

Private Sub optcab_Click(Index As Integer)
    If Index = 0 Then
        lblcab.Visible = False
        txtcab.Visible = False
        LblCab2.Visible = False
        txtcab2.Visible = False
    Else
        lblcab.Visible = True
        txtcab.Visible = True
        LblCab2.Visible = True
        txtcab2.Visible = True
        txtcab.Text = "1"
        txtcab2.Text = "1"
        txtcab.SetFocus
    End If
        
End Sub

Private Sub OptCatalogo_Click()
    lbltitulo.Caption = "Catalogo"
End Sub

Private Sub OptFicha_Click()
    lbltitulo.Caption = "Ficha"
End Sub


Private Sub TxtBusqueda_Change()
            Set RBusqueda = New ADODB.Recordset
            'DESCRIPCION
            If OptBusqueda.Item(0).Value = True Then
                If BRutina = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Rutina, Descrip From Rutinas where Descrip Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Rutina, Descrip From Rutinas where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                ElseIf BFicha = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Descrip Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                ElseIf BCatalogo = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From VariablesDescripcion where DescripcionVariable Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From VariablesDescripcion where UPPER(DescripcionVariable) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                ElseIf BLinea = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where Descrip Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where UPPER(Descrip) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                End If
            'CODIGO
            ElseIf OptBusqueda.Item(1).Value = True Then
                If BRutina = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Rutina, Descrip From Rutinas where Rutina Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Rutina, Descrip From Rutinas where UPPER(Rutina) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                ElseIf BFicha = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where Esp_Tec Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip From FichaTecnica where UPPER(Esp_Tec) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                ElseIf BCatalogo = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From CodigoVariable where DescripcionVariable Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable From CodigoVariable where UPPER(DescripcionVariable) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                ElseIf BLinea = True Then
                        If GOrigenDeDatos = "AmaproAccess" Then
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where Linea Like '%" & TxtBusqueda.Text & "%'")
                        Else
                            Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip From Lineas where UPPER(Linea) Like '%" & UCase(TxtBusqueda.Text) & "%'")
                        End If
                        
                End If
            End If
                    
                    Set DBGridBusqueda.DataSource = RBusqueda
                    DBGridBusqueda.Columns(1).Width = "5000"

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

Private Sub TxtCab_GotFocus()
        txtcab.SelStart = 0
        txtcab.SelLength = Len(txtcab.Text)
End Sub

Private Sub TxtCab_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub txtcab2_GotFocus()
        txtcab2.SelStart = 0
        txtcab2.SelLength = Len(txtcab2.Text)
End Sub

Private Sub txtcab2_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtFicTec_Change()
    Set RBuscaFicha = New ADODB.Recordset
    If OptFicha.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFicha, "Select Descrip From Fichatecnica Where Esp_Tec = '" & TxtFicTec.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaFicha, "Select Descrip From Fichatecnica Where UPPER(Esp_Tec) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            If RBuscaFicha.RecordCount > 0 Then
                LblFichaTecnica.Caption = RBuscaFicha(0)
            Else
                LblFichaTecnica.Caption = ""
            End If
    Else
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaFicha, "Select DescripcionVariable From VariablesDescripcion Where CodigoVariable = '" & TxtFicTec.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaFicha, "Select DescripcionVariable From VariablesDescripcion Where UPPER(CodigoVariable) = '" & UCase(TxtFicTec.Text) & "'")
            End If
            If RBuscaFicha.RecordCount > 0 Then
                LblFichaTecnica.Caption = RBuscaFicha(0)
            Else
                LblFichaTecnica.Caption = ""
            End If
    End If
End Sub

Private Sub TxtFicTec_DblClick()
    BRutina = False
    BLinea = False
    If OptCatalogo.Value = True Then
        BFicha = False
        BCatalogo = True
    ElseIf OptFicha.Value = True Then
        BFicha = True
        BCatalogo = False
    End If
    Set RBusqueda = New ADODB.Recordset
    If OptCatalogo.Value = True Then
        Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable from VariablesDescripcion")
    ElseIf OptFicha.Value = True Then
        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Activa = -1")
    End If
    Set DBGridBusqueda.DataSource = RBusqueda
    DBGridBusqueda.Columns(1).Width = "5000"
    FrameBusqueda.Visible = True
    TxtBusqueda.SetFocus
End Sub

Private Sub TxtFicTec_GotFocus()
        TxtFicTec.SelStart = 0
        TxtFicTec.SelLength = Len(TxtFicTec.Text)
End Sub

Private Sub TxtFicTec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{tab}"
End If
    
    
If KeyAscii = 43 Then

    BRutina = False
    BLinea = False
    If OptCatalogo.Value = True Then
        BFicha = False
        BCatalogo = True
    ElseIf OptFicha.Value = True Then
        BFicha = True
        BCatalogo = False
    End If
    Set RBusqueda = New ADODB.Recordset
    If OptCatalogo.Value = True Then
        Call Abrir_Recordset(RBusqueda, "Select CodigoVariable, DescripcionVariable from VariablesDescripcion")
    ElseIf OptFicha.Value = True Then
        Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Nombre_Comercial from FichaTecnica Where Activa = -1")
    End If
    
    Set DBGridBusqueda.DataSource = RBusqueda
    DBGridBusqueda.Columns(1).Width = "5000"
    FrameBusqueda.Visible = True
    TxtBusqueda.SetFocus
    
End If

End Sub

Private Sub TxtFicTec_LostFocus()
    TxtFicTec.Text = UCase(TxtFicTec.Text)
End Sub

Private Sub TxtLin_Change()
    Set RBuscaLinea = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLin.Text & "'")
        Else
            Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
        End If
    If RBuscaLinea.RecordCount > 0 Then
        LblLineas.Caption = RBuscaLinea!Descrip
    Else
        LblLineas.Caption = ""
    End If
    
End Sub

Private Sub Txtlin_DblClick()

    BRutina = False
    BFicha = False
    BCatalogo = False
    BLinea = True
    Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas")
    Set DBGridBusqueda.DataSource = RBusqueda
    DBGridBusqueda.Columns(1).Width = "5000"
    FrameBusqueda.Visible = True
    TxtBusqueda.SetFocus
    


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
            
                BRutina = False
                BFicha = False
                BCatalogo = False
                BLinea = True
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select Linea, Descrip from Lineas")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "5000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
                
            End If
            
End Sub

Private Sub TxtRut_Change()
    Set RBuscaRutina = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaRutina, "Select Descrip From Rutinas Where Rutina = '" & TxtRut.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaRutina, "Select Descrip From Rutinas Where UPPER(Rutina) = '" & UCase(TxtRut.Text) & "'")
        End If
    If RBuscaRutina.RecordCount > 0 Then
            LblRutina.Caption = RBuscaRutina(0)
    Else
            LblRutina.Caption = ""
    End If
    
    
End Sub

Private Sub TxtRut_DblClick()
    BRutina = True
    BFicha = False
    BCatalogo = False
    BLinea = False
    Set RBusqueda = New ADODB.Recordset
        Call Abrir_Recordset(RBusqueda, "Select Rutina, Descrip from Rutinas")
        Set DBGridBusqueda.DataSource = RBusqueda
        DBGridBusqueda.Columns(1).Width = "5000"
        FrameBusqueda.Visible = True
        TxtBusqueda.SetFocus
    
    
    
End Sub

Private Sub TxtRut_GotFocus()
        TxtRut.SelStart = 0
        TxtRut.SelLength = Len(TxtRut.Text)
End Sub

Private Sub TxtRut_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
            
            
        If KeyAscii = 43 Then
        
            BRutina = True
            BFicha = False
            BCatalogo = False
            BLinea = False
            Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select Rutina, Descrip from Rutinas")
                Set DBGridBusqueda.DataSource = RBusqueda
                DBGridBusqueda.Columns(1).Width = "5000"
                FrameBusqueda.Visible = True
                TxtBusqueda.SetFocus
        End If
End Sub


Public Sub ImprimeHistograma()
On Error Resume Next

Conexion.Execute ("Delete from Histograma")


Set RHistograma = New ADODB.Recordset
                      
               If GOrigenDeDatos = "AmaproAccess" Then
                                     'POR UNA LINEA
                                     If ChkLin.Value = 1 Then
                                           If OptFicha.Value = True Then
                                                   Call Abrir_Recordset(RHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas As CR, FichaTecnica as FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and Cr.Esp_Tec = Ft.Esp_Tec And CR.Esp_Tec = '" & TxtFicTec.Text & "' And CR.Valor > 0 And Linea = '" & TxtLin.Text & "' Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                           Else
                                                   Call Abrir_Recordset(RHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas As CR, FichaTecnica as FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 And Linea = '" & TxtLin.Text & "' Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                           End If
                                     'TODAS LAS LINEAS
                                     Else
                                           If OptFicha.Value = True Then
                                                   Call Abrir_Recordset(RHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas As CR, FichaTecnica as FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and Cr.Esp_Tec = Ft.Esp_Tec And CR.Esp_Tec = '" & TxtFicTec.Text & "' And CR.Valor > 0 Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                           Else
                                                   Call Abrir_Recordset(RHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas As CR, FichaTecnica as FT where CR.Fec_Rut >= #" & Format(PFecRutIni.Value, "mm/dd/yyyy") & "# and CR.Fec_Rut <= #" & Format(PFecRutFin.Value, "mm/dd/yyyy") & "# and CR.Rutina = '" & TxtRut.Text & "' and CR.Esp_Tec = FT.Esp_tec And FT.Variables = '" & TxtFicTec.Text & "' And CR.Valor > 0 Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                           End If
                                     End If
                                               Do Until RHistograma.EOF
                                                           VTexto = "'" & RHistograma(0) & "', #"
                                                           VTexto = VTexto & Format(RHistograma(1), "mm/dd/yyyy") & "#, '"
                                                           VTexto = VTexto & RHistograma(2) & "', " 'HORA
                                                           VTexto = VTexto & RHistograma(3) & ", " 'CABEZAL
                                                           VTexto = VTexto & RHistograma(4) & ", '" 'VALOR
                                                           VTexto = VTexto & RHistograma(5) & "', '" 'FICHA
                                                           VTexto = VTexto & RHistograma(6) & "'" 'DESCRIPCION
                                                           
                                                           Conexion.Execute "Insert Into Histograma Values(" & VTexto & ")"
                                                           
                                                           If Err <> 0 Then
                                                                MsgBox "error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                                Err.Clear
                                                           End If
                                                           
                                                   RHistograma.MoveNext
                                               Loop
                Else 'ORACLE
                                'POR UNA LINEA
                                     If ChkLin.Value = 1 Then
                                           If OptFicha.Value = True Then
                                                   Call Abrir_Recordset(RHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas As CR, FichaTecnica as FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Cr.Esp_Tec) = UPPER(Ft.Esp_Tec) And UPPER(CR.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                           Else
                                                   Call Abrir_Recordset(RHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas As CR, FichaTecnica as FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Cr.Esp_Tec) = UPPER(Ft.Esp_Tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 And UPPER(Linea) = '" & UCase(TxtLin.Text) & "' Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                           End If
                                     'TODAS LAS LINEAS
                                     Else
                                           If OptFicha.Value = True Then
                                                   Call Abrir_Recordset(RHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas As CR, FichaTecnica as FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Cr.Esp_Tec) = UPPER(Ft.Esp_Tec) And UPPER(CR.Esp_Tec) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                           Else
                                                   Call Abrir_Recordset(RHistograma, "Select CR.Linea, CR.Fec_Rut, CR.Hor_Rut, CR.Cabezal, CR.Valor, CR.Esp_Tec, FT.Descrip from CapturaRutinas As CR, FichaTecnica as FT where CR.Fec_Rut >= To_Date('" & PFecRutIni.Value & "', 'dd/mm/yyyy')" & " and CR.Fec_Rut <= To_Date('" & PFecRutFin.Value & "', 'dd/mm/yyyy')" & " and UPPER(CR.Rutina) = '" & UCase(TxtRut.Text) & "' and UPPER(Cr.Esp_Tec) = UPPER(Ft.Esp_Tec) And UPPER(FT.Variables) = '" & UCase(TxtFicTec.Text) & "' And CR.Valor > 0 Order By CR.Esp_Tec, cR.Fec_Rut, CR.Valor")
                                           End If
                                     End If
                                
                                     
                                               Do Until RHistograma.EOF
                                                           VTexto = "'" & RHistograma(0) & "', "
                                                           VTexto = VTexto & "To_Date('" & RHistograma(1) & "', 'dd/mm/yyyy')" & ", '"
                                                           VTexto = VTexto & RHistograma(2) & "', " 'HORA
                                                           VTexto = VTexto & RHistograma(3) & ", " 'CABEZAL
                                                           VTexto = VTexto & RHistograma(4) & ", '" 'VALOR
                                                           VTexto = VTexto & RHistograma(5) & "', '" 'FICHA
                                                           VTexto = VTexto & RHistograma(6) & "'" 'DESCRIPCION
                                                           
                                                           Conexion.Execute "Insert Into Histograma Values(" & VTexto & ")"
                                                           
                                                           If Err <> 0 Then
                                                                MsgBox "error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Informacion"
                                                                Err.Clear
                                                           End If
                                                       
                                                   RHistograma.MoveNext
                                               Loop
                End If
 
                GTituloReporte = "Rutina = '" & TxtRut.Text & " " & LblRutina.Caption & "     " & TxtFicTec.Text & " " & LblFichaTecnica.Caption & "'"
                GNombreReporte = "\Histograma.rpt"
                FrmReporte.Show
 
                
End Sub
