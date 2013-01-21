VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form InventarioEntradas 
   BackColor       =   &H00FF8080&
   Caption         =   "Entradas A Inventario"
   ClientHeight    =   9060
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11955
   Icon            =   "InventarioEntradas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9060
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameBuscar 
      Caption         =   "Busqueda De Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8895
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   11895
      Begin MSDataGridLib.DataGrid DbGridBusqueda 
         Height          =   7815
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   13785
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
      Begin VB.CommandButton Command1 
         Height          =   615
         Left            =   11040
         Picture         =   "InventarioEntradas.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Sale de Lista"
         Top             =   240
         Width           =   735
      End
      Begin VB.OptionButton OptDescripcion 
         Caption         =   "Descripcion"
         Height          =   195
         Left            =   120
         TabIndex        =   30
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptCodigo 
         Caption         =   "Codigo"
         Height          =   195
         Left            =   1560
         TabIndex        =   29
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtBuscar 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   4935
      End
   End
   Begin TabDlg.SSTab TabEntradas 
      Height          =   8895
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   15690
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   1058
      BackColor       =   16744576
      TabCaption(0)   =   "Encabezado"
      TabPicture(0)   =   "InventarioEntradas.frx":293C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "FrameEncabezado"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Detalle"
      TabPicture(1)   =   "InventarioEntradas.frx":2D8E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "FrameDetalle"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "DBGridDetalle"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin MSDataGridLib.DataGrid DBGridDetalle 
         Height          =   3375
         Left            =   240
         TabIndex        =   83
         Top             =   3360
         Width           =   11415
         _ExtentX        =   20135
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
         ColumnCount     =   21
         BeginProperty Column00 
            DataField       =   "Documento"
            Caption         =   "Documento"
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
            DataField       =   "FechaProduccion"
            Caption         =   "Fecha Produccion"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "dd/MM/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   3
            EndProperty
         EndProperty
         BeginProperty Column02 
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
         BeginProperty Column03 
            DataField       =   "FichaTecnica"
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
         BeginProperty Column05 
            DataField       =   "Tarima"
            Caption         =   "Tarima"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column06 
            DataField       =   "Estado"
            Caption         =   "Estado"
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
            DataField       =   "Calidad"
            Caption         =   "Calidad"
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
            DataField       =   "Bodega"
            Caption         =   "Bodega"
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
            DataField       =   "CantidadEntrada"
            Caption         =   "Entrada"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column10 
            DataField       =   "Saldo"
            Caption         =   "Saldo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column11 
            DataField       =   "OrdenProduccion"
            Caption         =   "Orden"
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
         BeginProperty Column12 
            DataField       =   "PesoEntrada"
            Caption         =   "Peso Entrada"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   4106
               SubFormatType   =   2
            EndProperty
         EndProperty
         BeginProperty Column13 
            DataField       =   "Batch"
            Caption         =   "Batch"
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
         BeginProperty Column14 
            DataField       =   "Barra"
            Caption         =   "Barra"
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
         BeginProperty Column15 
            DataField       =   "SerieBoleta"
            Caption         =   "Serie Boleta"
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
         BeginProperty Column16 
            DataField       =   "OrdenBoleta"
            Caption         =   "Orden Boleta"
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
         BeginProperty Column17 
            DataField       =   "BultoBoleta"
            Caption         =   "Bulto Boleta"
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
         BeginProperty Column18 
            DataField       =   "FechaBoleta"
            Caption         =   "Fecha Boleta"
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
         BeginProperty Column19 
            DataField       =   "BobinaBoleta"
            Caption         =   "Bobina Boleta"
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
         BeginProperty Column20 
            DataField       =   "Observaciones"
            Caption         =   "Observaciones"
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
               ColumnWidth     =   689.953
            EndProperty
            BeginProperty Column01 
               Locked          =   -1  'True
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column02 
               Locked          =   -1  'True
               ColumnWidth     =   434.835
            EndProperty
            BeginProperty Column03 
               Locked          =   -1  'True
               ColumnWidth     =   1184.882
            EndProperty
            BeginProperty Column04 
               Locked          =   -1  'True
               ColumnWidth     =   3404.977
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   540.284
            EndProperty
            BeginProperty Column06 
               Locked          =   -1  'True
               ColumnWidth     =   945.071
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   329.953
            EndProperty
            BeginProperty Column08 
               ColumnWidth     =   450.142
            EndProperty
            BeginProperty Column09 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   884.976
            EndProperty
            BeginProperty Column10 
               Alignment       =   1
               Locked          =   -1  'True
               ColumnWidth     =   840.189
            EndProperty
            BeginProperty Column11 
               ColumnWidth     =   1349.858
            EndProperty
            BeginProperty Column12 
               Alignment       =   1
               ColumnWidth     =   1049.953
            EndProperty
            BeginProperty Column13 
            EndProperty
            BeginProperty Column14 
               Locked          =   -1  'True
               ColumnWidth     =   2564.788
            EndProperty
            BeginProperty Column15 
            EndProperty
            BeginProperty Column16 
            EndProperty
            BeginProperty Column17 
            EndProperty
            BeginProperty Column18 
            EndProperty
            BeginProperty Column19 
            EndProperty
            BeginProperty Column20 
               ColumnWidth     =   7694.93
            EndProperty
         EndProperty
      End
      Begin VB.Frame FrameDetalle 
         Caption         =   "Detalle Entradas A Inventario"
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
         ForeColor       =   &H00FF0000&
         Height          =   8055
         Left            =   120
         TabIndex        =   46
         Top             =   720
         Width           =   11685
         Begin MSDataGridLib.DataGrid dg_OCompra_MP 
            Height          =   1215
            Left            =   120
            TabIndex        =   118
            Top             =   6120
            Width           =   11415
            _ExtentX        =   20135
            _ExtentY        =   2143
            _Version        =   393216
            BackColor       =   -2147483634
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
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
            Caption         =   "Ordenes de Compra de MP Pendientes"
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
                  LCID            =   2058
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
                  LCID            =   2058
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
         Begin VB.TextBox TxtCueTar 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            TabIndex        =   117
            Top             =   7560
            Width           =   855
         End
         Begin VB.Frame FrameDetalleCompras 
            Enabled         =   0   'False
            Height          =   2295
            Left            =   120
            TabIndex        =   47
            Top             =   240
            Width           =   11415
            Begin VB.TextBox TxtTarUlt 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
               BackColor       =   &H80000016&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   10200
               TabIndex        =   115
               Top             =   360
               Width           =   555
            End
            Begin VB.CheckBox ChkMultiplica 
               BackColor       =   &H80000016&
               Height          =   255
               Left            =   8520
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   960
               Width           =   255
            End
            Begin VB.TextBox TxtObs2 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2160
               MaxLength       =   20
               TabIndex        =   65
               Top             =   1920
               Width           =   9165
            End
            Begin VB.ComboBox CboCal 
               Height          =   315
               ItemData        =   "InventarioEntradas.frx":30A8
               Left            =   2160
               List            =   "InventarioEntradas.frx":30B5
               TabIndex        =   55
               Text            =   "A"
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox TxtEst 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   9000
               Locked          =   -1  'True
               MaxLength       =   16
               TabIndex        =   64
               TabStop         =   0   'False
               Top             =   1560
               Width           =   2295
            End
            Begin VB.TextBox TxtBobBol 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   6720
               MaxLength       =   20
               TabIndex        =   63
               Top             =   1560
               Width           =   2205
            End
            Begin VB.TextBox TxtBulBol 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3840
               MaxLength       =   20
               TabIndex        =   61
               Top             =   1560
               Width           =   1485
            End
            Begin VB.TextBox TxtOrdBol 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   2160
               MaxLength       =   20
               TabIndex        =   60
               Top             =   1560
               Width           =   1605
            End
            Begin VB.TextBox TxtSerBol 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               MaxLength       =   20
               TabIndex        =   59
               Top             =   1560
               Width           =   1965
            End
            Begin MSMask.MaskEdBox MskPesEnt 
               Height          =   285
               Left            =   8880
               TabIndex        =   57
               Top             =   960
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,###,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.TextBox TxtBarra 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0FFFF&
               Height          =   288
               Left            =   5880
               Locked          =   -1  'True
               MaxLength       =   35
               TabIndex        =   48
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.TextBox TxtCodPro 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   4800
               MaxLength       =   15
               TabIndex        =   52
               Top             =   360
               Width           =   1812
            End
            Begin VB.TextBox TxtDocDet 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   840
               TabIndex        =   71
               TabStop         =   0   'False
               Top             =   120
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox TxtBod 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   3360
               MaxLength       =   3
               TabIndex        =   56
               Top             =   960
               Width           =   555
            End
            Begin VB.TextBox TxtTar 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   10800
               TabIndex        =   53
               Top             =   360
               Width           =   555
            End
            Begin VB.TextBox TxtLin 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   3360
               MaxLength       =   2
               TabIndex        =   51
               Top             =   360
               Width           =   435
            End
            Begin VB.TextBox TxtBat 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   54
               Top             =   960
               Width           =   1935
            End
            Begin VB.TextBox TxtOrd 
               Appearance      =   0  'Flat
               Height          =   288
               Left            =   120
               MaxLength       =   15
               TabIndex        =   49
               Top             =   360
               Width           =   1935
            End
            Begin MSMask.MaskEdBox MskFecPro 
               Height          =   285
               Left            =   2160
               TabIndex        =   50
               Top             =   360
               Width           =   1155
               _ExtentX        =   2037
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   8438015
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox TxtCanPro 
               Height          =   285
               Left            =   10200
               TabIndex        =   58
               Top             =   960
               Width           =   1095
               _ExtentX        =   1931
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "#,###,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MskFecBol 
               Height          =   285
               Left            =   5400
               TabIndex        =   62
               Top             =   1560
               Width           =   1275
               _ExtentX        =   2249
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               BackColor       =   16777215
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Ultima Tarima"
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
               Index           =   29
               Left            =   9600
               TabIndex        =   116
               Top             =   120
               Width           =   1170
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "laminas x cuerpos"
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
               Index           =   8
               Left            =   7200
               TabIndex        =   114
               Top             =   720
               Width           =   1530
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000016&
               Caption         =   "Observaciones De Bulto"
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
               Index           =   28
               Left            =   120
               TabIndex        =   111
               Top             =   1920
               Width           =   2070
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Estado"
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
               Left            =   9000
               TabIndex        =   109
               Top             =   1320
               Width           =   600
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Bobina Boleta"
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
               Index           =   26
               Left            =   6720
               TabIndex        =   103
               Top             =   1320
               Width           =   1200
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Boleta"
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
               Index           =   25
               Left            =   5400
               TabIndex        =   102
               Top             =   1320
               Width           =   1140
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Bulto Boleta"
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
               Index           =   24
               Left            =   3840
               TabIndex        =   101
               Top             =   1320
               Width           =   1050
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Orden Boleta"
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
               Index           =   23
               Left            =   2160
               TabIndex        =   100
               Top             =   1320
               Width           =   1125
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Serie Boleta"
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
               Index           =   22
               Left            =   120
               TabIndex        =   99
               Top             =   1320
               Width           =   1050
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Peso Entrada"
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
               Left            =   8880
               TabIndex        =   98
               Top             =   720
               Width           =   1155
            End
            Begin VB.Label LblDes 
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
               Left            =   6720
               TabIndex        =   97
               Top             =   360
               Width           =   3375
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Orden"
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
               TabIndex        =   96
               Top             =   120
               Width           =   525
            End
            Begin VB.Label Label1 
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
               Index           =   0
               Left            =   4680
               TabIndex        =   82
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label2 
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
               Left            =   6600
               TabIndex        =   81
               Top             =   120
               Width           =   1575
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Cantidad"
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
               Left            =   10200
               TabIndex        =   80
               Top             =   720
               Width           =   765
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H80000016&
               Caption         =   "Tarima"
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
               Left            =   10800
               TabIndex        =   79
               Top             =   120
               Width           =   585
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
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
               Height          =   195
               Index           =   8
               Left            =   2160
               TabIndex        =   78
               Top             =   120
               Width           =   540
            End
            Begin VB.Label Label6 
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
               Index           =   9
               Left            =   3360
               TabIndex        =   77
               Top             =   120
               Width           =   480
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Batch"
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
               Index           =   10
               Left            =   120
               TabIndex        =   76
               Top             =   720
               Width           =   510
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Bodega"
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
               Left            =   3360
               TabIndex        =   75
               Top             =   720
               Width           =   660
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Calidad"
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
               Index           =   6
               Left            =   2160
               TabIndex        =   74
               Top             =   720
               Width           =   645
            End
            Begin VB.Label LblLin2 
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
               Left            =   3840
               TabIndex        =   73
               Top             =   360
               Width           =   855
            End
            Begin VB.Label LblBod 
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
               TabIndex        =   72
               Top             =   960
               Width           =   4335
            End
         End
         Begin VB.CommandButton CmdAgregar2 
            Caption         =   "A&gregar"
            Height          =   495
            Left            =   120
            Picture         =   "InventarioEntradas.frx":30C2
            Style           =   1  'Graphical
            TabIndex        =   66
            Top             =   7440
            Visible         =   0   'False
            Width           =   2200
         End
         Begin VB.CommandButton CmdGrabar2 
            Caption         =   "G&rabar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   2400
            Picture         =   "InventarioEntradas.frx":35F4
            Style           =   1  'Graphical
            TabIndex        =   67
            Top             =   7440
            Visible         =   0   'False
            Width           =   2200
         End
         Begin VB.CommandButton CmdTerminar 
            Caption         =   "&Terminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   9240
            Picture         =   "InventarioEntradas.frx":3B26
            TabIndex        =   70
            Top             =   7440
            Visible         =   0   'False
            Width           =   1245
         End
         Begin VB.CommandButton CmdCancelar2 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   495
            Left            =   4680
            Picture         =   "InventarioEntradas.frx":4058
            Style           =   1  'Graphical
            TabIndex        =   68
            Top             =   7440
            Visible         =   0   'False
            Width           =   2200
         End
         Begin VB.CommandButton CmdBorrar2 
            Caption         =   "B&orrar"
            Height          =   495
            Left            =   6960
            Picture         =   "InventarioEntradas.frx":458A
            Style           =   1  'Graphical
            TabIndex        =   69
            Top             =   7440
            Visible         =   0   'False
            Width           =   2200
         End
      End
      Begin VB.Frame FrameEncabezado 
         Caption         =   "Encabezado Entradas De Inventario"
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
         Height          =   7215
         Left            =   -74880
         TabIndex        =   35
         Top             =   720
         Width           =   11655
         Begin VB.CommandButton CmdCedulas 
            BackColor       =   &H0080C0FF&
            Caption         =   "Cedulas"
            Height          =   700
            Left            =   8640
            Picture         =   "InventarioEntradas.frx":4ABC
            Style           =   1  'Graphical
            TabIndex        =   112
            Top             =   6360
            Width           =   900
         End
         Begin VB.CommandButton CmdBotones2 
            BackColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   1
            Left            =   120
            MouseIcon       =   "InventarioEntradas.frx":6BE6
            Picture         =   "InventarioEntradas.frx":7028
            Style           =   1  'Graphical
            TabIndex        =   107
            ToolTipText     =   "Primer Registro"
            Top             =   6480
            Width           =   375
         End
         Begin VB.CommandButton CmdBotones2 
            BackColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   2
            Left            =   480
            MouseIcon       =   "InventarioEntradas.frx":755A
            Picture         =   "InventarioEntradas.frx":799C
            Style           =   1  'Graphical
            TabIndex        =   106
            ToolTipText     =   "Registro Anterior"
            Top             =   6480
            Width           =   375
         End
         Begin VB.CommandButton CmdBotones2 
            BackColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   3
            Left            =   10800
            MouseIcon       =   "InventarioEntradas.frx":7ECE
            Picture         =   "InventarioEntradas.frx":8310
            Style           =   1  'Graphical
            TabIndex        =   105
            ToolTipText     =   "Siguiente Registro"
            Top             =   6480
            Width           =   375
         End
         Begin VB.CommandButton CmdBotones2 
            BackColor       =   &H00C0C0C0&
            Height          =   465
            Index           =   4
            Left            =   11160
            MouseIcon       =   "InventarioEntradas.frx":8842
            Picture         =   "InventarioEntradas.frx":8C84
            Style           =   1  'Graphical
            TabIndex        =   104
            ToolTipText     =   "Ultimo Registro"
            Top             =   6480
            Width           =   375
         End
         Begin VB.Frame FrameCompras 
            Enabled         =   0   'False
            Height          =   6015
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   11415
            Begin VB.TextBox TxtLib 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   600
               Width           =   1575
            End
            Begin VB.TextBox TxtEstado 
               Appearance      =   0  'Flat
               BackColor       =   &H0080C0FF&
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   12
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   960
               Width           =   1575
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   7
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   15
               Top             =   5640
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   6
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   14
               Top             =   5280
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   5
               Left            =   1560
               MaxLength       =   30
               TabIndex        =   13
               Top             =   4920
               Width           =   3495
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   4
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   6
               Top             =   2400
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   3
               Left            =   1560
               MaxLength       =   15
               TabIndex        =   5
               Top             =   2040
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   2
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   4
               Top             =   1680
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   1
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   3
               Top             =   1320
               Width           =   1215
            End
            Begin VB.TextBox TxtTexto 
               Appearance      =   0  'Flat
               Height          =   285
               Index           =   0
               Left            =   1560
               MaxLength       =   10
               TabIndex        =   2
               Top             =   960
               Width           =   1215
            End
            Begin VB.TextBox TxtDocIng 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   0
               TabStop         =   0   'False
               Top             =   240
               Width           =   1215
            End
            Begin VB.TextBox TxtBodega 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               MaxLength       =   3
               TabIndex        =   9
               Top             =   3480
               Width           =   1215
            End
            Begin VB.TextBox TxtBatch 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1560
               TabIndex        =   10
               Top             =   3840
               Width           =   1215
            End
            Begin VB.TextBox TxtReq 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   9720
               Locked          =   -1  'True
               MaxLength       =   20
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   240
               Width           =   1575
            End
            Begin VB.TextBox TxtObs 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1560
               MaxLength       =   100
               TabIndex        =   12
               Top             =   4560
               Width           =   6855
            End
            Begin VB.CheckBox ChkProInt 
               Caption         =   "Produccion"
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
               Left            =   1560
               TabIndex        =   7
               Top             =   2760
               Width           =   1335
            End
            Begin VB.CheckBox ChkProLib 
               Caption         =   "Produccion Liberada"
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
               Left            =   1560
               TabIndex        =   8
               Top             =   3120
               Width           =   2175
            End
            Begin VB.TextBox TxtLinea 
               Appearance      =   0  'Flat
               BackColor       =   &H008080FF&
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
               Left            =   1560
               MaxLength       =   2
               TabIndex        =   11
               Top             =   4200
               Width           =   1215
            End
            Begin MSMask.MaskEdBox MskFec 
               Height          =   285
               Left            =   1560
               TabIndex        =   1
               Top             =   600
               Width           =   1215
               _ExtentX        =   2143
               _ExtentY        =   503
               _Version        =   393216
               Appearance      =   0
               Format          =   "dd/mm/yyyy"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Liberado"
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
               Index           =   27
               Left            =   8880
               TabIndex        =   110
               Top             =   600
               Width           =   750
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Estado"
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
               Left            =   9000
               TabIndex        =   108
               Top             =   960
               Width           =   600
            End
            Begin VB.Label LblTra 
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
               Left            =   2880
               TabIndex        =   95
               Top             =   2400
               Width           =   5535
            End
            Begin VB.Label LblTipDoc 
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
               Left            =   2880
               TabIndex        =   94
               Top             =   1680
               Width           =   5535
            End
            Begin VB.Label LblPro 
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
               Left            =   2880
               TabIndex        =   93
               Top             =   1320
               Width           =   5535
            End
            Begin VB.Label LblTipEnt 
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
               Left            =   2880
               TabIndex        =   92
               Top             =   960
               Width           =   5535
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Entrada"
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
               Index           =   21
               Left            =   120
               TabIndex        =   91
               Top             =   960
               Width           =   1110
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Proveedor"
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
               Left            =   120
               TabIndex        =   90
               Top             =   1320
               Width           =   885
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Tipo Documento"
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
               Left            =   120
               TabIndex        =   89
               Top             =   1680
               Width           =   1410
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "# Documento"
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
               TabIndex        =   88
               Top             =   2040
               Width           =   1155
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Transportista"
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
               Index           =   17
               Left            =   120
               TabIndex        =   87
               Top             =   2400
               Width           =   1125
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Nombre Piloto"
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
               Left            =   120
               TabIndex        =   86
               Top             =   4920
               Width           =   1200
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Placas Camion"
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
               Index           =   15
               Left            =   120
               TabIndex        =   85
               Top             =   5280
               Width           =   1260
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Placas Furgon"
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
               Index           =   14
               Left            =   120
               TabIndex        =   84
               Top             =   5640
               Width           =   1230
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Fecha Entrada"
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
               TabIndex        =   45
               Top             =   600
               Width           =   1260
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Transaccion"
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
               TabIndex        =   44
               Top             =   240
               Width           =   1065
            End
            Begin VB.Label Label6 
               Caption         =   "Bodega"
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
               Index           =   3
               Left            =   120
               TabIndex        =   43
               Top             =   3480
               Width           =   975
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
               Left            =   2880
               TabIndex        =   42
               Top             =   3480
               Width           =   5535
            End
            Begin VB.Label Label6 
               Caption         =   "No. Batch"
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
               Index           =   13
               Left            =   120
               TabIndex        =   41
               Top             =   3840
               Width           =   975
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Requerido"
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
               Index           =   12
               Left            =   8760
               TabIndex        =   40
               Top             =   240
               Width           =   885
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Observaciones"
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
               Index           =   4
               Left            =   120
               TabIndex        =   39
               Top             =   4560
               Width           =   1275
            End
            Begin VB.Label Label6 
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
               Index           =   11
               Left            =   120
               TabIndex        =   38
               Top             =   4200
               Width           =   615
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
               Left            =   2880
               TabIndex        =   37
               Top             =   4200
               Width           =   5535
            End
         End
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar"
            Height          =   700
            Left            =   960
            Picture         =   "InventarioEntradas.frx":91B6
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   6360
            Width           =   900
         End
         Begin VB.CommandButton CmdGrabar 
            Caption         =   "&Grabar"
            Enabled         =   0   'False
            Height          =   700
            Left            =   2880
            Picture         =   "InventarioEntradas.frx":9533
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   6360
            Width           =   900
         End
         Begin VB.CommandButton CmdCancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   700
            Left            =   3840
            Picture         =   "InventarioEntradas.frx":9A8F
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   6360
            Width           =   900
         End
         Begin VB.CommandButton CmdBorrar 
            Caption         =   "&Borrar"
            Height          =   700
            Left            =   4800
            Picture         =   "InventarioEntradas.frx":9FC6
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   6360
            Width           =   900
         End
         Begin VB.CommandButton CmdSalida 
            Appearance      =   0  'Flat
            Height          =   700
            Left            =   9600
            Picture         =   "InventarioEntradas.frx":A58E
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Salida"
            Top             =   6360
            Width           =   1020
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "B&uscar"
            Height          =   700
            Left            =   5760
            Picture         =   "InventarioEntradas.frx":AAA9
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   6360
            Width           =   900
         End
         Begin VB.CommandButton CmdEditar 
            Caption         =   "&Editar"
            Height          =   700
            Left            =   1920
            Picture         =   "InventarioEntradas.frx":AF31
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   6360
            Width           =   900
         End
         Begin VB.CommandButton CmdImprimir 
            BackColor       =   &H0080C0FF&
            Caption         =   "Entrada"
            Height          =   700
            Left            =   6720
            Picture         =   "InventarioEntradas.frx":B308
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   6360
            Width           =   900
         End
         Begin VB.CommandButton CmdImprimirAmarillas 
            BackColor       =   &H0080C0FF&
            Caption         =   "Amarillas"
            Height          =   700
            Left            =   7680
            Picture         =   "InventarioEntradas.frx":B842
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   6360
            Width           =   900
         End
      End
   End
End
Attribute VB_Name = "InventarioEntradas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mensaje As String
Dim VDocumento As Double
Dim VDocumentoDetalle As Double
Dim VCantidad As Currency
Dim VCodigoProducto As String
Dim VCantidadProducto As Currency
Dim VBodega As String
Dim VBatch As Double
Dim VClasificacion As String

Dim Bandera As Boolean
Dim Bandera2 As Boolean
Dim Bandera3 As Boolean
Dim Bandera4 As Boolean

Dim BanderaOCompraMP As Boolean
Dim BGrabo As Boolean
Dim BEdito As Boolean
Dim BBorro As Boolean

'Dim BanderaBodegaMP As Boolean

Dim BBodega As Boolean
Dim BProducto As Boolean
Dim BBodegaDetalle As Boolean
Dim BLineas As Boolean
Dim BTipoEntrada As Boolean
Dim BProveedor As Boolean
Dim BTipoDocumento As Boolean
Dim BTransportista As Boolean
Dim BProduccionInterna As Boolean
Dim BProduccionLiberada As Boolean


Dim RBuscaProducto As New ADODB.Recordset
Dim RMaximo As New ADODB.Recordset
Dim RBuscaBodega As New ADODB.Recordset
Dim RBuscaDetalle As New ADODB.Recordset
Dim RBuscaEncabezado As New ADODB.Recordset
Dim RBuscaProduccion As New ADODB.Recordset
Dim RBuscaTarima As New ADODB.Recordset
Dim RCuentaTarimas As New ADODB.Recordset
Dim RBuscaLinea As New ADODB.Recordset
Dim RBuscaFichaOrden As New ADODB.Recordset

Dim RBusqueda As New ADODB.Recordset
Dim REncabezado As New ADODB.Recordset
Dim RDetalle As New ADODB.Recordset
Dim RBuscaTipoEntrada As New ADODB.Recordset
Dim RBuscaProveedor As New ADODB.Recordset
Dim RBuscaTipoDocumento As New ADODB.Recordset
Dim RBuscaTransportista As New ADODB.Recordset
Dim RBuscaTipoInventario As New ADODB.Recordset

Dim RBuscaCuerpos As New ADODB.Recordset
Dim RBuscaUltimaTarima As New ADODB.Recordset

Dim RBuscaOrdenCompraMP As New ADODB.Recordset

Dim VUltimaFichaTecnica As String
Dim VUltimaEnvases As Currency
Dim VUltimaFecha As String
Dim VUltimaTarima As Integer
Dim VUltimaLinea As String
Dim VUltimaCalidad As String
Dim VLinea As String
Dim VUltimaOrden As String

Dim VUltimaBobina As String
Dim VUltimaSerie As String
Dim VUltimaOrden2 As String
Dim VUltimaBulto As String
Dim VUltimaFecha2 As String

Dim VTexto As String
Dim BEditarEncabezado As Boolean
Dim BEditarDetalle As Boolean


Sub Botones1()
    If Bandera = True Then
         FrameCompras.Enabled = True
         CmdAgregar.Enabled = False
         CmdEditar.Enabled = False
         CmdGrabar.Enabled = True
         CmdBorrar.Enabled = False
         CmdCancelar.Enabled = True
         CmdBuscar.Enabled = False
         CmdSalida.Enabled = False
         CmdImprimir.Enabled = False
         CmdImprimirAmarillas.Enabled = False
         CmdBotones2.Item(1).Visible = False
         CmdBotones2.Item(2).Visible = False
         CmdBotones2.Item(3).Visible = False
         CmdBotones2.Item(4).Visible = False
         
    Else
         FrameCompras.Enabled = False
         CmdAgregar.Enabled = True
         CmdEditar.Enabled = True
         CmdGrabar.Enabled = False
         CmdBorrar.Enabled = True
         CmdCancelar.Enabled = False
         CmdBuscar.Enabled = True
         CmdSalida.Enabled = True
         CmdImprimir.Enabled = True
         CmdImprimirAmarillas.Enabled = True
         CmdBotones2.Item(1).Visible = True
         CmdBotones2.Item(2).Visible = True
         CmdBotones2.Item(3).Visible = True
         CmdBotones2.Item(4).Visible = True
         
         
    End If
End Sub

Sub Botones2()
    If Bandera2 = True Then
         FrameDetalleCompras.Enabled = True
         CmdAgregar2.Enabled = False
         CmdGrabar2.Enabled = True
         CmdTerminar.Enabled = False
         CmdBorrar2.Enabled = False
         CmdCancelar2.Enabled = True
         
    Else
         FrameDetalleCompras.Enabled = False
         CmdAgregar2.Enabled = True
         CmdGrabar2.Enabled = False
         CmdTerminar.Enabled = True
         CmdBorrar2.Enabled = True
         CmdCancelar2.Enabled = False
         
    End If

End Sub

Sub BotonesVisiblesDetalle()
    If Bandera3 = True Then
         CmdAgregar2.Visible = True
         
         CmdGrabar2.Visible = True
         CmdTerminar.Visible = True
         CmdBorrar2.Visible = True
         CmdCancelar2.Visible = True
    Else
         CmdAgregar2.Visible = False

         CmdGrabar2.Visible = False
         CmdTerminar.Visible = False
         CmdBorrar2.Visible = False
         CmdCancelar2.Visible = False
    End If

End Sub


Private Sub ChkProInt_Click()
        If ChkProInt.Value = 1 Then
            ChkProLib.Value = 0
        End If
End Sub

Private Sub ChkProInt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub ChkProLib_Click()
        If ChkProLib.Value = 1 Then
            ChkProInt.Value = 0
        End If
End Sub

Private Sub ChkProLib_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub Cmd_Cancelar_Click()
    'Frame_Compras_MP.Visible = False
End Sub

Private Sub Cmd_Salir_Click()
    'Frame_Compras_MP.Visible = False
End Sub

Private Sub CmdAgregar2_Click()
    
    ' VER STATUS DE QUE HIZO EL USUARIO
    BGrabo = True
    BEdito = False
    BBorro = False
    
    Bandera2 = True
    Botones2
    Limpia_CamposDetalle
    
    'INABILITA EL GRID PARA QUE NO PUEDAN MOVERSE POR EL GRID
    DBGridDetalle.Enabled = False
    
    BEditarDetalle = False
    TxtDocDet.Text = VDocumento
    
    'ASIGNA LOS DATOS DEL ENCABEZADO
    TxtBod.Text = VBodega
    TxtBat.Text = VBatch
    
    'ASIGNA LOS ULTIMOS DATOS DIGITADOS
    TxtCodPro.Text = VUltimaFichaTecnica
    TxtCanPro.Text = VUltimaEnvases
    TxtLin.Text = VUltimaLinea
    MskFecPro.Text = VUltimaFecha
    TxtTar.Text = VUltimaTarima
    CboCal.Text = VUltimaCalidad
    TxtOrd.Text = VUltimaOrden
    
    TxtBobBol.Text = VUltimaBobina
    TxtSerBol.Text = VUltimaSerie
    TxtOrdBol.Text = VUltimaOrden2
    TxtBulBol.Text = VUltimaBulto
    MskFecBol.Text = VUltimaFecha2
    
    TxtEst.Text = "NO INSPECCIONADO"
   
    TxtOrd.SetFocus
    LblDes.Caption = ""
End Sub


Private Sub CmdBorrar_Click()
On Error Resume Next
            If GBorrar = True Then
                'NO HACE NADA PORQUE SI TIENE ACCESO
            ElseIf TxtEstado.Text = "LIBERADO" Then
                'VERIFICA SI YA FUE LIBERADA LA ENTRADA
                    MsgBox "Este Documento No Se Puede BORRAR Porque Ya Fue Liberado", vbOKOnly + vbExclamation, "Informacion"
                    Exit Sub
            End If
            
            VDocumento = TxtDocIng.Text
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                MousePointer = 11
                        
                    REncabezado.Delete
                    
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147217873 Then
                                
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                    
                    
                    'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        REncabezado.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        REncabezado.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
                        
                        Llena_CamposEncabezado
                        
                            Set RDetalle = New ADODB.Recordset
                                 If GOrigenDeDatos = "AmaproAccess" Then
                                     Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & TxtDocIng.Text & " And E.Documento = D.Documento And D.FichaTecnica = F.Esp_Tec")
                                 Else 'ORACLE
                                     Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & TxtDocIng.Text & " And E.Documento = D.Documento And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
                                 End If
                                        Llena_CamposDetalle
                                        Set DBGridDetalle.DataSource = RDetalle

                    
                MousePointer = 0
            End If
      
                
            
            
            
End Sub

Private Sub CmdBorrar2_Click()
On Error Resume Next
            
    ' VER STATUS DE QUE HIZO EL USUARIO
    BGrabo = False
    BEdito = False
    BBorro = True
            
            'ASIGANMOS A UNA VARIABLE EL DOCUMENTO DETALLE
            VDocumentoDetalle = TxtDocDet.Text
            VBodega = TxtBod.Text
            VCodigoProducto = TxtCodPro.Text
            VCantidad = TxtCanPro.Text
    
            mensaje = MsgBox("¿Está seguro de Borrar el registro?", vbOKCancel + vbCritical + vbDefaultButton2, "Eliminación de Registros")

            'SI CONTESTA QUE SI QUIERE BORRAR
            If mensaje = vbOK Then
                
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                                Conexion.Execute "Delete From DetalleEntradasInventario Where FichaTecnica = '" & TxtCodPro.Text & "' And Tarima = " & TxtTar.Text & " And FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Linea = '" & TxtLin.Text & "'"
                        Else 'ORACLE
                                Conexion.Execute "Delete From DetalleEntradasInventario Where UPPER(FichaTecnica) = '" & UCase(TxtCodPro.Text) & "' And Tarima = " & TxtTar.Text & " And FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'"
                        End If
                        
                        If GOrigenDeDatos = "AmaproAccess" Then
                            If Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        Else 'ORACLE
                            'SI HAY ERRORES
                            If Err = -2147217873 Then
                                Conexion.RollbackTrans
                                MsgBox "No Se Puede Borrar Porque Tiene Registros Relacionados ", vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            ElseIf Err <> -2147217873 And Err <> 0 Then
                                Conexion.RollbackTrans
                                MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                Err.Clear
                            End If
                        End If
                    
                    
                    'VUELVE A LLENAR EL RECORDSET DE SU ESTADO ORIGINAL
                        RDetalle.Requery
                        'MUEVE AL SIGUIENTE REGISTRO
                        RDetalle.MoveLast
                        'SI HAY ERRORES
                        If Err <> 0 Then
                            'MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Err.Clear
                        End If
            End If
                                        
                   
            
End Sub

Private Sub CmdBotones2_Click(Index As Integer)
On Error Resume Next
MousePointer = 11
    If Index = 1 Then
        REncabezado.MoveFirst
    'REGISTRO ANTERIOR
    ElseIf Index = 2 Then
        REncabezado.MovePrevious
    'SIGUIENTE REGISTRO
    ElseIf Index = 3 Then
        REncabezado.MoveNext
    'ULTIMO REGISTRO
    ElseIf Index = 4 Then
        REncabezado.MoveLast
    End If
    
    'SI LLEGA AL PRIMERO O FINAL DEL REGISTRO
    If REncabezado.BOF Then
        REncabezado.MoveFirst
    ElseIf REncabezado.EOF Then
        REncabezado.MoveLast
    End If
    
    If Err <> 0 Then
    End If
    
    'SI PRESIONA LOS BOTONES DE SIGUIENTE O ANTERIOR O PRIMER O ULTIMO REGISTRO
    Llena_CamposEncabezado
    
            Set RDetalle = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & TxtDocIng.Text & " And E.Documento = D.Documento And D.FichaTecnica = F.Esp_Tec")
                Else 'ORACLE
                    Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & TxtDocIng.Text & " And E.Documento = D.Documento And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
                End If
                
                Llena_CamposDetalle
                Set DBGridDetalle.DataSource = RDetalle
    
MousePointer = 0

End Sub

Private Sub CmdBuscar_Click()
On Error Resume Next
    mensaje = InputBox("Transaccion a Buscar")
    If IsNumeric(mensaje) Then
                REncabezado.MoveFirst
                REncabezado.Find " Documento = " & mensaje
                                                
                'Set REncabezado = New ADODB.Recordset
                '        If GOrigenDeDatos = "AmaproAccess" Then
                '            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoEntradasInventario Where Documento = " & mensaje)
                '        Else 'ORACLE
                '            Call Abrir_Recordset(REncabezado, "Select * From EncabezadoEntradasInventario Where Documento = " & mensaje)
                '        End If
    
                   
                  '      Llena_CamposEncabezado
                                        
                Llena_CamposEncabezado
                
                        Set RDetalle = New ADODB.Recordset
                                If GOrigenDeDatos = "AmaproAccess" Then
                                     Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & TxtDocIng.Text & " And E.Documento = D.Documento And D.FichaTecnica = F.Esp_Tec")
                                Else 'ORACLE
                                     Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & TxtDocIng.Text & " And E.Documento = D.Documento And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
                                End If
                                Llena_CamposDetalle
                                Set DBGridDetalle.DataSource = RDetalle
    End If
    
    
End Sub

Private Sub CmdCancelar_Click()
On Error Resume Next
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    Llena_CamposEncabezado
    FrameDetalle.Visible = True
    DBGridDetalle.Visible = True
    
End Sub

Private Sub CmdCancelar2_Click()
On Error Resume Next
    
    DBGridDetalle.Enabled = True
    Bandera2 = False
    Botones2
    Llena_CamposDetalle
    
    ' VER STATUS DE QUE HIZO EL USUARIO
    'BGrabo = False
    'BEdito = False
    'BBorro = False
        
End Sub


Private Sub CmdEditar_Click()
On Error Resume Next
    
    BEditarEncabezado = True
    
    ' VER STATUS DE QUE HIZO EL USUARIO
    BGrabo = False
    BEdito = True
    BBorro = False
    
    'VALIDA SI TIENE ACCESO
    
            If GEditar = True Then
            Else
                If TxtEstado.Text = "LIBERADO" Then
                    MsgBox "Transaccion Ya Esta Liberada", vbOKOnly + vbExclamation, "Informacion"
                    Exit Sub
                Else
                End If
            End If
    

    
    Bandera = True
    Botones1
    MskFec.SetFocus
    FrameDetalle.Visible = False
    DBGridDetalle.Visible = False
    TxtReq.Text = GUsuario
        
    
End Sub



Private Sub CmdGrabar2_Click()
On Error Resume Next
    
    Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLin.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
               
            Else
                MsgBox "Codigo Linea No Existe", vbOKOnly + vbInformation, "Inforamcion"
                TxtLin.SetFocus
                Exit Sub
            End If

    Set RBuscaProducto = New ADODB.Recordset
        If GOrigenDeDatos = "AmaproAccess" Then
            Call Abrir_Recordset(RBuscaProducto, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCodPro.Text & "'")
        Else 'ORACLE
            Call Abrir_Recordset(RBuscaProducto, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodPro.Text) & "'")
        End If
        If RBuscaProducto.RecordCount > 0 Then
        
        Else
            MsgBox "Codigo Ficha Tecnica No Existe", vbOKOnly + vbInformation, "Inforamcion"
            TxtCodPro.SetFocus
            Exit Sub
        End If
    
    
    'GUARDA VARIABLES
    VCantidad = TxtCanPro.Text
    VCodigoProducto = TxtCodPro.Text
    
    VUltimaFichaTecnica = TxtCodPro.Text
    VUltimaEnvases = TxtCanPro.Text
    VUltimaLinea = TxtLin.Text
    VUltimaFecha = MskFecPro.Text
    VUltimaTarima = Val(TxtTar.Text) + 1
    VUltimaCalidad = CboCal.Text
    VUltimaOrden = TxtOrd.Text
    VBodega = TxtBod.Text
    
    VUltimaBobina = TxtBobBol.Text
    VUltimaSerie = TxtSerBol.Text
    VUltimaOrden2 = TxtOrdBol.Text
    VUltimaBulto = TxtBulBol.Text
    VUltimaFecha2 = MskFecBol.Text
        
    'REVISAMOS DATOS
    If Not IsNumeric(TxtCanPro.Text) Then
       MsgBox "Cantidad De Producto Incorrecta", vbOKOnly + vbCritical, "Error"
       TxtCanPro.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS EL BATCH DE DETALLE
    If Not IsNumeric(TxtBat.Text) Then
       MsgBox "Numero De Bath Incorrecto", vbOKOnly + vbCritical, "Error"
       TxtBat.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS LA TARIMA
    If Not IsNumeric(TxtTar.Text) Then
       MsgBox "Numero De Tarima Incorrecto", vbOKOnly + vbCritical, "Error"
       TxtBat.SetFocus
       Exit Sub
    End If
    
    'REVISAMOS EL PESO
    If Not IsNumeric(MskPesEnt.Text) Then
       MsgBox "Peso Incorrecto", vbOKOnly + vbCritical, "Error"
       MskPesEnt.SetFocus
       Exit Sub
    End If
    
    If CboCal.Text <> "A" And CboCal.Text <> "R" And CboCal.Text <> "I" Then
        MsgBox "Calidad Incorrecta", vbOKOnly + vbInformation, "Informacion"
        CboCal.SetFocus
        Exit Sub
    End If
    
        
    'ASIGNA LA BARRA
    TxtBarra.Text = Format(MskFecPro.Text, "ddmmyy") & TxtLin.Text & TxtCodPro.Text & TxtTar.Text
        
    
    'SELECCIONA TODAS LAS TARIMAS DE PRODUCCION DE ACUERDO AL BATCH Y DE DONDE PROVIENEN LAS TARIMAS
    'QUE PUEDE SER POR PRODUCCION INTERNA, PRODUCCION LIBERADA O PRODUCCION EXTERNA
        Set RBuscaProduccion = New ADODB.Recordset
        If BProduccionInterna = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProduccion, "Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From Produccion Where Batch = " & VBatch & " And Linea = '" & VLinea & "' And (Calidad = 'A' or Calidad = 'R' Or Calidad = 'I')")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaProduccion, "Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From Produccion Where Batch = " & VBatch & " And UPPER(Linea) = '" & UCase(VLinea) & "' And (UPPER(Calidad) = 'A' or UPPER(Calidad) = 'R' Or UPPER(Calidad) = 'I')")
                End If
                If RBuscaProduccion.RecordCount > 0 Then
                Else
                    MsgBox "Tarima/Bulto No Existe En Produccion Interna", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
                End If
        ElseIf BProduccionLiberada = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProduccion, "Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From ProduccionLiberada Where Batch = " & VBatch & " And Linea = '" & VLinea & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaProduccion, "Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From ProduccionLiberada Where Batch = " & VBatch & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                End If
                If RBuscaProduccion.RecordCount > 0 Then
                Else
                    MsgBox "Tarima/Bulto No Existe En Produccion Liberada", vbOKOnly + vbInformation, "Informacion"
                    Exit Sub
                End If
        End If
    
                        'BUSCAMOS SI EXISTE LA TARIMA EN INVENTARIO
                        Set RBuscaTarima = New ADODB.Recordset
                            If GOrigenDeDatos = "AmaproAccess" Then
                                Call Abrir_Recordset(RBuscaTarima, "Select FichaTecnica From DetalleEntradasInventario Where FichaTecnica = '" & TxtCodPro.Text & "' And Tarima = " & TxtTar.Text & " And FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Linea = '" & TxtLin.Text & "'")
                            Else 'ORACLE
                                Call Abrir_Recordset(RBuscaTarima, "Select FichaTecnica From DetalleEntradasInventario Where UPPER(FichaTecnica) = '" & UCase(TxtCodPro.Text) & "' And Tarima = " & TxtTar.Text & " And FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                            End If
                        
                            'SI ENCUENTRA LA TARIMA LA EDITA
                            If RBuscaTarima.RecordCount > 0 Then
                                MsgBox "Tarima/Bulto " & TxtTar.Text & " Ya Fue Ingresada", vbOKOnly + vbInformation, "Revise Por Favor"
                            'AGREGA AL DETALLE DE LA ENTRADA DE PRODUCTO LO QUE SE CAPTURO EN PRODUCCION
                            Else
                                    
                                    VTexto = TxtDocDet.Text & ", " 'DOCUMENTO
                                    If GOrigenDeDatos = "AmaproAccess" Then
                                        VTexto = VTexto & "#" & Format(MskFecPro.Text, "mm/dd/yyyy") & "#, '" 'FECHA
                                    Else 'ORACLE
                                        VTexto = VTexto & "To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                                    End If
                                    VTexto = VTexto & TxtLin.Text & "', '" 'LINEA
                                    VTexto = VTexto & TxtCodPro.Text & "', " 'FICHA TECNICA
                                    VTexto = VTexto & TxtTar.Text & ", " 'TARIMA
                                    VTexto = VTexto & TxtBat.Text & ", '" 'BATCH
                                    VTexto = VTexto & CboCal.Text & "', '" 'CALIDAD
                                    VTexto = VTexto & TxtBod & "', '" 'BODEGA
                                    VTexto = VTexto & "" & "', '" 'PASILLO
                                    VTexto = VTexto & "" & "', '" 'CASILLA
                                    VTexto = VTexto & "" & "', " 'BIN
                                    VTexto = VTexto & TxtCanPro.Text & ", '" 'SALDO
                                    VTexto = VTexto & TxtOrd.Text & "', '" 'ORDEN
                                    VTexto = VTexto & TxtBarra.Text & "', " 'BARRA
                                    VTexto = VTexto & MskPesEnt.Text & ", '" 'PESO
                                    VTexto = VTexto & "NO INSPECCIONADO" & "', '" 'ESTADO
                                    VTexto = VTexto & TxtSerBol.Text & "', '" 'SERIE
                                    VTexto = VTexto & TxtOrdBol.Text & "', '" 'ORDEN
                                    VTexto = VTexto & TxtBulBol.Text & "', '" 'BULTO
                                    VTexto = VTexto & MskFecBol.Text & "', '" 'FECHA
                                    VTexto = VTexto & TxtBobBol.Text & "', " 'BOBINA
                                    VTexto = VTexto & TxtCanPro.Text & ", '" 'CANTIDAD DE ENTRADA
                                    VTexto = VTexto & TxtObs2.Text & "'" 'OBSERVACIONES
                                    
                                    Conexion.Execute "Insert Into DetalleEntradasInventario Values(" & VTexto & ")"
                            End If
                                        
                                     'SI SE DUPLICA LA LLAVE
                                     If GOrigenDeDatos = "AmaproAccess" Then
                                      'SI ES CUALQUIER OTRO ERROR
                                        If Err <> 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    Else 'ORACLE
                                        If Err = -2147217873 Then
                                            MsgBox "Tarima/Bulto Ya Existe", vbOKOnly + vbInformation, "Informacion"
                                            Exit Sub
                                      'SI ES CUALQUIER OTRO ERROR
                                        ElseIf Err <> -2147217873 And Err <> 0 Then
                                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                                            Exit Sub
                                        End If
                                    End If
        
                                    'CUENTA LAS TARIMAS QUE HAN SALIDO
                                    Set RCuentaTarimas = New ADODB.Recordset
                                        Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From DetalleEntradasInventario Where Documento = " & TxtDocIng.Text)
                                        If RCuentaTarimas.RecordCount > 0 Then
                                            TxtCueTar.Text = RCuentaTarimas(0)
                                        Else
                                            TxtCueTar.Text = 0
                                        End If

        
                        ' MUESTRA LAS ORDENES DE COMPRA DEL PRODUCTO CAPTURADO
                        Afecta_OC_MP
                        
                        Bandera2 = False
                        Botones2
                        RDetalle.Requery
                        RDetalle.MoveLast
                        Llena_CamposDetalle
                        DBGridDetalle.Enabled = True
                        CmdAgregar2.SetFocus
                        LblDes.Caption = ""
    
End Sub


Private Sub CmdAgregar_Click()
On Error Resume Next
    
    
    Bandera = True
    Botones1
    BEditarEncabezado = False
    FrameDetalle.Visible = False
    DBGridDetalle.Visible = False
    Limpia_CamposEncabezado
    'ASIGNA EL USUARIO
    TxtReq.Text = GUsuario
    'ASIGNA LA FECHA ACTUAL
    MskFec.Text = Format(Date, "dd/mm/yyyy")
    MskFec.SetFocus
    TxtEstado.Text = "NO LIBERADA"
    
    
    
    'ASIGNA VALOR AL CHECK DE PRODUCCION INTERNA
    ChkProInt.Value = 1
    TxtCueTar.Text = ""
    
    

End Sub


Private Sub CmdGrabar_Click()
On Error Resume Next

MousePointer = 11

    
    
    
    'ASIGNA VALORES A LAS VARIABLES PARA PODER CONTROLAR DE DONDE VIENEN LAS TARIMAS
    BProduccionInterna = ChkProInt.Value
    BProduccionLiberada = ChkProLib.Value
    
    
    'REVISA LA FECHA
    If Not IsDate(MskFec.Text) Then
        MsgBox "Fecha Incorrecta", vbOKOnly + vbInformation, "Informacion"
        MskFec.SetFocus
        Exit Sub
    End If
            
    
    'REVISA EL BATCH
    If Not IsNumeric(TxtBatch.Text) Then
        MsgBox "Numero De Batch Debe Ser Numerico", vbOKOnly + vbInformation, "Informacion"
        TxtBatch.SetFocus
        Exit Sub
    End If
    
    If TxtLinea.Text = "" Then
        MsgBox "Linea No Puede Estar En Blanco", vbOKOnly + vbInformation, "Informacion"
        TxtLinea.SetFocus
        Exit Sub
    End If
    
    
    VDocumento = TxtDocIng.Text
    VBodega = TxtBodega.Text
    VBatch = TxtBatch.Text
    VLinea = TxtLinea.Text
    MskFec.Text = Format(MskFec.Text, "dd/mm/yyyy")
               
    'GRABA DATOS
    'AGREGAR
                    If BEditarEncabezado = False Then
                    
                            'BUSCA EL DOCUMENTO MAXIMO Y LE AGREGA UNO MAS
                            Set RMaximo = New ADODB.Recordset
                            Call Abrir_Recordset(RMaximo, "Select max(Documento) from EncabezadoEntradasInventario")
                                If RMaximo.RecordCount > 0 Then
                                    If IsNull(RMaximo(0)) Then
                                        VDocumento = "1"
                                    Else
                                        VDocumento = Val(RMaximo(0)) + 1
                                    End If
                                End If
    
    
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = "#" & Format(MskFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = "To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & VDocumento & ", '" 'DOCUMENTO
                            VTexto = VTexto & TxtBodega.Text & "', " 'BODEGA
                            VTexto = VTexto & TxtBatch.Text & ", '" 'BATCH
                            VTexto = VTexto & TxtLinea.Text & "', '" 'LINEA
                            VTexto = VTexto & TxtObs.Text & "', " 'OBSERVACIONES
                            If ChkProInt.Value = "1" Then
                                VTexto = VTexto & "-1" & ", " 'PRODUCCION INTERNA
                            Else
                                VTexto = VTexto & "0" & ", " 'PRODUCCION INTERNA
                            End If
                            If ChkProLib.Value = "1" Then
                                VTexto = VTexto & "-1" & ", '" 'PRODUCCIO LIBERADA
                            Else
                                VTexto = VTexto & "0" & ", '" 'PRODUCCION LIBERADA
                            End If
                            VTexto = VTexto & TxtTexto.Item(0).Text & "', '" 'TIPO ENTRADA
                            VTexto = VTexto & TxtTexto.Item(1).Text & "', '" 'PROVEEDOR
                            VTexto = VTexto & TxtTexto.Item(3).Text & "', '" 'NUMERO DOCUMENTO
                            VTexto = VTexto & TxtTexto.Item(2).Text & "', '" 'TIPO DE DOCUMENTO
                            VTexto = VTexto & TxtTexto.Item(4).Text & "', '" 'TRANSPORTISTA
                            VTexto = VTexto & TxtTexto.Item(5).Text & "', '" 'PILOTO
                            VTexto = VTexto & TxtTexto.Item(6).Text & "', '" 'PLACAS CAMION
                            VTexto = VTexto & TxtTexto.Item(7).Text & "', '" 'PLACAS FURGON
                            VTexto = VTexto & GUsuario & "', '" 'REQUERIDO
                            VTexto = VTexto & "', '" 'LIBERADO
                            VTexto = VTexto & TxtEstado.Text & "'" 'ESTADO LIBERADO NO LIBERADO
                            
                            Conexion.Execute "Insert Into EncabezadoEntradasInventario Values(" & VTexto & ")"
                    'EDITAR
                    Else
                            If GOrigenDeDatos = "AmaproAccess" Then
                                VTexto = "FechaEntrada = #" & Format(MskFec.Text, "mm/dd/yyyy") & "#, " 'FECHA
                            Else 'ORACLE
                                VTexto = "FechaEntrada = To_Date('" & MskFec.Text & "', 'dd/mm/yyyy')" & ", " 'FECHA
                            End If
                            VTexto = VTexto & "Documento = " & TxtDocIng.Text & ", " 'DOCUMENTO
                            VTexto = VTexto & "Bodega = '" & TxtBodega.Text & "', " 'BODEGA
                            VTexto = VTexto & "Batch = " & TxtBatch.Text & ", " 'BATCH
                            VTexto = VTexto & "Linea = '" & TxtLinea.Text & "', " 'LINEA
                            VTexto = VTexto & "Observaciones = '" & TxtObs.Text & "', " 'OBSERVACIONES
                            If ChkProInt.Value = "1" Then
                                VTexto = VTexto & "ProduccionInterna = -1" & ", " 'PRODUCCION INTERNA
                            Else
                                VTexto = VTexto & "ProduccionInterna = 0" & ", " 'PRODUCCION INTERNA
                            End If
                            If ChkProLib.Value = "1" Then
                                VTexto = VTexto & "ProduccionLiberada = -1" & ", " 'PRODUCCIO LIBERADA
                            Else
                                VTexto = VTexto & "ProduccionLiberada = 0" & ", " 'PRODUCCION LIBERADA
                            End If
                            VTexto = VTexto & "TipoEntrada = '" & TxtTexto.Item(0).Text & "', " 'TIPO ENTRADA
                            VTexto = VTexto & "Proveedor = '" & TxtTexto.Item(1).Text & "', " 'PROVEEDOR
                            VTexto = VTexto & "NumeroDocumento = '" & TxtTexto.Item(3).Text & "', " 'NUMERO DOCUMENTO
                            VTexto = VTexto & "TipoDeDocumento = '" & TxtTexto.Item(2).Text & "', " 'TIPO DE DOCUMENTO
                            VTexto = VTexto & "Transportista = '" & TxtTexto.Item(4).Text & "', " 'TRANSPORTISTA
                            VTexto = VTexto & "NombreDePiloto = '" & TxtTexto.Item(5).Text & "', " 'PILOTO
                            VTexto = VTexto & "PlacasCamion = '" & TxtTexto.Item(6).Text & "', " 'PLACAS CAMION
                            VTexto = VTexto & "PlacasFurgon = '" & TxtTexto.Item(7).Text & "', " 'PLACAS FURGON
                            VTexto = VTexto & "Requerido = '" & TxtReq.Text & "'" 'REQUERIDO
                            
                            VTexto = VTexto & " Where Documento = " & VDocumento & " " 'DOCUMENTO
                            
                            Conexion.Execute "UPDATE EncabezadoEntradasInventario SET " & VTexto
                    End If
                    
                    'SI SE DUPLICA LA LLAVE
                     If GOrigenDeDatos = "AmaproAccess" Then
                        
                      'SI ES CUALQUIER OTRO ERROR
                        If Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    Else 'ORACLE
                        If Err = -2147217873 Then
                            MsgBox "Transaccion Ya Existe", vbOKOnly + vbInformation, "Informacion"
                            TxtDocIng.SetFocus
                            Exit Sub
                      'SI ES CUALQUIER OTRO ERROR
                        ElseIf Err <> -2147217873 And Err <> 0 Then
                            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbInformation, "Error"
                            Exit Sub
                        End If
                    End If
                    
    'CAMBIA BOTONES
    Bandera = False
    Botones1
    
    
    'SELECCIONA TODAS LAS TARIMAS DE PRODUCCION DE ACUERDO AL BATCH Y DE DONDE PROVIENEN LAS TARIMAS
    'QUE PUEDE SER POR PRODUCCION INTERNA, PRODUCCION LIBERADA O PRODUCCION EXTERNA
        Set RBuscaProduccion = New ADODB.Recordset
        If BProduccionInterna = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProduccion, "Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From Produccion Where Batch = " & VBatch & " And Linea = '" & VLinea & "' And (Calidad = 'A' or Calidad = 'R' Or Calidad = 'I')")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaProduccion, "Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From Produccion Where Batch = " & VBatch & " And UPPER(Linea) = '" & UCase(VLinea) & "' And (UPPER(Calidad) = 'A' or UPPER(Calidad) = 'R' Or UPPER(Calidad) = 'I')")
                End If
        ElseIf BProduccionLiberada = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProduccion, "Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From ProduccionLiberada Where Batch = " & VBatch & " And Linea = '" & VLinea & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaProduccion, "Select Fec_Prd, Tarima, Linea, Esp_Tec, Batch, Envases, Calidad, Orden, Barra From ProduccionLiberada Where Batch = " & VBatch & " And UPPER(Linea) = '" & UCase(VLinea) & "'")
                End If
        End If
                
            'SI ES PRODUCCION, O PRODUCCION LIBERADA
            If BProduccionInterna = True Or BProduccionLiberada = True Then
                
                         If RBuscaProduccion.RecordCount > 0 Then
                                 
                                 'CREA UN CICLO CON LOS DATOS DE PRODUCCION DEL BATCH
                                 Do Until RBuscaProduccion.EOF
                                     
                                     'BUSCAMOS SI EXISTE LA TARIMA
                                     Set RBuscaTarima = New ADODB.Recordset
                                         If GOrigenDeDatos = "AmaproAccess" Then
                                             Call Abrir_Recordset(RBuscaTarima, "Select FichaTecnica From DetalleEntradasInventario Where FichaTecnica = '" & RBuscaProduccion!Esp_Tec & "' And Tarima = " & RBuscaProduccion!Tarima & " And FechaProduccion = #" & Format(RBuscaProduccion!fec_prd, "mm/dd/yyyy") & "# And Linea = '" & RBuscaProduccion!Linea & "'")
                                         Else 'ORACLE
                                             Call Abrir_Recordset(RBuscaTarima, "Select FichaTecnica From DetalleEntradasInventario Where UPPER(FichaTecnica) = '" & UCase(RBuscaProduccion!Esp_Tec) & "' And Tarima = " & RBuscaProduccion!Tarima & " And FechaProduccion = To_Date('" & RBuscaProduccion!fec_prd & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(RBuscaProduccion!Linea) & "'")
                                         End If
                                     
                                         'SI ENCUENTRA LA TARIMA LA EDITA
                                         If RBuscaTarima.RecordCount > 0 Then
                                             MsgBox "Tarima " & RBuscaProduccion!Tarima & " Ya Fue Ingresada", vbOKOnly + vbInformation, "Revise Por Favor"
                                         'AGREGA AL DETALLE DE LA ENTRADA DE PRODUCTO LO QUE SE CAPTURO EN PRODUCCION
                                         Else
                                                 
                                                 VTexto = VDocumento & ", " 'DOCUMENTO
                                                 If GOrigenDeDatos = "AmaproAccess" Then
                                                     VTexto = VTexto & "#" & Format(RBuscaProduccion!fec_prd, "mm/dd/yyyy") & "#, '" 'FECHA
                                                 Else 'ORACLE
                                                     VTexto = VTexto & "To_Date('" & RBuscaProduccion!fec_prd & "', 'dd/mm/yyyy')" & ", '" 'FECHA
                                                 End If
                                                 VTexto = VTexto & RBuscaProduccion!Linea & "', '" 'LINEA
                                                 VTexto = VTexto & RBuscaProduccion!Esp_Tec & "', " 'FICHA TECNICA
                                                 VTexto = VTexto & RBuscaProduccion!Tarima & ", " 'TARIMA
                                                 VTexto = VTexto & VBatch & ", '" 'BATCH
                                                 VTexto = VTexto & RBuscaProduccion!Calidad & "', '" 'CALIDAD
                                                 VTexto = VTexto & VBodega & "', '" 'BODEGA
                                                 VTexto = VTexto & "" & "', '" 'PASILLO
                                                 VTexto = VTexto & "" & "', '" 'CASILLA
                                                 VTexto = VTexto & "" & "', " 'BIN
                                                 VTexto = VTexto & RBuscaProduccion!Envases & ", '" 'SALDO
                                                 VTexto = VTexto & RBuscaProduccion!Orden & "', '" 'ORDEN
                                                 VTexto = VTexto & RBuscaProduccion!Barra & "', " 'BARRA
                                                 VTexto = VTexto & "0" & ", '" 'PESO
                                                 
                                                 Set RBuscaTipoInventario = New ADODB.Recordset
                                                 Call Abrir_Recordset(RBuscaTipoInventario, "Select TipoInventario From FichaTecnica Where Esp_Tec = '" & RBuscaProduccion!Esp_Tec & "'")
                                                    If RBuscaTipoInventario.RecordCount > 0 Then
                                                        If RBuscaTipoInventario!TipoInventario = "PRODUCTO TERMINADO" Then
                                                            VTexto = VTexto & "INSPECCIONADO" & "', '" 'ESTADO
                                                        Else
                                                            VTexto = VTexto & "NO INSPECCIONADO" & "', '" 'ESTADO
                                                        End If
                                                    Else
                                                        VTexto = VTexto & "NO INSPECCIONADO" & "', '" 'ESTADO
                                                    End If
                                                 VTexto = VTexto & "" & "', '" 'SERIE
                                                 VTexto = VTexto & "" & "', '" 'ORDEN
                                                 VTexto = VTexto & "" & "', '" 'BULTO
                                                 VTexto = VTexto & "" & "', '" 'FECHA
                                                 VTexto = VTexto & "" & "', " 'BOBINA
                                                 VTexto = VTexto & RBuscaProduccion!Envases & ", " 'CANTIDAD DE ENTRADA
                                                 VTexto = VTexto & "''" 'CANTIDAD DE ENTRADA
                                                 
                                                 Conexion.Execute "Insert Into DetalleEntradasInventario Values(" & VTexto & ")"
                                         End If
                                                 
                                                 If Err <> 0 Then
                                                     MsgBox Err.Number & Err.Description & "Ojo Tarima " & RBuscaProduccion!Tarima & " Ya Existe, No Se Grabara, Pero Revise las Tarimas Por Favor", vbOKOnly + vbInformation, "Informacion"
                                                 End If
                                     'SE MUEVE AL SIGUIENTE REGISTRO
                                     RBuscaProduccion.MoveNext
                                 Loop
                         Else
                                                          
                         End If
            
            End If 'TERMINA LA PRODUCCION O PRODUCCION LIBERADA
            
                REncabezado.Requery
                REncabezado.MoveFirst
                REncabezado.Find ("Documento = " & VDocumento)

                Llena_CamposEncabezado

                                 Set RDetalle = New ADODB.Recordset
                                              If GOrigenDeDatos = "AmaproAccess" Then
                                                  Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & VDocumento & " And E.Documento = D.Documento And D.FichaTecnica = F.Esp_Tec")
                                              Else 'ORACLE
                                                  Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & VDocumento & " And E.Documento = D.Documento And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
                                              End If
                                                  Llena_CamposDetalle
                                                  Set DBGridDetalle.DataSource = RDetalle
    
            
            
            
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Visible = True
    FrameDetalle.Enabled = True
    FrameEncabezado.Enabled = False
    'VISUALIZA EL GRID DE DETALEE
    DBGridDetalle.Visible = True
    
    'HABILITA LAS COLUMNAS PARA PODER MODIFICARLAS PERO SOLO LA UBICACION
    DBGridDetalle.AllowUpdate = True
    DBGridDetalle.AllowDelete = True
        
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = True
    BotonesVisiblesDetalle
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = False
    BotonesVisiblesEncabezado
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    'BOTONES DE DATA
    CmdBotones2.Item(1).Visible = False
    CmdBotones2.Item(2).Visible = False
    CmdBotones2.Item(3).Visible = False
    CmdBotones2.Item(4).Visible = False
    
    TabEntradas.Tab = 1
       
    CmdAgregar2.SetFocus
    
MousePointer = 0
    
End Sub



Private Sub CmdImprimir_Click()
On Error Resume Next
        
        MousePointer = 11
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "InventarioEntradas.rpt"
                Else
                    GNombreReporte = "InventarioEntradasO.rpt"
                End If
                GCriteriaReporte = "{EncabezadoEntradasInventario.Documento} = " & TxtDocIng
                FrmReporte.Show
            
        MousePointer = 0

                

End Sub

Private Sub CmdImprimirAmarillas_Click()
On Error Resume Next
        MousePointer = 11
                    If TxtEstado.Text = "NO LIBERADA" Then
                        MsgBox "No se puede imprimir cedula porque no esta liberada", vbOKOnly + vbInformation, "Informacion"
                        Exit Sub
                    End If

                
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "BoletaAmarilla.rpt"
                Else
                    GNombreReporte = "BoletaAmarillaO.rpt"
                End If
                GCriteriaReporte = "{EncabezadoEntradasInventario.Documento} = " & TxtDocIng
                FrmReporte.Show
            
        MousePointer = 0
        
End Sub

Private Sub CmdSalida_Click()
    Unload Me
End Sub


Private Sub CmdTerminar_Click()
On Error Resume Next

'REVISA SI EL USUARIO ELEGIÓ UNA OC MP
If BGrabo = True Or BBorro = True Then
    Pregunta_OC_MP
    Exit Sub
Else
    
End If

'CONTINUA EL PROCEDIMIENTO DE TERMINAR
If CmdCancelar2.Enabled = True Then
     CmdCancelar2_Click
End If
    
    'VISUALIZA LOS BOTONES DEL DETALLE
    Bandera3 = False
    BotonesVisiblesDetalle
        
    
    'HABILITA EL DETALLE Y DESABILITA EL ENCABEZADO
    FrameDetalle.Enabled = False
    'FrameDetalle.Visible = False
    FrameEncabezado.Enabled = True
    
    'DESHABILITA LAS COLUMNAS PARA PODER MODIFICARLAS PERO SOLO LA UBICACION
    DBGridDetalle.AllowUpdate = True
    DBGridDetalle.AllowDelete = True
    
    CmdBotones2.Item(1).Visible = True
    CmdBotones2.Item(2).Visible = True
    CmdBotones2.Item(3).Visible = True
    CmdBotones2.Item(4).Visible = True
    
    'VISUALIZA LOS BOTONES DEL ENCABEZADO
    Bandera4 = True
    BotonesVisiblesEncabezado
        
    If Err <> 0 Then
        MsgBox Err.Description
    End If
    
    BanderaOCompraMP = False
    TabEntradas.Tab = 0
  
End Sub

Private Sub Command1_Click()
    FrameBuscar.Visible = False
End Sub



Private Sub CmdCedulas_Click()
On Error Resume Next
        If TxtEstado.Text = "NO LIBERADA" Then
            MsgBox "No se puede imprimir cedula porque no esta liberada", vbOKOnly + vbInformation, "Informacion"
            Exit Sub
        End If
        
        MousePointer = 11
                'MUESTRA EL REPORTE
                If GOrigenDeDatos = "AmaproAccess" Then
                    GNombreReporte = "CedulaMateriaPrima.rpt"
                Else
                    GNombreReporte = "CedulaMateriaPrimaO.rpt"
                End If
                GCriteriaReporte = "{EncabezadoEntradasInventario.Documento} = " & TxtDocIng
                FrmReporte.Show
            
        MousePointer = 0
        

End Sub

'Private Sub DataDetalleEntradas_Reposition()
'        If IsNumeric(TxtDocDet.Text) Then
'            'CUENTA CUANTAS TARIMAS TIENE EL DOCUMENTO
'            Set RCuentaTarimas = Db.OpenRecordset("Select Count(*) From DetalleEntradasInventario Where Documento = " & TxtDocDet.Text)
'                If RCuentaTarimas.RecordCount > 0 Then
'                    If IsNull(RCuentaTarimas(0)) Then
'                        TxtCueTar.Text = "0 Tarimas"
'                    Else
'                        TxtCueTar.Text = RCuentaTarimas(0) & " Tarimas"
'                    End If
'                Else
'                    TxtCueTar.Text = "0 Tarimas"
'                End If
'        End If
'End Sub
Private Sub DBGridBusqueda_DblClick()
    'BODEGA
    If BBodega = True Then
        TxtBodega.Text = DbGridBusqueda.Columns(0)
        TxtBodega.SetFocus
    'BODEGA DETALLE
    ElseIf BBodegaDetalle = True Then
        TxtBod.Text = DbGridBusqueda.Columns(0)
        TxtBod.SetFocus
    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        TxtCodPro.Text = DbGridBusqueda.Columns(0)
        TxtCodPro.SetFocus
    'LINEAS
    ElseIf BLineas = True Then
        TxtLin.Text = DbGridBusqueda.Columns(0)
        TxtLin.SetFocus
    'TIPO ENTRADAS
    ElseIf BTipoEntrada = True Then
        TxtTexto.Item(0).Text = DbGridBusqueda.Columns(0)
        TxtTexto.Item(0).SetFocus
    'PROVEEDOR
    ElseIf BProveedor = True Then
        TxtTexto.Item(1).Text = DbGridBusqueda.Columns(0)
        TxtTexto.Item(1).SetFocus
    'TIPO DE DOCUMENTO
    ElseIf BTipoDocumento = True Then
        TxtTexto.Item(2).Text = DbGridBusqueda.Columns(0)
        TxtTexto.Item(2).SetFocus
    'TARNSPORTISTA
    ElseIf BTransportista = True Then
        TxtTexto.Item(4).Text = DbGridBusqueda.Columns(0)
        TxtTexto.Item(4).SetFocus
    End If
        TxtBuscar.Text = ""
        FrameBuscar.Visible = False
End Sub

Private Sub DbGridBusqueda_HeadClick(ByVal ColIndex As Integer)
        RBusqueda.Sort = RBusqueda.Fields(ColIndex).Name
End Sub

Private Sub DBGridBusqueda_KeyPress(KeyAscii As Integer)
            If KeyAscii = 43 Then
                'BODEGA
                If BBodega = True Then
                    TxtBodega.Text = DbGridBusqueda.Columns(0)
                    TxtBodega.SetFocus
                'BODEGA DETALLE
                ElseIf BBodegaDetalle = True Then
                    TxtBod.Text = DbGridBusqueda.Columns(0)
                    TxtBod.SetFocus
                'PRODUCTO TERMINADO
                ElseIf BProducto = True Then
                    TxtCodPro.Text = DbGridBusqueda.Columns(0)
                    TxtCodPro.SetFocus
                'LINEAS
                ElseIf BLineas = True Then
                    TxtLin.Text = DbGridBusqueda.Columns(0)
                    TxtLin.SetFocus
                'TIPO ENTRADAS
                ElseIf BTipoEntrada = True Then
                    TxtTexto.Item(0).Text = DbGridBusqueda.Columns(0)
                    TxtTexto.Item(0).SetFocus
                'PROVEEDOR
                ElseIf BProveedor = True Then
                    TxtTexto.Item(1).Text = DbGridBusqueda.Columns(0)
                    TxtTexto.Item(1).SetFocus
                'TIPO DE DOCUMENTO
                ElseIf BTipoDocumento = True Then
                    TxtTexto.Item(2).Text = DbGridBusqueda.Columns(0)
                    TxtTexto.Item(2).SetFocus
                'TARNSPORTISTA
                ElseIf BTransportista = True Then
                    TxtTexto.Item(4).Text = DbGridBusqueda.Columns(0)
                    TxtTexto.Item(4).SetFocus
                End If
                    TxtBuscar.Text = ""
                    FrameBuscar.Visible = False
            End If

End Sub

Private Sub DbGridDetalle_BeforeUpdate(Cancel As Integer)
        If Err <> 0 Then
            MsgBox "Error " & Err.Number & " " & Err.Description, vbOKOnly + vbCritical, "Error"
        End If
    
End Sub


Private Sub DbGridDetalle_HeadClick(ByVal ColIndex As Integer)
        RDetalle.Sort = RDetalle.Fields(ColIndex).Name
End Sub


Private Sub DbGridDetalle_SelChange(Cancel As Integer)
        Llena_CamposDetalle
        
        'HABILITA EL DBGRID PARA OBTENER OC MP
        Afecta_OC_MP
        
End Sub

Private Sub dg_OCompra_MP_Click()
On Error Resume Next
    
    ' SE INDICA QUE EL USUARIO ELIGIÓ UNA OC DE MP
    BanderaOCompraMP = True
    
    MsgBox "Eligió con un click una OC MP ", vbOKOnly + vbInformation, "Informacion"
    
    ' SE ENVIA AL PROCESO QUE GUARDA EL CAMBIO EN LA TABLA DE "DETALLE PEDIDOS PROVEEDORES"
    ' CON LA CANTIDAD Y EL PRODUCTO
    
    ' TERMINA EL PROCESO DE ACTUALZIACION EN LA TABLA
    
    ' SE REGRESA EL CONTROL AL FORMULARIO PRIMARIO DE CAPTURA DE ENTRADA DE INVENTARIO
    'SE QUITAN LAS BANDERAS PARA PODER CERRAR LA VENTANA
    
    
    ' VER STATUS DE QUE HIZO EL USUARIO
    BGrabo = False
    BEdito = False
    BBorro = False
    BanderaOCompraMP = False

End Sub

Private Sub dg_OCompra_MP_KeyPress(KeyAscii As Integer)
On Error Resume Next
    
    ' SE INDICA QUE EL USUARIO ELIGIÓ UNA OC DE MP
    BanderaOCompraMP = True
    
    MsgBox "Eligió con un click una OC MP ", vbOKOnly + vbInformation, "Informacion"
    
    ' SE ENVIA AL PROCESO QUE GUARDA EL CAMBIO EN LA TABLA DE "DETALLE PEDIDOS PROVEEDORES"
    ' CON LA CANTIDAD Y EL PRODUCTO
    
    ' TERMINA EL PROCESO DE ACTUALZIACION EN LA TABLA
    
    ' SE REGRESA EL CONTROL AL FORMULARIO PRIMARIO DE CAPTURA DE ENTRADA DE INVENTARIO


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then
        CmdTerminar_Click
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next

    TabEntradas.Tab = 0

    BanderaOCompraMP = False
    ' VER STATUS DE QUE HIZO EL USUARIO
    BGrabo = False
    BEdito = False
    BBorro = False

Set REncabezado = New ADODB.Recordset
    If GOrigenDeDatos = "AmaproAccess" Then
        Call Abrir_Recordset(REncabezado, "Select * From EncabezadoEntradasInventario Where FechaEntrada >= #" & Format((Date - 730), "mm/dd/yyyy") & "# And FechaEntrada <= #" & Format(Date, "mm/dd/yyyy") & "# Order By Documento")
    Else 'ORACLE
        Call Abrir_Recordset(REncabezado, "Select * From EncabezadoEntradasInventario Where FechaEntrada >= To_Date('" & (Date - 730) & "', 'dd/mm/yyyy') And FechaEntrada <= To_Date('" & Date & "', 'dd/mm/yyyy') Order By Documento")
    End If
    'SE VA AL ULTIMO REGISTRO
    REncabezado.MoveLast
    
    Llena_CamposEncabezado
            
    Set RDetalle = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & TxtDocIng & " And E.Documento = D.Documento And D.FichaTecnica = F.Esp_Tec")
            Else 'ORACLE
                Call Abrir_Recordset(RDetalle, "Select D.Documento, D.FechaProduccion, D.Linea, D.FichaTecnica, F.Descrip, D.Tarima, D.Batch, D.Calidad, D.Bodega, D.CantidadEntrada, D.Saldo, D.OrdenProduccion, D.PesoEntrada, D.Estado, D.Barra, D.SerieBoleta, D.OrdenBoleta, D.BultoBoleta, D.FechaBoleta, D.BobinaBoleta, D.Observaciones From EncabezadoEntradasInventario E, DetalleEntradasInventario D, FichaTecnica F Where E.Documento = " & TxtDocIng & " And E.Documento = D.Documento And UPPER(D.FichaTecnica) = UPPER(F.Esp_Tec)")
            End If
            Llena_CamposDetalle
            Set DBGridDetalle.DataSource = RDetalle
            
            If Err <> 0 Then
                MsgBox Err.Description
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


Private Sub MskFecBol_GotFocus()
        MskFecBol.SelStart = 0
        MskFecBol.SelLength = Len(MskFecBol.Text)
End Sub

Private Sub MskFecBol_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub MskFecPro_GotFocus()
        MskFecPro.SelStart = 0
        MskFecPro.SelLength = Len(MskFecPro.Text)
End Sub
Private Sub MskFecPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub MskPesEnt_GotFocus()
        MskPesEnt.SelStart = 0
        MskPesEnt.SelLength = Len(MskPesEnt.Text)
End Sub

Private Sub MskPesEnt_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub


Private Sub TabEntradas_Click(PreviousTab As Integer)
                        If TabEntradas.Tab = 1 Then
                                    'CUENTA LAS TARIMAS QUE HAN SALIDO
                                    Set RCuentaTarimas = New ADODB.Recordset
                                        Call Abrir_Recordset(RCuentaTarimas, "Select Count(*) From DetalleEntradasInventario Where Documento = " & TxtDocIng.Text)
                                        If RCuentaTarimas.RecordCount > 0 Then
                                            TxtCueTar.Text = RCuentaTarimas(0)
                                        Else
                                            TxtCueTar.Text = 0
                                        End If
                        End If
End Sub

Private Sub TxtBarra_GotFocus()
        TxtBarra.SelStart = 0
        TxtBarra.SelLength = Len(TxtBarra.Text)
End Sub

Private Sub TxtBarra_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub Txtbat_GotFocus()
        TxtBat.SelStart = 0
        TxtBat.SelLength = Len(TxtBat.Text)
End Sub

Private Sub Txtbat_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtBatch_GotFocus()
        TxtBatch.SelStart = 0
        TxtBatch.SelLength = Len(TxtBatch.Text)
End Sub

Private Sub TxtBatch_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtBobBol_GotFocus()
        TxtBobBol.SelStart = 0
        TxtBobBol.SelLength = Len(TxtBobBol.Text)
End Sub

Private Sub TxtBobBol_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtBod_Change()
        Set RBuscaBodega = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBod.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBod.Text) & "'")
            End If
            If RBuscaBodega.RecordCount > 0 Then
                LblBod.Caption = RBuscaBodega!Descripcion
            Else
                LblBod.Caption = ""
            End If
End Sub

Private Sub TxtBod_DblClick()
            TxtBuscar.Visible = True
            OptDescripcion.Visible = True
            OptCodigo.Visible = True
            
            BBodega = False
            BProducto = False
            BBodegaDetalle = True
            BLineas = False
            
            BTipoEntrada = False
            BProveedor = False
            BTipoDocumento = False
            BTransportista = False
            
            FrameBuscar.Visible = True
            TxtBuscar.SetFocus
            Set RBusqueda = New ADODB.Recordset
            Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Order by CodigoBodega")
            Set DbGridBusqueda.DataSource = RBusqueda
            DbGridBusqueda.Columns(1).Width = "4000"

End Sub

Private Sub TxtBod_GotFocus()
            TxtBod.SelStart = 0
            TxtBod.SelLength = Len(TxtBod.Text)
End Sub

Private Sub TxtBod_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
               SendKeys "{tab}"
            End If
            
            If KeyAscii = 43 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = True
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = False
                BTipoDocumento = False
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Order by CodigoBodega")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            End If
End Sub
Private Sub TxtBodega_Change()
            Set RBuscaBodega = New ADODB.Recordset
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where CodigoBodega = '" & TxtBodega.Text & "'")
                Else 'ORACLE
                    Call Abrir_Recordset(RBuscaBodega, "Select Descripcion From BodegasInventario Where UPPER(CodigoBodega) = '" & UCase(TxtBodega.Text) & "'")
                End If
                If RBuscaBodega.RecordCount > 0 Then
                    LblBodega.Caption = RBuscaBodega!Descripcion
                Else
                    LblBodega.Caption = ""
                End If
End Sub
Private Sub TxtBodega_DblClick()
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                BBodega = True
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = False
                BTipoDocumento = False
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Order by CodigoBodega")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"

End Sub
Private Sub TxtBodega_GotFocus()
                TxtBodega.SelStart = 0
                TxtBodega.SelLength = Len(TxtBodega.Text)
End Sub

Private Sub TxtBodega_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
               SendKeys "{tab}"
            End If
                
            If KeyAscii = 43 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                BBodega = True
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = False
                BTipoDocumento = False
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Order by CodigoBodega")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            End If
End Sub


Private Sub TxtBulBol_GotFocus()
            TxtBulBol.SelStart = 0
            TxtBulBol.SelLength = Len(TxtBulBol.Text)
End Sub

Private Sub TxtBulBol_KeyPress(KeyAscii As Integer)
            If KeyAscii = 13 Then
               SendKeys "{tab}"
            End If
End Sub

Private Sub Txtbuscar_Change()
    Set RBusqueda = New ADODB.Recordset
    'BODEGA
    If (BBodega = True Or BBodegaDetalle = True) Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Where CodigoBodega Like '%" & TxtBuscar.Text & "%' Order by CodigoBodega")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Where UPPER(CodigoBodega) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoBodega")
                End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by CodigoBodega")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select * from BodegasInventario Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoBodega")
                End If
        End If
    'PRODUCTO TERMINADO
    ElseIf BProducto = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Esp_Tec Like '%" & TxtBuscar.Text & "%' And Activa = -1 Order by Esp_Tec")
                Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where UPPER(Esp_Tec) Like '%" & UCase(TxtBuscar.Text) & "%' And Activa = -1 Order by Esp_Tec")
                End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
                If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Descrip Like '%" & TxtBuscar.Text & "%' And Activa = -1 Order by Esp_Tec")
                Else
                    Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' And Activa = -1 Order by Esp_Tec")
                End If
        End If
    'LINEAS
    ElseIf BLineas = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select * from Lineas Where Linea Like '%" & TxtBuscar.Text & "%' Order by Linea")
            Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select * from Lineas Where UPPER(Linea) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Linea")
            End If
            
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select * from Lineas Where Descrip Like '%" & TxtBuscar.Text & "%' Order by Linea")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select * from Lineas Where UPPER(Descrip) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Linea")
            End If
        End If
    'TIPO DE ENTRADA
    ElseIf BTipoEntrada = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from TiposEntradasInventario Where Codigo Like '%" & TxtBuscar.Text & "%' Order by Codigo")
            Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from TiposEntradasInventario Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Codigo")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from TiposEntradasInventario Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by Codigo")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from TiposEntradasInventario Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Codigo")
            End If
        End If
    'PROVEEDOR
    ElseIf BProveedor = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where CodigoProveedor Like '%" & TxtBuscar.Text & "%' Order by CodigoProveedor")
            Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(CodigoProveedor) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoProveedor")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by Codigoproveedor")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoProveedor")
            End If
        End If
    'TIPO DE DOCUMENTO
    ElseIf BTipoDocumento = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where CodigoDocumento Like '%" & TxtBuscar.Text & "%' Order by CodigoDocumento")
            Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where UPPER(CodigoDocumento) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Codigodocumento")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by CodigoDocumento")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by CodigoDocumento")
            End If
        End If
    'TIPO DE DOCUMENTO
    ElseIf BTransportista = True Then
        'SI VA A BUSCAR POR CODIGO
        If OptCodigo.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas Where Codigo Like '%" & TxtBuscar.Text & "%' Order by Codigo")
            Else 'ORACLE
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas Where UPPER(Codigo) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Codigo")
            End If
        'SI VA A BUSCAR POR DESCRIPCION
        ElseIf OptDescripcion.Value = True Then
            If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas Where Descripcion Like '%" & TxtBuscar.Text & "%' Order by Codigo")
            Else
                    Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas Where UPPER(Descripcion) Like '%" & UCase(TxtBuscar.Text) & "%' Order by Codigo")
            End If
        End If

    End If
        
        Set DbGridBusqueda.DataSource = RBusqueda
        DbGridBusqueda.Columns(1).Width = "4000"
End Sub
Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Private Sub CboCal_GotFocus()
        CboCal.SelStart = 0
        CboCal.SelLength = Len(CboCal.Text)
End Sub

Private Sub CboCal_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
        
End Sub

Private Sub TxtCanPro_GotFocus()
        TxtCanPro.SelStart = 0
        TxtCanPro.SelLength = Len(TxtCanPro.Text)
End Sub
Private Sub TxtCanPro_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub TxtCanPro_LostFocus()
            'SI DESEA MULTIPLICAR LAS LAMINAS QUE VIENEN, BUSCA POR EL CODIGO CUANTAS LAMINA TIENE CADA CODIGO Y LAS MULTIPLICA
            If ChkMultiplica.Value = 1 Then
                If IsNumeric(TxtCanPro.Text) Then
                    Set RBuscaCuerpos = New ADODB.Recordset
                        Call Abrir_Recordset(RBuscaCuerpos, "Select UnidadesxLamina From FichaTecnica Where Esp_Tec = '" & TxtCodPro.Text & "'")
                        If RBuscaCuerpos.RecordCount > 0 Then
                            TxtCanPro.Text = TxtCanPro.Text * RBuscaCuerpos!UnidadesxLamina
                        End If
                End If
            End If

End Sub

Private Sub TxtCodPro_Change()
                 Set RBuscaProducto = New ADODB.Recordset
                 If GOrigenDeDatos = "AmaproAccess" Then
                    Call Abrir_Recordset(RBuscaProducto, "Select Descrip From FichaTecnica Where Esp_Tec = '" & TxtCodPro.Text & "'")
                 Else 'ORACLE
                    Call Abrir_Recordset(RBuscaProducto, "Select Descrip From FichaTecnica Where UPPER(Esp_Tec) = '" & UCase(TxtCodPro.Text) & "'")
                 End If
                 If RBuscaProducto.RecordCount > 0 Then
                        LblDes.Caption = RBuscaProducto(0)
                 Else
                        LblDes.Caption = ""
                 End If
End Sub

Private Sub TxtCodPro_DblClick()
                Set RBusqueda = New ADODB.Recordset
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                BBodega = False
                BProducto = True
                BBodegaDetalle = False
                BLineas = False
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1 Order by Esp_Tec")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
End Sub

Private Sub TxtCodPro_GotFocus()
                TxtCodPro.SelStart = 0
                TxtCodPro.SelLength = Len(TxtCodPro.Text)
End Sub

Private Sub TxtCodPro_KeyPress(KeyAscii As Integer)
                'SI PRECIONA ENTER
                If KeyAscii = 13 Then
                   SendKeys "{tab}"
                End If
                'SI PRECIONA LA TECLA DE SIGNO +
                If KeyAscii = 43 Then
                   Set RBusqueda = New ADODB.Recordset
                   TxtBuscar.Visible = True
                   OptDescripcion.Visible = True
                   OptCodigo.Visible = True
                   
                   BBodega = False
                   BProducto = True
                   BBodegaDetalle = False
                   BLineas = False
                   FrameBuscar.Visible = True
                   TxtBuscar.SetFocus
                   Call Abrir_Recordset(RBusqueda, "Select Esp_Tec, Descrip, MaterialEmpaque, Size from FichaTecnica Where Activa = -1 Order by Esp_Tec")
                   Set DbGridBusqueda.DataSource = RBusqueda
                   DbGridBusqueda.Columns(1).Width = "4000"
                End If
End Sub

Private Sub TxtCodPro_LostFocus()
            Set RBuscaUltimaTarima = New ADODB.Recordset
                        If GOrigenDeDatos = "AmaproAccess" Then
                            'Call Abrir_Recordset(RBuscaUltimaTarima, "Select Max(Tarima) From DetalleEntradasInventario Where FichaTecnica = '" & TxtCodPro.Text & "' And FechaProduccion = #" & Format(MskFecPro.Text, "mm/dd/yyyy") & "# And Linea = '" & TxtLin.Text & "'")
                            Call Abrir_Recordset(RBuscaUltimaTarima, "Select Max(Tarima) From DetalleEntradasInventario Where FichaTecnica = '" & TxtCodPro.Text & "' And Linea = '" & TxtLin.Text & "'")
                        Else 'ORACLE
                            'Call Abrir_Recordset(RBuscaUltimaTarima, "Select Max(Tarima) From DetalleEntradasInventario Where UPPER(FichaTecnica) = '" & UCase(TxtCodPro.Text) & "' And FechaProduccion = To_Date('" & MskFecPro.Text & "', 'dd/mm/yyyy')" & " And UPPER(Linea) = '" & UCase(TxtLin.Text) & "'")
                            Call Abrir_Recordset(RBuscaUltimaTarima, "Select Max(Tarima) From DetalleEntradasInventario Where UPPER(FichaTecnica) = '" & UCase(TxtCodPro.Text) & "' And Linea = '" & TxtLin.Text & "'")
                        End If
                 
                 If RBuscaUltimaTarima.RecordCount > 0 Then
                        If IsNull(RBuscaUltimaTarima(0)) Then
                            TxtTarUlt.Text = "0"
                        Else
                            TxtTarUlt.Text = RBuscaUltimaTarima(0)
                        End If
                 Else
                        TxtTarUlt.Text = "0"
                 End If

End Sub

Private Sub TxtDocing_KeyPress(KeyAscii As Integer)
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
                LblLin2.Caption = RBuscaLinea!Descrip
            Else
                LblLin2.Caption = ""
            End If

End Sub

Private Sub Txtlin_DblClick()
        Set RBusqueda = New ADODB.Recordset
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        
        BBodega = False
        BProducto = False
        BBodegaDetalle = False
        BLineas = True
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        Call Abrir_Recordset(RBusqueda, "Select * from Lineas")
        Set DbGridBusqueda.DataSource = RBusqueda
        DbGridBusqueda.Columns(1).Width = "4000"
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
        Set RBusqueda = New ADODB.Recordset
        TxtBuscar.Visible = True
        OptDescripcion.Visible = True
        OptCodigo.Visible = True
        
        BBodega = False
        BProducto = False
        BBodegaDetalle = False
        BLineas = True
        FrameBuscar.Visible = True
        TxtBuscar.SetFocus
        Call Abrir_Recordset(RBusqueda, "Select * from Lineas")
        Set DbGridBusqueda.DataSource = RBusqueda
        DbGridBusqueda.Columns(1).Width = "4000"
    End If

End Sub


Private Sub TxtLinea_Change()
        Set RBuscaLinea = New ADODB.Recordset
            If GOrigenDeDatos = "AmaproAccess" Then
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where Linea = '" & TxtLinea.Text & "'")
            Else 'ORACLE
                Call Abrir_Recordset(RBuscaLinea, "Select Descrip From Lineas Where UPPER(Linea) = '" & UCase(TxtLinea.Text) & "'")
            End If
            If RBuscaLinea.RecordCount > 0 Then
                LblLinea.Caption = RBuscaLinea!Descrip
            Else
                LblLinea.Caption = ""
            End If
        
End Sub

Private Sub TxtLinea_GotFocus()
        TxtLinea.SelStart = 0
        TxtLinea.SelLength = Len(TxtLinea.Text)
End Sub

Private Sub TxtLinea_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtObs_GotFocus()
        TxtObs.SelStart = 0
        TxtObs.SelLength = Len(TxtObs.Text)
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtObs2_GotFocus()
        TxtObs2.SelStart = 0
        TxtObs2.SelLength = Len(TxtObs2.Text)
End Sub

Private Sub TxtObs2_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
            SendKeys "{tab}"
        End If
End Sub

Private Sub TxtOrd_GotFocus()
        TxtOrd.SelStart = 0
        TxtOrd.SelLength = Len(TxtOrd.Text)
End Sub

Private Sub TxtOrd_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
       SendKeys "{tab}"
    End If
End Sub

Private Sub TxtOrd_LostFocus()
        'ORDEN EN DETALLE DE PRODUCCION
                Set RBuscaFichaOrden = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaFichaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where Documento = '" & TxtOrd.Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaFichaOrden, "Select FichaTecnica From EncabezadoOrdenProduccion Where UPPER(Documento) = '" & UCase(TxtOrd.Text) & "'")
                    End If
                    If RBuscaFichaOrden.RecordCount > 0 Then
                        TxtCodPro.Text = RBuscaFichaOrden!FichaTecnica
                    Else
                        
                    End If

End Sub

Private Sub TxtOrdBol_GotFocus()
        TxtOrdBol.SelStart = 0
        TxtOrdBol.SelLength = Len(TxtOrdBol.Text)
End Sub

Private Sub TxtOrdBol_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub


Private Sub TxtSerBol_GotFocus()
        TxtSerBol.SelStart = 0
        TxtSerBol.SelLength = Len(TxtSerBol.Text)
End Sub

Private Sub TxtSerBol_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
End Sub

Private Sub Txttar_GotFocus()
        TxtTar.SelStart = 0
        TxtTar.SelLength = Len(TxtTar.Text)
End Sub

Private Sub Txttar_KeyPress(KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If

End Sub

Public Sub BotonesVisiblesEncabezado()
    If Bandera4 = True Then
         CmdAgregar.Visible = True
         CmdEditar.Visible = True
         CmdGrabar.Visible = True
         CmdCancelar.Visible = True
         CmdBorrar.Visible = True
         CmdBuscar.Visible = True
         CmdImprimir.Visible = True
         CmdImprimirAmarillas.Visible = True
         CmdSalida.Visible = True
    Else
         CmdAgregar.Visible = False
         CmdEditar.Visible = False
         CmdGrabar.Visible = False
         CmdCancelar.Visible = False
         CmdBorrar.Visible = False
         CmdBuscar.Visible = False
         CmdImprimir.Visible = False
         CmdImprimirAmarillas.Visible = False
         CmdSalida.Visible = False
    End If

End Sub

Public Sub Llena_CamposEncabezado()
On Error Resume Next
            
                            If Not IsNull(REncabezado!FechaEntrada) Then
                                MskFec.Text = REncabezado!FechaEntrada
                            Else
                                MskFec.Text = ""
                            End If
                            If Not IsNull(REncabezado!Documento) Then
                                TxtDocIng.Text = REncabezado!Documento
                            Else
                                TxtDocIng.Text = ""
                            End If
                            If Not IsNull(REncabezado!Bodega) Then
                                TxtBodega.Text = REncabezado!Bodega
                            Else
                                TxtBodega.Text = ""
                            End If
                            If Not IsNull(REncabezado!Batch) Then
                                TxtBatch.Text = REncabezado!Batch
                            Else
                                TxtBatch.Text = ""
                            End If
                            If Not IsNull(REncabezado!Linea) Then
                                TxtLinea.Text = REncabezado!Linea
                            Else
                                TxtLinea.Text = ""
                            End If
                            If Not IsNull(REncabezado!Observaciones) Then
                                TxtObs.Text = REncabezado!Observaciones
                            Else
                                TxtObs.Text = ""
                            End If
                            If REncabezado!ProduccionInterna = "-1" Then
                                ChkProInt.Value = "1"
                            Else
                                ChkProInt.Value = "0"
                            End If
                            If REncabezado!ProduccionLiberada = "-1" Then
                                ChkProLib.Value = "1"
                            Else
                                ChkProLib.Value = "0"
                            End If
                            If Not IsNull(REncabezado!TipoEntrada) Then
                                TxtTexto.Item(0).Text = REncabezado!TipoEntrada
                            Else
                                TxtTexto.Item(0).Text = ""
                            End If
                            If Not IsNull(REncabezado!Proveedor) Then
                                TxtTexto.Item(1).Text = REncabezado!Proveedor
                            Else
                                TxtTexto.Item(1).Text = ""
                            End If
                            If Not IsNull(REncabezado!NumeroDocumento) Then
                                TxtTexto.Item(3).Text = REncabezado!NumeroDocumento
                            Else
                                TxtTexto.Item(3).Text = ""
                            End If
                            If Not IsNull(REncabezado!TipoDeDocumento) Then
                                TxtTexto.Item(2).Text = REncabezado!TipoDeDocumento
                            Else
                                TxtTexto.Item(2).Text = ""
                            End If
                            If Not IsNull(REncabezado!Transportista) Then
                                TxtTexto.Item(4).Text = REncabezado!Transportista
                            Else
                                TxtTexto.Item(4).Text = ""
                            End If
                            If Not IsNull(REncabezado!NombreDePiloto) Then
                                TxtTexto.Item(5).Text = REncabezado!NombreDePiloto
                            Else
                                TxtTexto.Item(5).Text = ""
                            End If
                            If Not IsNull(REncabezado!PlacasCamion) Then
                                TxtTexto.Item(6).Text = REncabezado!PlacasCamion
                            Else
                                TxtTexto.Item(6).Text = ""
                            End If
                            If Not IsNull(REncabezado!PlacasFurgon) Then
                                TxtTexto.Item(7).Text = REncabezado!PlacasFurgon
                            Else
                                TxtTexto.Item(7).Text = ""
                            End If
                            If Not IsNull(REncabezado!Requerido) Then
                                TxtReq.Text = REncabezado!Requerido
                            Else
                                TxtReq.Text = ""
                            End If
                            If Not IsNull(REncabezado!Liberado) Then
                                TxtLib.Text = REncabezado!Liberado
                            Else
                                TxtLib.Text = ""
                            End If
                            If Not IsNull(REncabezado!Estado) Then
                                TxtEstado.Text = REncabezado!Estado
                            Else
                                TxtEstado.Text = ""
                            End If
            
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Llena_CamposDetalle()
On Error Resume Next
                If RDetalle.RecordCount > 0 Then
                                If Not IsNull(RDetalle!Documento) Then
                                    TxtDocDet.Text = RDetalle!Documento
                                Else
                                    TxtDocDet.Text = ""
                                End If
                                If Not IsNull(RDetalle!FechaProduccion) Then
                                    MskFecPro = RDetalle!FechaProduccion
                                Else
                                    MskFecPro = ""
                                End If
                                If Not IsNull(RDetalle!Linea) Then
                                    TxtLin.Text = RDetalle!Linea
                                Else
                                    TxtLin.Text = ""
                                End If
                                If Not IsNull(RDetalle!FichaTecnica) Then
                                    TxtCodPro.Text = RDetalle!FichaTecnica
                                Else
                                    TxtCodPro.Text = ""
                                End If
                                If Not IsNull(RDetalle!Tarima) Then
                                    TxtTar.Text = RDetalle!Tarima
                                Else
                                    TxtTar.Text = ""
                                End If
                                If Not IsNull(RDetalle!Tarima) Then
                                    TxtBat.Text = RDetalle!Batch
                                Else
                                    TxtBat.Text = ""
                                End If
                                If Not IsNull(RDetalle!Calidad) Then
                                    CboCal.Text = RDetalle!Calidad
                                Else
                                    CboCal.Text = ""
                                End If
                                If Not IsNull(RDetalle!Bodega) Then
                                    TxtBod.Text = RDetalle!Bodega
                                Else
                                    TxtBod.Text = ""
                                End If
                                If Not IsNull(RDetalle!OrdenProduccion) Then
                                    TxtOrd.Text = RDetalle!OrdenProduccion
                                Else
                                    TxtOrd.Text = ""
                                End If
                                If Not IsNull(RDetalle!Barra) Then
                                    TxtBarra.Text = RDetalle!Barra
                                Else
                                    TxtBarra.Text = ""
                                End If
                                If Not IsNull(RDetalle!PesoEntrada) Then
                                    MskPesEnt.Text = RDetalle!PesoEntrada
                                Else
                                    MskPesEnt.Text = 0
                                End If
                                If Not IsNull(RDetalle!Estado) Then
                                    TxtEst.Text = RDetalle!Estado
                                Else
                                    TxtEst.Text = ""
                                End If
                                If Not IsNull(RDetalle!SerieBoleta) Then
                                    TxtSerBol.Text = RDetalle!SerieBoleta
                                Else
                                    TxtSerBol.Text = ""
                                End If
                                If Not IsNull(RDetalle!OrdenBoleta) Then
                                    TxtOrdBol.Text = RDetalle!OrdenBoleta
                                Else
                                    TxtOrdBol.Text = ""
                                End If
                                If Not IsNull(RDetalle!BultoBoleta) Then
                                    TxtBulBol.Text = RDetalle!BultoBoleta
                                Else
                                    TxtBulBol.Text = ""
                                End If
                                If Not IsNull(RDetalle!FechaBoleta) Then
                                    MskFecBol.Text = RDetalle!FechaBoleta
                                Else
                                    MskFecBol.Text = ""
                                End If
                                If Not IsNull(RDetalle!BobinaBoleta) Then
                                    TxtBobBol.Text = RDetalle!BobinaBoleta
                                Else
                                    TxtBobBol.Text = ""
                                End If
                                If Not IsNull(RDetalle!CantidadEntrada) Then
                                    TxtCanPro.Text = RDetalle!CantidadEntrada
                                Else
                                    TxtCanPro.Text = ""
                                End If
                                If Not IsNull(RDetalle!Observaciones) Then
                                    TxtObs2.Text = RDetalle!Observaciones
                                Else
                                    TxtObs2.Text = ""
                                End If
                    Else
                                    TxtDocDet.Text = "0"
                                    MskFecPro = ""
                                    TxtLin.Text = ""
                                    TxtCodPro.Text = ""
                                    TxtTar.Text = "0"
                                    TxtBat.Text = "0"
                                    CboCal.Text = ""
                                    TxtBod.Text = ""
                                    TxtOrd.Text = ""
                                    TxtBarra.Text = ""
                                    MskPesEnt.Text = "0"
                                    TxtEst.Text = ""
                                    TxtSerBol.Text = ""
                                    TxtOrdBol.Text = ""
                                    TxtBulBol.Text = ""
                                    MskFecBol.Text = ""
                                    TxtBobBol.Text = ""
                                    TxtCanPro.Text = "0"
                                    TxtObs2.Text = ""
                    End If
                    
            'HABILITA EL DBGRID PARA OBTENER OC MP
            Afecta_OC_MP
            
            If Err <> 0 Then
            End If
End Sub

Public Sub Limpia_CamposEncabezado()
                            MskFec.Text = ""
                            TxtDocIng.Text = "0"
                            TxtBod.Text = ""
                            TxtBatch.Text = 0
                            TxtLinea.Text = ""
                            TxtObs.Text = ""
                            ChkProInt.Value = "0"
                            ChkProLib.Value = "0"
                            TxtTexto.Item(0).Text = ""
                            TxtTexto.Item(1).Text = ""
                            TxtTexto.Item(3).Text = ""
                            TxtTexto.Item(2).Text = ""
                            TxtTexto.Item(4).Text = ""
                            TxtTexto.Item(5).Text = ""
                            TxtTexto.Item(6).Text = ""
                            TxtTexto.Item(7).Text = ""
                            TxtReq.Text = ""
                            TxtLib.Text = ""
                            TxtEstado.Text = ""
End Sub

Public Sub Limpia_CamposDetalle()
                                    TxtDocDet.Text = "0"
                                    MskFecPro = ""
                                    TxtLin.Text = ""
                                    TxtCodPro.Text = ""
                                    TxtTar.Text = "0"
                                    TxtBat.Text = "0"
                                    CboCal.Text = ""
                                    TxtBod.Text = ""
                                    TxtOrd.Text = ""
                                    TxtBarra.Text = ""
                                    MskPesEnt.Text = "0"
                                    TxtEst.Text = ""
                                    TxtSerBol.Text = ""
                                    TxtOrdBol.Text = ""
                                    TxtBulBol.Text = ""
                                    MskFecBol.Text = ""
                                    TxtBobBol.Text = ""
                                    TxtCanPro.Text = "0"
                                    TxtObs2.Text = ""
End Sub



Private Sub TxtTexto_Change(Index As Integer)
            If Index = 0 Then
                Set RBuscaTipoEntrada = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaTipoEntrada, "select Descripcion From TiposEntradasInventario Where Codigo = '" & TxtTexto.Item(0).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaTipoEntrada, "select Descripcion From TiposEntradasInventario Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(0).Text) & "'")
                    End If
                        If RBuscaTipoEntrada.RecordCount > 0 Then
                            LblTipEnt.Caption = RBuscaTipoEntrada!Descripcion
                        Else
                            LblTipEnt.Caption = ""
                        End If
                            
            ElseIf Index = 1 Then
                Set RBuscaProveedor = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaProveedor, "select Descripcion From Proveedores Where CodigoProveedor = '" & TxtTexto.Item(1).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaProveedor, "select Descripcion From Proveedores Where UPPER(CodigoProveedor) = '" & UCase(TxtTexto.Item(1).Text) & "'")
                    End If
                        If RBuscaProveedor.RecordCount > 0 Then
                            LblPro.Caption = RBuscaProveedor!Descripcion
                        Else
                            LblPro.Caption = ""
                        End If
            ElseIf Index = 2 Then
                Set RBuscaTipoDocumento = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaTipoDocumento, "select Descripcion From Documentos Where CodigoDocumento = '" & TxtTexto.Item(2).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaTipoDocumento, "select Descripcion From Documentos Where UPPER(CodigoDocumento) = '" & UCase(TxtTexto.Item(2).Text) & "'")
                    End If
                        If RBuscaTipoDocumento.RecordCount > 0 Then
                            LblTipDoc.Caption = RBuscaTipoDocumento!Descripcion
                        Else
                            LblTipDoc.Caption = ""
                        End If
            ElseIf Index = 4 Then
                Set RBuscaTransportista = New ADODB.Recordset
                    If GOrigenDeDatos = "AmaproAccess" Then
                        Call Abrir_Recordset(RBuscaTransportista, "select Descripcion From Transportistas Where Codigo = '" & TxtTexto.Item(4).Text & "'")
                    Else 'ORACLE
                        Call Abrir_Recordset(RBuscaTransportista, "select Descripcion From Transportistas Where UPPER(Codigo) = '" & UCase(TxtTexto.Item(4).Text) & "'")
                    End If
                        If RBuscaTransportista.RecordCount > 0 Then
                            LblTra.Caption = RBuscaTransportista!Descripcion
                        Else
                            LblTra.Caption = ""
                        End If

            End If
End Sub

Private Sub Txttexto_DblClick(Index As Integer)
            
            'TIPO ENTRADA
            If Index = 0 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = True
                BProveedor = False
                BTipoDocumento = False
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from TiposEntradasInventario")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            'PROVEEDOR
            ElseIf Index = 1 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = True
                BTipoDocumento = False
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            'TIPO DE DOCUMENTO
            ElseIf Index = 2 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = False
                BTipoDocumento = True
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            'TRANSPORTISTA
            ElseIf Index = 4 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = False
                BTipoDocumento = False
                BTransportista = True
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            End If
End Sub

Private Sub TxtTexto_KeyPress(Index As Integer, KeyAscii As Integer)
        If KeyAscii = 13 Then
           SendKeys "{tab}"
        End If
        If KeyAscii = 43 Then
            'TIPO ENTRADA
            If Index = 0 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = True
                BProveedor = False
                BTipoDocumento = False
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from TiposEntradasInventario")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            'PROVEEDOR
            ElseIf Index = 1 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = True
                BTipoDocumento = False
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select CodigoProveedor, Descripcion from Proveedores")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            'TIPO DE DOCUMENTO
            ElseIf Index = 2 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = False
                BTipoDocumento = True
                BTransportista = False
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select CodigoDocumento, Descripcion from Documentos")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            'TRANSPORTISTA
            ElseIf Index = 4 Then
                TxtBuscar.Visible = True
                OptDescripcion.Visible = True
                OptCodigo.Visible = True
                
                
                BBodega = False
                BProducto = False
                BBodegaDetalle = False
                BLineas = False
                
                BTipoEntrada = False
                BProveedor = False
                BTipoDocumento = False
                BTransportista = True
                
                FrameBuscar.Visible = True
                TxtBuscar.SetFocus
                Set RBusqueda = New ADODB.Recordset
                Call Abrir_Recordset(RBusqueda, "Select Codigo, Descripcion from Transportistas")
                Set DbGridBusqueda.DataSource = RBusqueda
                DbGridBusqueda.Columns(1).Width = "4000"
            End If
        End If

End Sub

Public Sub Afecta_OC_MP()
On Error Resume Next
    
    'MsgBox "Debe Elegir una Orden de Compra de MP a Afectar ", vbOKOnly + vbInformation, "Informacion"
            
        'Frame_Compras_MP.Visible = True
                  
    ' ABRIMOS RECORDSET PARA SACAR LAS O.C. DE MP DE DICHO PRODUCTO CAPTURADO
    Set RBuscaOrdenCompraMP = New ADODB.Recordset
    If GOrigenDeDatos = "AmaproAccess" Then
        Call Abrir_Recordset(RBuscaOrdenCompraMP, "SELECT DISTINCT " & _
                            "E.Fecha, E.Documento, E.Proveedor, " & _
                            "P.Descripcion, D.Codigo, D.CantidadPedido, " & _
                            "D.SaldoPorEntregar " & _
                            "FROM " & _
                            "EncabezadoPedidosProveedores E, " & _
                            "DetallePedidosProveedores D, " & _
                            "Proveedores P " & _
                            "WHERE " & _
                            "E.Documento = D.Documento AND " & _
                            "D.SaldoPorEntregar > 0 AND " & _
                            "E.Proveedor = P.CodigoProveedor AND " & _
                            "D.Codigo LIKE '" & TxtCodPro.Text & "' ")
    Else 'ORACLE

    End If
        
        Set dg_OCompra_MP.DataSource = RBuscaOrdenCompraMP
 
    'OCompraMP.SetFocus
    'MsgBox "Pulse OK para continuar ", vbOKOnly + vbInformation, "Informacion"

End Sub

Public Sub Pregunta_OC_MP()
On Error Resume Next
    
    'REVISA SI EL USUARIO ELEGIÓ UNA OC MP
    If VLinea = "77" Then  'Es MP
        'EJECUTA PROCEDIMIENTO PARA AFECTAR LAS OC DE MP
        If BanderaOCompraMP = True Then
            
        Else
            MsgBox "Debe Elegir una Orden de Compra de MP a Afectar ", vbOKOnly + vbInformation, "Informacion"
            dg_OCompra_MP.SetFocus
            Exit Sub
        End If
        
   Else
    
   End If

End Sub
